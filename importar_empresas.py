"""
Importação de empresas da planilha empresas_DB.xlsx
"""
import sys
import re
import pandas as pd
import psycopg2

XLSX = r'C:\Users\usuario\Downloads\empresas_DB.xlsx'

def limpar(v):
    if v is None: return None
    s = str(v).strip()
    return None if s in ('nan', 'None', '') else s

def limpar_float(v):
    s = limpar(v)
    if not s: return None
    try: return float(s.replace(',', '.'))
    except: return None

def montar_endereco(end, bairro):
    e = limpar(end)
    b = limpar(bairro)
    if e and b:  return f'{e}, Bairro {b}'
    if e:        return e
    if b:        return f'Bairro {b}'
    return None

def separar_nome_registro(v):
    """
    'COREN 914937 - Fulano' → nome='Fulano', registro='COREN 914937'
    'Fulano da Silva - CRO 123' → nome='Fulano da Silva', registro='CRO 123'
    'Fulano da Silva' → nome='Fulano da Silva', registro=None
    """
    s = limpar(v)
    if not s: return None, None
    # Padrões de registro no final: COREN, CRO, CRM, CREA, CRP, OAB, CFF, CREFITO, CFBM + número
    m = re.search(r'-\s*((?:COREN|CRO|CRM|CREA|CRP|OAB|CFF|CREFITO|CFBM)\s*[\w\-/]+)\s*$', s, re.I)
    if m:
        reg = m.group(1).strip()
        nome = s[:m.start()].strip().rstrip('-').strip()
        return nome or None, reg
    # Padrão no início: 'COREN 914937 - Nome'
    m2 = re.match(r'^((?:COREN|CRO|CRM|CREA|CRP|OAB|CFF|CREFITO|CFBM)\s*[\w\-/]+)\s*-\s*(.+)', s, re.I)
    if m2:
        return m2.group(2).strip(), m2.group(1).strip()
    return s, None

def main():
    if len(sys.argv) < 2:
        print("Uso: python importar_empresas.py \"postgresql://...\"")
        sys.exit(1)
    db_url = sys.argv[1]

    print(f"Lendo planilha: {XLSX}")
    df = pd.read_excel(XLSX, dtype=str, header=None, skiprows=1)
    n_sup_cols = df.shape[1] - 16
    df.columns = ['codigo','nome_fantasia','razao_social','constituicao','natureza','cnpj',
                  'endereco','bairro','cidade','estado','email','representante','cargo',
                  'cpf_rep','bolsa','aux_transp'] + [f'sup_{i}' for i in range(n_sup_cols)]

    df = df[df['codigo'] != 'CÓDIGO'].copy()
    df = df[df['razao_social'].notna() & df['razao_social'].str.strip().ne('')].copy()
    df = df[df['razao_social'].str.strip() != 'nan'].copy()
    print(f"Total de registros: {len(df)}")

    conn = psycopg2.connect(db_url)
    conn.autocommit = True
    cur = conn.cursor()

    inseridas = 0
    ignoradas = 0
    sups_inseridos = 0
    cnpjs_vistos = set()

    sup_cols = [c for c in df.columns if c.startswith('sup_')]

    for _, row in df.iterrows():
        nome = limpar(row.get('razao_social'))
        if not nome: continue

        cnpj = limpar(row.get('cnpj'))
        if cnpj in cnpjs_vistos:
            print(f"  [DUP] {nome} — CNPJ {cnpj} já processado, ignorando.")
            ignoradas += 1
            continue
        if cnpj:
            cnpjs_vistos.add(cnpj)

        endereco = montar_endereco(row.get('endereco'), row.get('bairro'))
        cidade = limpar(row.get('cidade'))
        email = limpar(row.get('email'))
        representante = limpar(row.get('representante'))
        cargo_rep = limpar(row.get('cargo'))
        cpf_rep = limpar(row.get('cpf_rep'))
        bolsa = limpar_float(row.get('bolsa'))
        aux = limpar_float(row.get('aux_transp'))

        cur.execute("""
            INSERT INTO empresa
                (nome, cnpj, endereco, cidade, email,
                 representante, cargo_representante, cpf_representante,
                 bolsa_padrao, aux_transporte_padrao, status)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'ativo')
            RETURNING id
        """, (nome, cnpj, endereco, cidade or 'Vitória da Conquista',
              email, representante, cargo_rep, cpf_rep, bolsa, aux))

        emp_id = cur.fetchone()[0]
        inseridas += 1

        for i, sc in enumerate(sup_cols):
            sv = limpar(row.get(sc))
            if not sv or sv == 'nan': continue
            sup_nome, sup_reg = separar_nome_registro(sv)
            if not sup_nome: continue
            cur.execute("""
                INSERT INTO empresa_supervisor (empresa_id, nome, registro, ordem)
                VALUES (%s, %s, %s, %s)
            """, (emp_id, sup_nome, sup_reg, i))
            sups_inseridos += 1

    cur.close()
    conn.close()

    print(f"\nResultado:")
    print(f"  Empresas inseridas:   {inseridas}")
    print(f"  Ignoradas (dup CNPJ): {ignoradas}")
    print(f"  Supervisores:         {sups_inseridos}")

if __name__ == '__main__':
    main()
