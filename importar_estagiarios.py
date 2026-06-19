"""
Script de importação de estagiários da planilha lista_estagiarios_DB.xlsx
Uso: python importar_estagiarios.py "postgresql://..."
"""
import sys
import re
import pandas as pd
import psycopg2
from psycopg2.extras import execute_values

XLSX = r'C:\Users\usuario\Downloads\lista_estagiarios_DB.xlsx'
SHEET = 'cadastro estagiários'

def limpar_cpf(v):
    if not v or str(v).strip() in ('nan', ''):
        return None
    return str(v).strip()

def limpar_semestre(v):
    if not v or str(v).strip() in ('nan', ''):
        return None
    m = re.search(r'\d+', str(v))
    return int(m.group()) if m else None

def limpar_str(v):
    if not v or str(v).strip() in ('nan', ''):
        return None
    return str(v).strip()

def montar_endereco(end, bairro, cidade):
    partes = [p for p in [limpar_str(end), limpar_str(bairro), limpar_str(cidade)] if p]
    return ', '.join(partes) or None

def main():
    if len(sys.argv) < 2:
        print("Uso: python importar_estagiarios.py \"postgresql://user:pass@host/db\"")
        sys.exit(1)

    db_url = sys.argv[1]

    print(f"Lendo planilha: {XLSX}")
    df = pd.read_excel(XLSX, sheet_name=SHEET, dtype=str, skiprows=1, header=0)
    # Renomear colunas pela posição (planilha pode ter nomes variados)
    cols = list(df.columns)
    # Col 0=NOME, 1=SEMESTRE, 2=CURSO, 3=CPF, 4=EMAIL, 5=ENDEREÇO, 6=BAIRRO, 7=CIDADE
    df.columns = ['nome', 'semestre', 'curso', 'cpf', 'email',
                  'endereco', 'bairro', 'cidade'] + [f'_x{i}' for i in range(len(cols) - 8)]

    # Remover linha de header duplicado que aparece como dado
    df = df[df['nome'] != 'NOME'].copy()
    # Manter só registros com nome e CPF preenchidos
    df = df[df['nome'].notna() & df['cpf'].notna()].copy()

    registros = []
    seen_cpf = set()
    for _, row in df.iterrows():
        cpf = limpar_cpf(row.get('cpf'))
        nome = limpar_str(row.get('nome'))
        if not cpf or not nome:
            continue
        if cpf in seen_cpf:
            continue  # deduplicar dentro da planilha
        seen_cpf.add(cpf)
        registros.append({
            'nome': nome,
            'cpf': cpf,
            'email': limpar_str(row.get('email')),
            'endereco': montar_endereco(row.get('endereco'), row.get('bairro'), row.get('cidade')),
            'semestre': limpar_semestre(row.get('semestre')),
            'tipo_ensino': 'superior',
            'status': 'ativo',
        })

    print(f"Registros únicos a inserir: {len(registros)}")

    conn = psycopg2.connect(db_url)
    cur = conn.cursor()

    inseridos = 0
    ignorados = 0
    for r in registros:
        cur.execute("""
            INSERT INTO estagiario (nome, cpf, email, endereco, semestre, tipo_ensino, status)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (cpf) DO NOTHING
        """, (r['nome'], r['cpf'], r['email'], r['endereco'],
              r['semestre'], r['tipo_ensino'], r['status']))
        if cur.rowcount == 1:
            inseridos += 1
        else:
            ignorados += 1

    conn.commit()
    cur.close()
    conn.close()

    print(f"\nResultado:")
    print(f"  Inseridos: {inseridos}")
    print(f"  Ignorados (CPF já existia): {ignorados}")
    print(f"  Total processado: {inseridos + ignorados}")

if __name__ == '__main__':
    main()
