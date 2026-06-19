"""
Importa IEs novas e atualiza dados faltantes nas IEs já cadastradas.
Match por SIGLA → CNPJ → NOME (nessa ordem de prioridade).
"""
import sys, re
import pandas as pd
import psycopg2
from psycopg2.extras import RealDictCursor

XLSX = r'C:\Users\usuario\Downloads\instituições_DB.xlsx'

def lim(v):
    if v is None: return None
    s = str(v).strip()
    return None if s in ('nan','None','','_______________________________') else s

def montar_endereco(end, bairro):
    e, b = lim(end), lim(bairro)
    if e and b:  return f'{e}, Bairro {b}'
    if e:        return e
    if b:        return f'Bairro {b}'
    return None

def limpar_representante(v):
    """Extrai só o nome antes de informações extras (CPF, Residente, portador...)"""
    s = lim(v)
    if not s: return None
    # Corta no primeiro indicador de informação adicional
    for sep in [', portador', ', CPF', ', residente', ', Residente', ' portador', ' CPF']:
        if sep.lower() in s.lower():
            idx = s.lower().index(sep.lower())
            s = s[:idx].strip().rstrip(',').strip()
            break
    return s or None

def parse_orientador(v):
    """'Nome da Silva (cargo)' → (nome, cargo)"""
    s = lim(v)
    if not s: return None, None
    m = re.match(r'^(.+?)\s*\(([^)]+)\)\s*$', s)
    if m:
        return m.group(1).strip().rstrip('.'), m.group(2).strip()
    return s.rstrip('.').strip(), None

def norm_sigla(v):
    s = lim(v)
    return s.upper().strip() if s else None

def norm_cnpj(v):
    s = lim(v)
    if not s: return None
    return re.sub(r'[\.\-\/\s]', '', s)

def main():
    if len(sys.argv) < 2:
        print("Uso: python importar_ies.py \"postgresql://...\"")
        sys.exit(1)

    db_url = sys.argv[1]
    print(f"Lendo planilha: {XLSX}")

    df = pd.read_excel(XLSX, dtype=str, header=None, skiprows=2)
    df.columns = ['tipo','x1','codigo','sigla','nome','cnpj','endereco','bairro','cidade','estado',
                  'representante','cargo'] + [f'orientador_{i}' for i in range(1, 11)]
    df = df[df['nome'].notna() & df['nome'].str.strip().ne('') & df['nome'].str.strip().ne('nan')].copy()
    df = df[df['codigo'].notna() & df['codigo'].str.strip().ne('nan')].copy()
    print(f"Total na planilha: {len(df)}")

    conn = psycopg2.connect(db_url, cursor_factory=RealDictCursor)
    conn.autocommit = True
    cur = conn.cursor()

    # Carregar IEs existentes do banco
    cur.execute("SELECT * FROM ie")
    db_ies = list(cur.fetchall())

    # Índices para match rápido
    idx_sigla = {norm_sigla(r['sigla']): dict(r) for r in db_ies if r.get('sigla')}
    idx_cnpj  = {norm_cnpj(r.get('cnpj','')): dict(r) for r in db_ies if r.get('cnpj')}
    idx_nome  = {str(r['nome']).strip().upper(): dict(r) for r in db_ies}

    inseridas = atualizadas = profs_inseridos = 0

    orient_cols = [f'orientador_{i}' for i in range(1, 11)]

    for _, row in df.iterrows():
        nome      = lim(row.get('nome'))
        sigla     = lim(row.get('sigla'))
        cnpj      = lim(row.get('cnpj'))
        endereco  = montar_endereco(row.get('endereco'), row.get('bairro'))
        cidade    = lim(row.get('cidade')) or 'Vitória da Conquista'
        rep       = limpar_representante(row.get('representante'))
        cargo_r   = lim(row.get('cargo'))
        if not nome: continue

        # Match
        db_rec = (idx_sigla.get(norm_sigla(sigla))
               or idx_cnpj.get(norm_cnpj(cnpj))
               or idx_nome.get(nome.upper()))

        orientadores = []
        for oc in orient_cols:
            n, c = parse_orientador(row.get(oc))
            if n:
                orientadores.append((n, c))

        if db_rec:
            # ── ATUALIZAR campos faltantes ──────────────────────────────
            ie_id = db_rec['id']
            updates, vals = [], []

            def add(field, val):
                if val and not db_rec.get(field):
                    updates.append(f"{field}=%s")
                    vals.append(val)

            add('sigla', sigla)
            add('cnpj', cnpj)
            add('endereco', endereco)
            add('cidade', cidade)
            add('representante_legal', rep)
            add('cargo_representante_legal', cargo_r)

            if updates:
                vals.append(ie_id)
                cur.execute(f"UPDATE ie SET {', '.join(updates)} WHERE id=%s", vals)
                atualizadas += 1

            # Professores: inserir apenas os que ainda não existem (por nome)
            cur.execute("SELECT nome FROM ie_professor WHERE ie_id=%s", (ie_id,))
            nomes_existentes = {r['nome'].strip().lower() for r in cur.fetchall()}
            ordem = len(nomes_existentes)
            for (pn, pc) in orientadores:
                if pn.strip().lower() not in nomes_existentes:
                    cur.execute(
                        "INSERT INTO ie_professor (ie_id,nome,cargo,ordem) VALUES (%s,%s,%s,%s)",
                        (ie_id, pn, pc, ordem))
                    nomes_existentes.add(pn.strip().lower())
                    profs_inseridos += 1
                    ordem += 1
        else:
            # ── INSERIR nova IE ─────────────────────────────────────────
            cur.execute("""
                INSERT INTO ie (nome,sigla,cnpj,endereco,cidade,representante_legal,
                                cargo_representante_legal,signatario_tce)
                VALUES (%s,%s,%s,%s,%s,%s,%s,'coordenador') RETURNING id
            """, (nome, sigla, cnpj, endereco, cidade, rep, cargo_r))
            ie_id = cur.fetchone()['id']
            inseridas += 1

            for ordem, (pn, pc) in enumerate(orientadores):
                cur.execute(
                    "INSERT INTO ie_professor (ie_id,nome,cargo,ordem) VALUES (%s,%s,%s,%s)",
                    (ie_id, pn, pc, ordem))
                profs_inseridos += 1

            # Atualizar índices para evitar duplicatas dentro da própria planilha
            if sigla: idx_sigla[norm_sigla(sigla)] = {'id': ie_id}
            if cnpj:  idx_cnpj[norm_cnpj(cnpj)]   = {'id': ie_id}
            idx_nome[nome.upper()] = {'id': ie_id}

    cur.close()
    conn.close()

    print(f"\nResultado:")
    print(f"  Novas IEs inseridas:       {inseridas}")
    print(f"  IEs existentes atualizadas: {atualizadas}")
    print(f"  Orientadores adicionados:  {profs_inseridos}")

if __name__ == '__main__':
    main()
