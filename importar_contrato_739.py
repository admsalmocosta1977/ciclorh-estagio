"""
Importa contrato 739 — Gustavo de Matos Pereira / Metaplac
"""
import sys, json
import psycopg2
from psycopg2.extras import RealDictCursor

def main():
    if len(sys.argv) < 2:
        print("Uso: python importar_contrato_739.py \"postgresql://...\"")
        sys.exit(1)

    db_url = sys.argv[1]
    conn = psycopg2.connect(db_url, cursor_factory=RealDictCursor)
    conn.autocommit = True
    cur = conn.cursor()

    # Buscar estagiário por CPF
    cur.execute("SELECT id, nome FROM estagiario WHERE REPLACE(REPLACE(REPLACE(cpf,'.',''),'-',''),' ','') = '08030970501'")
    est = cur.fetchone()
    if not est:
        print("ERRO: Estagiário não encontrado (CPF 080.309.705-01)")
        sys.exit(1)
    print(f"Estagiário: {est['nome']} (id={est['id']})")

    # Buscar empresa por CNPJ
    cur.execute("SELECT id, nome FROM empresa WHERE REPLACE(REPLACE(REPLACE(REPLACE(cnpj,'.',''),'/',''),'-',''),' ','') = '07851050000121'")
    emp = cur.fetchone()
    if not emp:
        print("ERRO: Empresa não encontrada (CNPJ 07.851.050/0001-21)")
        sys.exit(1)
    print(f"Empresa: {emp['nome']} (id={emp['id']})")

    # Buscar supervisor da Metaplac
    cur.execute("SELECT id, nome, cargo, registro FROM empresa_supervisor WHERE empresa_id = %s AND nome ILIKE '%%elizangela%%'", (emp['id'],))
    sup = cur.fetchone()
    sup_nome = sup['nome'] if sup else 'ELIZÂNGELA ROCHA DÓREA'
    sup_cargo = sup['cargo'] if sup else None
    sup_reg = sup['registro'] if sup else None
    print(f"Supervisor: {sup_nome}")

    # Buscar IE por sigla
    cur.execute("SELECT id, nome FROM ie WHERE sigla ILIKE 'UESB' OR nome ILIKE '%SUDOESTE DA BAHIA%'")
    ie = cur.fetchone()
    if not ie:
        print("ERRO: IE UESB não encontrada")
        sys.exit(1)
    print(f"IE: {ie['nome']} (id={ie['id']})")

    # Verificar se contrato já existe
    cur.execute("SELECT id FROM contrato WHERE numero_contrato = '739' AND estagiario_id = %s", (est['id'],))
    if cur.fetchone():
        print("AVISO: Contrato 739 já existe para este estagiário. Abortando.")
        sys.exit(0)

    # Atividades
    atividades = "||".join([
        "Auxiliar de escritório",
        "Auxiliar de contabilidade",
        "Auxiliar de emissão de notas fiscais",
        "Auxiliar de faturamento",
        "Auxiliar de operação de ERP",
        "Auxiliar de gestão de documentos gerenciais e fiscais",
        "Participar de reuniões quando solicitado",
    ])

    # Jornada: seg-sex 08:00-14:00 (matutino)
    jornada = {}
    for dia in ['seg', 'ter', 'qua', 'qui', 'sex']:
        jornada[dia] = {'mat_ini': '08:00', 'mat_fim': '14:00'}
    jornada_json = json.dumps(jornada, ensure_ascii=False)

    cur.execute("""
        INSERT INTO contrato (
            estagiario_id, empresa_id, ie_id, orientador,
            supervisor_nome, supervisor_cargo, supervisor_registro,
            curso, tipo_estagio, area_atuacao,
            ch_diaria, ch_semanal,
            data_inicio, data_fim, numero_contrato,
            bolsa, bolsa_tipo, taxa, aux_transporte,
            atividades, jornada
        ) VALUES (
            %s, %s, %s, %s,
            %s, %s, %s,
            %s, %s, %s,
            %s, %s,
            %s, %s, %s,
            %s, %s, %s, %s,
            %s, %s
        ) RETURNING id
    """, (
        est['id'], emp['id'], ie['id'], 'Salmo Lima Costa',
        sup_nome, sup_cargo, sup_reg,
        'Ciências Contábeis', 'Não Obrigatório', 'Ciências Contábeis',
        6, 30,
        '2025-12-10', '2026-12-09', '739',
        900.00, 'mensal', None, 100.00,
        atividades, jornada_json
    ))
    contrato_id = cur.fetchone()['id']
    print(f"\nContrato inserido com sucesso! ID={contrato_id}")
    print(f"  Número: 739")
    print(f"  Vigência: 10/12/2025 a 09/12/2026")
    print(f"  Bolsa: R$ 900,00 | Aux. Transporte: R$ 100,00")
    print(f"  Atividades: 7")

    cur.close()
    conn.close()

if __name__ == '__main__':
    main()
