#!/usr/bin/env python3
"""
Extrator de empresas da Receita Federal — rodar LOCALMENTE (não no servidor).

Uso:
    python rfb_extrator_local.py
    python rfb_extrator_local.py "FEIRA DE SANTANA"

Saída: empresas_<municipio>.csv  (importe pelo sistema em Admin > Prospecção > Importar CSV)
"""

import csv
import io
import os
import sys
import tempfile
import unicodedata
import zipfile

import requests

RFB_BASE = 'https://dadosabertos.rfb.gov.br/CNPJ/'
HEADERS = {'User-Agent': 'Mozilla/5.0'}
MUNICIPIO = sys.argv[1].strip().upper() if len(sys.argv) > 1 else 'VITORIA DA CONQUISTA'


def norm(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s.upper())
                   if unicodedata.category(c) != 'Mn')


def baixar_zip(url, descricao):
    print(f'  Baixando {descricao}…', end=' ', flush=True)
    r = requests.get(url, headers=HEADERS, stream=True, timeout=300)
    r.raise_for_status()
    tmp = tempfile.NamedTemporaryFile(suffix='.zip', delete=False)
    total = 0
    for chunk in r.iter_content(131072):
        tmp.write(chunk)
        total += len(chunk)
    tmp.close()
    print(f'OK ({total // 1_048_576} MB)')
    return tmp.name


def ler_csv_zip(path):
    with zipfile.ZipFile(path) as z:
        nome = z.namelist()[0]
        with z.open(nome) as f:
            yield from csv.reader(io.TextIOWrapper(f, encoding='latin-1'), delimiter=';')


def main():
    print(f'\n=== Extrator RFB — {MUNICIPIO} ===\n')

    # 1. Código do município
    print('[1/4] Localizando município…')
    path = baixar_zip(RFB_BASE + 'Municipios.zip', 'Municipios.zip')
    mun_code = None
    for row in ler_csv_zip(path):
        if len(row) >= 2 and norm(row[1].strip()) == norm(MUNICIPIO):
            mun_code = row[0].strip()
            break
    os.unlink(path)
    if not mun_code:
        sys.exit(f'Município "{MUNICIPIO}" não encontrado. Tente sem acento, em maiúsculas.')
    print(f'    Código RFB: {mun_code}')

    # 2. CNAEs
    print('[2/4] Carregando CNAEs…')
    path = baixar_zip(RFB_BASE + 'Cnaes.zip', 'Cnaes.zip')
    cnae_desc = {}
    for row in ler_csv_zip(path):
        if len(row) >= 2:
            cnae_desc[row[0].strip()] = row[1].strip().title()
    os.unlink(path)
    print(f'    {len(cnae_desc)} CNAEs carregados')

    # 3. Estabelecimentos
    print('[3/4] Varrendo estabelecimentos (10 arquivos ~300 MB cada)…')
    estab = {}
    for i in range(10):
        url = f'{RFB_BASE}Estabelecimentos{i}.zip'
        path = None
        try:
            path = baixar_zip(url, f'Estabelecimentos{i}.zip')
            count = 0
            for row in ler_csv_zip(path):
                if len(row) < 22:
                    continue
                if row[5].strip() == '02' and row[20].strip() == mun_code:
                    cb = row[0].strip()
                    tel = (row[21].strip() + row[22].strip()).strip() or ''
                    estab[cb] = {
                        'cnpj': cb + row[1].strip() + row[2].strip(),
                        'nome_fantasia': row[4].strip().title(),
                        'cnae': row[11].strip(),
                        'endereco': ' '.join(x for x in [row[13].strip(), row[14].strip(), row[15].strip()] if x).title(),
                        'bairro': row[17].strip().title(),
                        'telefone': tel,
                        'email': row[27].strip().lower(),
                    }
                    count += 1
            print(f'    → {count} empresas encontradas neste arquivo')
        except Exception as e:
            print(f'    Erro: {e}')
        finally:
            if path and os.path.exists(path):
                os.unlink(path)

    print(f'    Total acumulado: {len(estab)} empresas')

    # 4. Razões sociais
    print('[4/4] Buscando razões sociais (10 arquivos)…')
    cnpj_basicos = set(estab.keys())
    porte_map = {'01': 'ME', '02': 'ME', '03': 'EPP', '05': 'Grande'}
    for i in range(10):
        url = f'{RFB_BASE}Empresas{i}.zip'
        path = None
        try:
            path = baixar_zip(url, f'Empresas{i}.zip')
            for row in ler_csv_zip(path):
                if len(row) >= 6 and row[0].strip() in cnpj_basicos:
                    estab[row[0].strip()]['razao_social'] = row[1].strip().title()
                    estab[row[0].strip()]['porte'] = porte_map.get(row[5].strip(), '')
        except Exception as e:
            print(f'    Erro: {e}')
        finally:
            if path and os.path.exists(path):
                os.unlink(path)

    # 5. Gerar CSV
    nome_arquivo = f'empresas_{norm(MUNICIPIO).replace(" ", "_")}.csv'
    city_title = MUNICIPIO.title()
    with open(nome_arquivo, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow([
            'empresa_nome', 'cnpj', 'cnae_codigo', 'cnae_descricao',
            'porte', 'cidade', 'bairro', 'endereco', 'telefone', 'email',
        ])
        for e in estab.values():
            cnpj = e.get('cnpj', '')
            if len(cnpj) != 14:
                continue
            writer.writerow([
                e.get('razao_social') or e.get('nome_fantasia') or 'Sem Nome',
                cnpj,
                e.get('cnae', ''),
                cnae_desc.get(e.get('cnae', ''), ''),
                e.get('porte', ''),
                city_title,
                e.get('bairro', ''),
                e.get('endereco', ''),
                e.get('telefone', ''),
                e.get('email', ''),
            ])

    print(f'\n✅ Arquivo gerado: {nome_arquivo}')
    print(f'   {len(estab)} empresas exportadas.')
    print('\nPróximo passo: importe o CSV em Comercial → Prospecção → Importar CSV\n')


if __name__ == '__main__':
    main()
