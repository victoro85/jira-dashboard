"""
JIRA Dashboard - Gerador Completo
Executa em sequência:
  1. gerar_relatorio.py       → relatorio_jira_completo.xlsx
  2. adicionar_precificacao.py → abas de custo no Excel
  3. gerar_html.py             → Apresentacao_Jira.html
  4. gerar_pdf.py              → Apresentacao_Jira.pdf
"""
import subprocess
import sys
import os
import time

SCRIPTS = [
    ('gerar_relatorio.py',        'Gerando relatório Excel...'),
    ('adicionar_precificacao.py', 'Adicionando precificação...'),
    ('gerar_html.py',             'Gerando apresentação HTML...'),
    ('gerar_pdf.py',              'Gerando apresentação PDF...'),
]

BASE = os.path.dirname(os.path.abspath(__file__))

def linha(char='-', n=52):
    print(char * n)

def rodar(script, descricao):
    print(f'\n>  {descricao}')
    inicio = time.time()
    result = subprocess.run(
        [sys.executable, os.path.join(BASE, script)],
        capture_output=True, text=True, cwd=BASE
    )
    duracao = round(time.time() - inicio, 1)

    if result.returncode == 0:
        print(f'   OK Concluído em {duracao}s')
        if result.stdout.strip():
            for linha_out in result.stdout.strip().splitlines():
                print(f'      {linha_out}')
    else:
        print(f'   ERRO ERRO em {script}:')
        for linha_err in result.stderr.strip().splitlines():
            print(f'      {linha_err}')
        sys.exit(1)

# ── Início ────────────────────────────────────────────────────────────────────
os.system('cls' if os.name == 'nt' else 'clear')
linha('=')
print('  JIRA DASHBOARD — GERADOR COMPLETO')
linha('=')
print(f'  Pasta: {BASE}')
linha()

# Verificar arquivos de entrada
entradas = ['export.csv', 'JIRA_separado.xlsx', 'rh_atual.xlsx']
print('\n  Verificando arquivos de entrada:')
faltando = []
for arq in entradas:
    caminho = os.path.join(BASE, arq)
    if os.path.exists(caminho):
        tam = round(os.path.getsize(caminho) / 1024, 1)
        print(f'   OK {arq} ({tam} KB)')
    else:
        print(f'   ERRO {arq} — NÃO ENCONTRADO')
        faltando.append(arq)

if faltando:
    linha()
    print('\n  Arquivos em falta. Coloque-os na pasta e tente novamente.')
    input('\n  Pressione Enter para sair...')
    sys.exit(1)

# Rodar scripts
linha()
for script, descricao in SCRIPTS:
    rodar(script, descricao)

# ── Resultado ─────────────────────────────────────────────────────────────────
linha()
print('\n  TUDO GERADO COM SUCESSO!\n')

saidas = [
    ('relatorio_jira_completo.xlsx', 'Relatório Excel'),
    ('Apresentacao_Jira.html',       'Apresentação HTML'),
    ('Apresentacao_Jira.pdf',        'Apresentação PDF'),
]
for arq, nome in saidas:
    caminho = os.path.join(BASE, arq)
    if os.path.exists(caminho):
        tam = round(os.path.getsize(caminho) / 1024, 1)
        print(f'   >> {nome}')
        print(f'       {caminho}  ({tam} KB)')

linha('=')
input('\n  Pressione Enter para fechar...')
