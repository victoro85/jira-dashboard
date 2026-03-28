"""
Gera Apresentacao_Jira.html com dados reais do relatorio_jira_completo.xlsx
Padrão visual: dark dashboard, valores SEMPRE fora das barras (direita, alinhados)
"""
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

# ── Dados — Relatorio Completo já contém todos os usuários ───────────────────
df = pd.read_excel('relatorio_jira_completo.xlsx', sheet_name='Relatorio Completo')
df['Diretoria'] = df['Diretoria'].fillna('NÃO LOCALIZADO').str.upper().str.strip()

def has_user(val):
    return pd.notna(val) and 'User' in str(val)

PRODS = {
    'Jira'                    : {'cor': '#CCFF00', 'label': 'Jira',        'th': 'Jira'},
    'Confluence'              : {'cor': '#00BFFF', 'label': 'Confluence',  'th': 'Confluence'},
    'Jira Service Management' : {'cor': '#FF6B35', 'label': 'Jira Service Management', 'th': 'Jira Service Management'},
    'Jira Product Discovery'  : {'cor': '#BF7FFF', 'label': 'Jira Product Discovery',  'th': 'Jira Product Discovery'},
}

PRECOS = {
    'Jira'                    : (1907.20, 220),
    'Confluence'              : (1252.50, 205),
    'Jira Service Management' : (1258.43,  54),
    'Jira Product Discovery'  : ( 711.18,  67),
    'Compass'                 : (  38.19,   4),
}
CUSTO_USER = {p: round(t/q, 4) for p,(t,q) in PRECOS.items()}

# Métricas gerais
total       = len(df)
ativos      = int((df['Status'] == 'Active').sum())
saneado_pct = round(df['Cargo'].notna().sum() / total * 100)
total_lic   = sum(df[p].apply(has_user).sum() for p in PRODS)

# Licenças por produto
prod_counts = {p: int(df[p].apply(has_user).sum()) for p in PRODS}

# Distribuição por diretoria
dc = df['Diretoria'].value_counts()

# Custo por diretoria
dir_custo = {}
for d, grp in df.groupby('Diretoria'):
    total_d = sum(grp[p].apply(has_user).sum() * CUSTO_USER.get(p, 0) for p in PRODS)
    dir_custo[d] = round(total_d, 2)
dir_custo_sorted = dict(sorted(dir_custo.items(), key=lambda x: -x[1]))

total_mensal = sum(PRECOS[p][0] for p in PRECOS)
total_anual  = round(total_mensal * 12, 2)

TOP = 10

def fmt_usd(v):
    return f'US$ {int(v):,}'.replace(',', '.')

# ── Helpers HTML ──────────────────────────────────────────────────────────────
def bar(label, value, max_val, is_outros=False, raw_val=None):
    num = raw_val if raw_val is not None else value
    pct = max(6, round(float(num) / float(max_val) * 100))
    bg  = 'background:#444;' if is_outros else ''
    lbl_color = 'color:#888;' if is_outros else ''
    return f'''
            <div class="bar">
                <div class="label">{label}</div>
                <div class="bar-inner"><div class="bar-fill" style="width:{pct}%;{bg}"></div></div>
                <span class="bar-label-out" style="{lbl_color}">{value}</span>
            </div>'''

def badge(n, cor):
    if n == 0:
        return '<span class="badge" style="background:#1a1a1a;color:#333;border:1px solid #2a2a2a;">—</span>'
    return f'<span class="badge" style="background:{cor}22;color:{cor};border:1px solid {cor}44;">{n}</span>'

def mini_card(n, label, cor):
    return f'''            <div class="mini-card" style="border-color:{cor}33;">
                <span class="mini-val" style="color:{cor};">{n}</span>
                <span class="mini-lbl">{label}</span>
            </div>'''

# ── Slide 3: Distribuição por Diretoria — TODAS ──────────────────────────────
max_dir  = int(dc.max())
n_dist   = len(dc)
bars_dist = ''
for d, v in dc.items():
    bars_dist += bar(d, int(v), max_dir)

if n_dist <= 10:
    dist_chart_class = 'chart-container'
elif n_dist <= 14:
    dist_chart_class = 'chart-container chart-xs'
else:
    dist_chart_class = 'chart-container chart-xxs'

# ── Slide 4: Custo por Diretoria — TODAS as diretorias ───────────────────────
max_custo   = float(list(dir_custo_sorted.values())[0])
items_custo = list(dir_custo_sorted.items())
n_custo     = len(items_custo)
bars_custo  = ''
for d, v in items_custo:
    bars_custo += bar(d, fmt_usd(v), max_custo, raw_val=float(v))

# Classe de barra dinâmica: quanto mais diretorias, menor a barra
if n_custo <= 10:
    custo_chart_class = 'chart-sm'
elif n_custo <= 14:
    custo_chart_class = 'chart-xs'
else:
    custo_chart_class = 'chart-xxs'

# ── Slides 5 e 6: Licenças por Diretoria — dividido em 2 partes ──────────────
all_dirs = dc.index.tolist()
SPLIT    = 10  # primeiras 10 no slide 5, restante no slide 6

def make_rows(dirs_list):
    rows = ''
    for d in dirs_list:
        sub   = df[df['Diretoria'] == d]
        n_usr = int(dc[d])
        cells = ''.join(f'<td>{badge(int(sub[p].apply(has_user).sum()), PRODS[p]["cor"])}</td>' for p in PRODS)
        rows += f'<tr><td>{d}</td><td class="total-cell">{n_usr}</td>{cells}</tr>\n'
    return rows

prod_rows_1 = make_rows(all_dirs[:SPLIT])
prod_rows_2 = make_rows(all_dirs[SPLIT:]) if len(all_dirs) > SPLIT else ''

# Cabeçalho da tabela pré-computado (usado em ambos os slides)
tbl_headers = '<th>Diretoria</th><th>Usuários</th>' + ''.join(
    f'<th style="color:{PRODS[p]["cor"]};">{PRODS[p]["th"]}</th>' for p in PRODS
)
tbl_colgroup = '<col class="c-dir"><col class="c-usr"><col class="c-prod"><col class="c-prod"><col class="c-prod"><col class="c-prod">'

def tbl_slide(title, rows, padding='50px 80px 70px'):
    return f'''<div class="slide-container" style="padding:{padding};">
    <h2 class="slide-title" style="margin-bottom:20px;">{title}</h2>
    <div class="content-area">
        <table class="prod-table">
            <colgroup>{tbl_colgroup}</colgroup>
            <thead><tr>{tbl_headers}</tr></thead>
            <tbody>{rows}</tbody>
        </table>
    </div>
</div>'''

slide_lic_1 = tbl_slide('Licenças por Diretoria', prod_rows_1)
slide_lic_2 = tbl_slide('Licenças por Diretoria', prod_rows_2) if prod_rows_2 else ''

# ── Slide 6: Totais de licenças ───────────────────────────────────────────────
prod_total_cards = ''
for p, info in PRODS.items():
    n   = prod_counts[p]
    pct = round(n / total * 100)
    prod_total_cards += f'''
            <div class="prod-card" style="border-color:{info['cor']}44;">
                <span class="prod-value" style="color:{info['cor']};">{n}</span>
                <span class="prod-pct" style="color:{info['cor']}88;">{pct}%</span>
                <span class="prod-label">{info['label'].replace('<br>',' ')}</span>
            </div>'''

# ── CSS ───────────────────────────────────────────────────────────────────────
CSS = """
        * { box-sizing: border-box; }
        body {
            background-color: #000;
            margin: 0; padding: 0; overflow: hidden;
            display: flex; justify-content: center; align-items: center; height: 100vh;
        }
        .slide-container {
            display: none; flex-direction: column; align-items: flex-start;
            background-color: #000; border-radius: 20px;
            box-shadow: 0 40px 80px rgba(0,0,0,0.9);
            font-family: 'Urbanist', sans-serif;
            height: 720px; width: 1280px;
            padding: 80px; position: relative;
            border: 1px solid #1a1a1a; justify-content: center;
        }
        .slide-container.active { display: flex; animation: slideIn 0.3s ease-out both; }
        @keyframes slideIn { from { opacity:0; } to { opacity:1; } }
        .slide-container::before {
            content:''; position:absolute; top:-20%; right:-10%;
            width:600px; height:600px;
            background:radial-gradient(circle, rgba(222,255,154,0.1) 0%, transparent 70%); z-index:0;
        }
        .slide-container::after {
            content:''; position:absolute; bottom:15px; left:80px; right:80px;
            height:2px; background:linear-gradient(90deg,transparent,#deff9a,transparent); opacity:0.3;
        }
        .slide-container > * { position:relative; z-index:2; }

        h1,h2,h3 { color:#fff; font-weight:800; margin:0; }
        h1 { font-size:150px; letter-spacing:-5px; line-height:0.8; }
        .slide-title {
            font-size:52px; margin-bottom:40px; width:100%;
            border-left:12px solid #deff9a; padding-left:30px; text-transform:uppercase;
        }
        .slide-title-sm { font-size:42px; margin-bottom:16px; }
        .content-area { display:flex; flex-direction:column; flex-grow:1; justify-content:center; width:100%; }

        /* MÉTRICAS */
        .highlight-numbers { display:flex; justify-content:space-between; width:100%; gap:30px; }
        .number-card {
            background:#0a0a0a; border:2px solid #222; border-radius:35px;
            padding:60px 30px; width:50%; text-align:center;
        }
        .number-card .value { font-size:110px; color:#deff9a; font-weight:800; display:block; line-height:1; margin-bottom:15px; }
        .number-card .label { font-size:22px; text-transform:uppercase; letter-spacing:5px; color:#555; }

        /* MINI CARDS */
        .mini-cards { display:flex; gap:14px; margin-top:28px; width:100%; }
        .mini-card {
            flex:1; background:#0d0d0d; border:1px solid #222; border-radius:14px;
            padding:16px 12px; text-align:center;
        }
        .mini-val { display:block; font-size:30px; font-weight:800; line-height:1; margin-bottom:6px; }
        .mini-lbl { font-size:10px; color:#555; text-transform:uppercase; letter-spacing:1.5px; }

        /* BARRAS */
        .chart-container { display:flex; flex-direction:column; gap:9px; width:100%; }
        .bar { align-items:center; display:flex; gap:20px; }
        .bar .label {
            color:#aaa; flex:0 0 300px; font-weight:700; text-align:right;
            font-size:15px; font-family:monospace; text-transform:uppercase;
            letter-spacing:1px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
        }
        .bar .bar-inner {
            background-color:#0f0f0f; border-radius:10px; flex-grow:1;
            height:50px; border:1px solid #1e1e1e;
            display:flex; align-items:center; overflow:hidden;
        }
        .bar .bar-fill {
            height:100%; border-radius:8px;
            background:linear-gradient(90deg,#deff9a,#a8cf5f);
            box-shadow:0 0 16px rgba(222,255,154,0.2);
        }
        .bar-label-out {
            color:#deff9a; font-weight:800; font-size:20px;
            white-space:nowrap; flex:0 0 130px; text-align:right;
        }
        .chart-sm  { gap:5px !important; }
        .chart-sm  .bar .bar-inner { height:38px; }
        .chart-sm  .bar-label-out  { font-size:17px; }
        .chart-sm  .bar .label     { font-size:13px; }
        .chart-xs  { gap:4px !important; }
        .chart-xs  .bar .bar-inner { height:30px; }
        .chart-xs  .bar-label-out  { font-size:14px; }
        .chart-xs  .bar .label     { font-size:11px; }
        .chart-xxs { gap:3px !important; }
        .chart-xxs .bar .bar-inner { height:28px; }
        .chart-xxs .bar-label-out  { font-size:13px; }
        .chart-xxs .bar .label     { font-size:11px; }

        /* TABELA PRODUTOS */
        .prod-table {
            width:100%; border-collapse:collapse; table-layout:fixed;
        }
        .prod-table col.c-dir  { width:22%; }
        .prod-table col.c-usr  { width:8%; }
        .prod-table col.c-prod { width:17.5%; }
        .prod-table thead tr {
            border-bottom:2px solid #222;
        }
        .prod-table thead th {
            text-align:center; padding:0 4px 12px;
            font-size:10px; text-transform:uppercase;
            letter-spacing:0.5px; vertical-align:bottom;
            white-space:nowrap; overflow:hidden;
        }
        .prod-table thead th:first-child { text-align:left; color:#deff9a; }
        .prod-table thead th:nth-child(2) { color:#666; }
        .prod-table tbody tr { height:52px; }
        .prod-table tbody tr + tr td { border-top:1px solid #111; }
        .prod-table tbody td {
            background:#0d0d0d; padding:0 4px;
            text-align:center; vertical-align:middle;
        }
        .prod-table tbody td:first-child {
            border-radius:10px 0 0 10px; color:#fff; font-family:monospace;
            font-size:13px; font-weight:700; text-align:left; padding-left:14px;
        }
        .prod-table tbody td:last-child { border-radius:0 10px 10px 0; }
        .total-cell { color:#555 !important; font-size:13px !important; }
        .badge {
            display:inline-flex; align-items:center; justify-content:center;
            width:40px; height:40px; border-radius:50%; font-weight:800; font-size:15px;
        }
        /* Tabela compacta para muitas diretorias */
        .prod-table-sm tbody tr { height:36px; }
        .prod-table-sm tbody td:first-child { font-size:11px; }
        .prod-table-sm .badge { width:30px; height:30px; font-size:12px; }
        .prod-table-sm .total-cell { font-size:11px !important; }
        .prod-table-xs tbody tr { height:28px; }
        .prod-table-xs tbody td:first-child { font-size:10px; padding-left:8px; }
        .prod-table-xs .badge { width:24px; height:24px; font-size:10px; }
        .prod-table-xs .total-cell { font-size:10px !important; }

        /* CARDS PRODUTO */
        .prod-cards { display:flex; gap:30px; width:100%; justify-content:center; }
        .prod-card {
            background:#0a0a0a; border:2px solid #222; border-radius:25px;
            padding:40px 30px; flex:1; text-align:center;
        }
        .prod-value { font-size:80px; font-weight:800; display:block; line-height:1; }
        .prod-pct   { font-size:24px; font-weight:700; display:block; margin:8px 0 12px; }
        .prod-label { font-size:16px; text-transform:uppercase; letter-spacing:3px; color:#555; display:block; }

        /* NAV */
        .nav-hint {
            position:fixed; bottom:25px; right:40px; color:#333; font-size:14px;
            text-transform:uppercase; letter-spacing:3px; font-weight:600; font-family:'Urbanist',sans-serif;
        }
        .slide-counter {
            position:fixed; bottom:25px; left:40px; color:#333; font-size:14px;
            letter-spacing:3px; font-family:'Urbanist',sans-serif; font-weight:600;
        }
"""

# ── HTML ──────────────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Jira - Análise de Centros de Custo</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Urbanist:wght@400;600;800&display=swap" rel="stylesheet">
    <style>{CSS}</style>
</head>
<body>

<!-- SLIDE 1: CAPA -->
<div class="slide-container active">
    <div style="text-align:center; width:100%;">
        <h1>JIRA</h1>
        <h3 style="letter-spacing:12px; text-transform:uppercase; color:#deff9a; margin-top:15px; font-size:35px;">
            Análise de Centros de Custo
        </h3>
        <p style="margin-top:50px; font-size:30px; color:#555;">Dashboard Executivo · Março 2026</p>
    </div>
</div>

<!-- SLIDE 2: MÉTRICAS GERAIS -->
<div class="slide-container">
    <h2 class="slide-title">Métricas Gerais</h2>
    <div class="content-area">
        <div class="highlight-numbers">
            <div class="number-card">
                <span class="value">{total}</span>
                <span class="label">Usuários Ativos</span>
            </div>
            <div class="number-card">
                <span class="value">{int(total_lic)}</span>
                <span class="label">Licenças Totais</span>
            </div>
        </div>
        <div class="mini-cards">
{chr(10).join(mini_card(prod_counts[p], PRODS[p]['label'].replace('<br>',' '), PRODS[p]['cor']) for p in PRODS)}
        </div>
    </div>
</div>

<!-- SLIDE 3: DISTRIBUIÇÃO POR DIRETORIA -->
<div class="slide-container" style="padding:50px 80px 75px;">
    <h2 class="slide-title slide-title-sm">Distribuição por Diretoria</h2>
    <div class="content-area">
        <div class="{dist_chart_class}">
{bars_dist}
        </div>
    </div>
</div>

<!-- SLIDE 4: CUSTO POR DIRETORIA -->
<div class="slide-container" style="padding:30px 80px 75px;">
    <h2 class="slide-title slide-title-sm" style="font-size:34px; margin-bottom:6px;">Custo por Diretoria</h2>
    <div style="display:flex; align-items:baseline; gap:16px; margin-bottom:8px;">
        <span style="font-size:11px; color:#555; text-transform:uppercase; letter-spacing:2px;">Total mensal</span>
        <span style="font-size:22px; font-weight:800; color:#deff9a;">{fmt_usd(total_mensal)}</span>
        <span style="font-size:11px; color:#555; text-transform:uppercase; letter-spacing:2px; margin-left:12px;">Anual estimado</span>
        <span style="font-size:22px; font-weight:800; color:#555;">{fmt_usd(total_anual)}</span>
    </div>
    <div class="content-area">
        <div class="chart-container {custo_chart_class}">
{bars_custo}
        </div>
    </div>
</div>

{slide_lic_1}

{slide_lic_2}

<!-- SLIDE 6: TOTAL DE LICENÇAS -->
<div class="slide-container">
    <h2 class="slide-title">Total de Licenças Ativas</h2>
    <div class="content-area">
        <div class="prod-cards">
{prod_total_cards}
        </div>
    </div>
</div>

<!-- SLIDE 7: ENCERRAMENTO -->
<div class="slide-container" style="text-align:center;">
    <div style="width:100%;">
        <h1 style="color:#deff9a;">FIM.</h1>
        <p style="font-size:36px; font-weight:600; color:#fff; margin-top:20px;">Dúvidas ou Esclarecimentos</p>
        <div style="margin-top:100px; padding:30px 80px; border:3px solid #deff9a; display:inline-block; border-radius:100px;">
            <span style="color:#fff; font-size:22px; letter-spacing:8px;">TI ASSETS MANAGEMENT | 2026</span>
        </div>
    </div>
</div>


<div class="slide-counter" id="counter">1 / 7</div>

<script>
    let current = 0;
    let animating = false;
    const slides = document.querySelectorAll('.slide-container');
    const counter = document.getElementById('counter');
    function show(i) {{
        if (animating) return;
        animating = true;
        slides.forEach((s, idx) => s.classList.toggle('active', idx === i));
        counter.textContent = (i + 1) + ' / ' + slides.length;
        setTimeout(() => {{ animating = false; }}, 350);
    }}
    window.addEventListener('keydown', e => {{
        if (e.key === 'ArrowRight' && current < slides.length - 1) show(++current);
        else if (e.key === 'ArrowLeft' && current > 0) show(--current);
    }});
    window.addEventListener('click', e => {{
        if (e.clientX > window.innerWidth / 2) {{ if (current < slides.length - 1) show(++current); }}
        else {{ if (current > 0) show(--current); }}
    }});
</script>
</body>
</html>"""

with open('Apresentacao_Jira.html', 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Gerado: Apresentacao_Jira.html')
print(f'Slides : 7')
print(f'Usuários: {total} | Licenças: {int(total_lic)} | Saneado: {saneado_pct}%')
print(f'Custo mensal: US$ {total_mensal:,.2f} | Anual: US$ {total_anual:,.2f}')
