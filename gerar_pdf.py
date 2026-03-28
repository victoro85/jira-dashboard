"""
Gera Apresentacao_Jira.pdf a partir do Apresentacao_Jira.html
Requer (executar uma vez):
    pip install playwright Pillow
    playwright install chromium
"""
import os
import io
import sys

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print('ERRO: playwright nao instalado.')
    print('Execute:  pip install playwright && playwright install chromium')
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print('ERRO: Pillow nao instalado.')
    print('Execute:  pip install Pillow')
    sys.exit(1)

BASE    = os.path.dirname(os.path.abspath(__file__))
HTML_IN = os.path.join(BASE, 'Apresentacao_Jira.html')
PDF_OUT = os.path.join(BASE, 'Apresentacao_Jira.pdf')

WIDTH, HEIGHT = 1280, 720
SCALE         = 2     # captura em 2560x1440 — bem mais nítido

if not os.path.exists(HTML_IN):
    print(f'ERRO: {HTML_IN} nao encontrado. Gere o HTML primeiro.')
    sys.exit(1)

# Contar slides no HTML
with open(HTML_IN, encoding='utf-8') as f:
    conteudo = f.read()
N_SLIDES = conteudo.count('class="slide-container')

print(f'Capturando {N_SLIDES} slides em {WIDTH * SCALE}x{HEIGHT * SCALE}px...')

# JS para trocar slide diretamente (sem animação)
JS_SHOW = """(function(i) {{
    document.querySelectorAll('.slide-container').forEach(function(s, idx) {{
        s.classList.toggle('active', idx === i);
    }});
}})(arguments[0])"""

images = []
with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(
        viewport={'width': WIDTH, 'height': HEIGHT},
        device_scale_factor=SCALE
    )

    url = 'file:///' + HTML_IN.replace('\\', '/')
    page.goto(url)
    page.wait_for_load_state('networkidle')
    page.wait_for_timeout(1500)  # aguarda fontes Google carregarem

    # Desabilita animações CSS para captura limpa
    page.add_style_tag(content="""
        * { animation-duration: 0s !important;
            animation-delay: 0s !important;
            transition-duration: 0s !important; }
    """)

    for i in range(N_SLIDES):
        # Troca slide via JS direto — sem dependência de teclado ou temporizadores
        page.evaluate(f"""
            document.querySelectorAll('.slide-container').forEach(function(s, idx) {{
                s.classList.toggle('active', idx === {i});
            }});
        """)
        page.wait_for_timeout(150)  # aguarda repintura

        shot = page.screenshot(type='png')
        img  = Image.open(io.BytesIO(shot)).convert('RGB')
        images.append(img)
        print(f'  Slide {i + 1}/{N_SLIDES}')

    browser.close()

# DPI = 144 * SCALE para manter tamanho proporcional ao 1280x720
DPI = 144 * SCALE
images[0].save(
    PDF_OUT, 'PDF',
    save_all=True,
    append_images=images[1:],
    resolution=DPI
)

tam = round(os.path.getsize(PDF_OUT) / 1024, 1)
print(f'Gerado: Apresentacao_Jira.pdf  ({tam} KB, {N_SLIDES} paginas)')
