from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

OUT = r"C:\Users\davi62985616\Desktop\Pre Hackathon\NORDEN_Pitch_3min.pptx"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(26, 35, 48)
SLATE = RGBColor(108, 114, 124)
ICE = RGBColor(244, 242, 238)
WHITE = RGBColor(255, 255, 255)
GOLD = RGBColor(142, 107, 71)
GREEN = RGBColor(91, 143, 123)
LIGHT = RGBColor(217, 214, 209)


def rect(slide, x, y, w, h, fill, radius=False, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line or fill
    return shape


def text(slide, s, x, y, w, h, size=24, color=NAVY, bold=False,
         font="Aptos", align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.text = s
    p.alignment = align
    p.font.name = font
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    return box


def rich_text(slide, lines, x, y, w, h, size=20, color=NAVY, bullet=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.margin_left = Inches(0.02); tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.space_after = Pt(10)
        if bullet:
            p.text = "• " + item
    return box


def title(slide, kicker, heading):
    text(slide, kicker.upper(), 0.65, 0.43, 3.8, 0.25, 11, GOLD, True)
    text(slide, heading, 0.65, 0.78, 12.0, 0.75, 30, NAVY, True)
    rect(slide, 0.65, 1.58, 12.03, 0.02, LIGHT)


def footer(slide, n):
    text(slide, "NORDEN  |  PITCH", 0.65, 7.13, 2.5, 0.18, 9, SLATE, True)
    text(slide, f"{n:02d}", 12.08, 7.10, 0.4, 0.20, 9, SLATE, True, align=PP_ALIGN.RIGHT)


# 1 — Abertura
slide = prs.slides.add_slide(prs.slide_layouts[6])
rect(slide, 0, 0, 13.333, 7.5, NAVY)
rect(slide, 9.35, 0, 3.983, 7.5, RGBColor(34, 48, 65))
text(slide, "NORDEN", 0.72, 1.05, 6.8, 0.75, 54, WHITE, True)
text(slide, "Men's Care & Experience", 0.75, 1.83, 6.8, 0.38, 15, LIGHT, False)
text(slide, "Seu estilo.", 0.75, 3.05, 6.2, 0.65, 34, WHITE, True)
text(slide, "Sua evolução.", 0.75, 3.72, 6.2, 0.65, 34, GOLD, True)
text(slide, "Um salão inteligente que conecta agendamento, atendimento, cuidado e compras.", 0.75, 4.72, 7.45, 0.90, 20, WHITE, False)
# abstract journey motif
for i, (lab, xx) in enumerate([("ANTES", 9.95), ("DURANTE", 10.95), ("DEPOIS", 12.05)]):
    rect(slide, xx, 2.15 + i*1.25, 0.62, 0.62, GOLD if i==1 else SLATE, True)
    text(slide, lab, xx-0.15, 2.95 + i*1.25, 1.0, 0.25, 8, WHITE, True, align=PP_ALIGN.CENTER)
text(slide, "NORDEN", 0.75, 6.68, 2.0, 0.22, 9, LIGHT, True)

# 2 — Dor
slide = prs.slides.add_slide(prs.slide_layouts[6])
rect(slide, 0, 0, 13.333, 7.5, ICE)
title(slide, "O problema", "Hoje, o cuidado masculino é fragmentado")
items = [
    ("01", "Agenda manual", "Conversas, remarcações e horários ociosos."),
    ("02", "Sem continuidade", "Serviços feitos ficam sem histórico acessível."),
    ("03", "Cuidados esquecidos", "O cliente não sabe como manter o resultado."),
    ("04", "Compra desconectada", "Produtos recomendados são difíceis de encontrar depois."),
]
for i, (num, head, desc) in enumerate(items):
    x = 0.65 + i*3.05
    rect(slide, x, 2.05, 2.72, 3.75, WHITE, True)
    text(slide, num, x+0.22, 2.28, 0.45, 0.35, 14, GOLD, True)
    text(slide, head, x+0.22, 3.05, 2.25, 0.62, 20, NAVY, True)
    text(slide, desc, x+0.22, 4.05, 2.25, 1.15, 16, SLATE, False)
text(slide, "Hipóteses de produto a serem validadas com o mercado.", 0.65, 6.35, 8.5, 0.28, 10, SLATE, False)
footer(slide, 2)

# 3 — Jornada
slide = prs.slides.add_slide(prs.slide_layouts[6])
rect(slide, 0, 0, 13.333, 7.5, WHITE)
title(slide, "A solução", "Uma única jornada: antes, durante e depois")
steps = [
    ("1", "AGENDA", "Escolha serviço, profissional e horário"),
    ("2", "ATENDIMENTO", "Preferências, procedimentos e produtos registrados"),
    ("3", "EVOLUÇÃO", "Linha do tempo e próxima manutenção"),
    ("4", "CUIDADO", "Manual personalizado + loja integrada"),
]
for i, (num, head, desc) in enumerate(steps):
    x = 0.75 + i*3.08
    rect(slide, x, 2.20, 2.52, 2.95, ICE, True)
    rect(slide, x+0.20, 2.43, 0.50, 0.50, NAVY, True)
    text(slide, num, x+0.20, 2.52, 0.50, 0.24, 13, WHITE, True, align=PP_ALIGN.CENTER)
    text(slide, head, x+0.20, 3.30, 2.10, 0.32, 13, GOLD, True)
    text(slide, desc, x+0.20, 3.82, 2.10, 1.05, 16, NAVY, False)
    if i < 3:
        text(slide, "→", x+2.52, 3.18, 0.52, 0.50, 25, SLATE, True, align=PP_ALIGN.CENTER)
rect(slide, 0.75, 5.55, 11.76, 0.76, NAVY, True)
text(slide, "Resultado: menos esforço operacional + mais continuidade para o cliente.", 1.02, 5.77, 11.20, 0.33, 19, WHITE, True, align=PP_ALIGN.CENTER)
footer(slide, 3)

# 4 — Diferenciais
slide = prs.slides.add_slide(prs.slide_layouts[6])
rect(slide, 0, 0, 13.333, 7.5, NAVY)
text(slide, "QUATRO DIFERENCIAIS", 0.65, 0.43, 4.0, 0.25, 11, GOLD, True)
text(slide, "Tecnologia a serviço da experiência", 0.65, 0.78, 11.8, 0.75, 30, WHITE, True)
rect(slide, 0.65, 1.58, 12.03, 0.02, SLATE)
features = [
    ("FLOW", "Realocação", "Cancelou? Sugere e oferece um novo horário sem alterar a reserva automaticamente."),
    ("JOURNEY", "Histórico", "Organiza procedimentos, produtos usados e próxima manutenção."),
    ("CARE", "Cuidados", "Cria um manual personalizado para preservar o resultado em casa."),
    ("STORE", "Comércio", "Vende os produtos recomendados, com retirada ou entrega."),
]
for i, (tag, head, desc) in enumerate(features):
    x = 0.65 + (i%2)*6.02
    y = 1.93 + (i//2)*2.10
    rect(slide, x, y, 5.72, 1.76, RGBColor(34, 48, 65), True)
    text(slide, tag, x+0.28, y+0.20, 1.20, 0.25, 11, GOLD, True)
    text(slide, head, x+0.28, y+0.62, 2.10, 0.38, 21, WHITE, True)
    text(slide, desc, x+2.10, y+0.48, 3.25, 0.82, 15, LIGHT, False)
footer(slide, 4)

# 5 — Benefícios
slide = prs.slides.add_slide(prs.slide_layouts[6])
rect(slide, 0, 0, 13.333, 7.5, ICE)
title(slide, "Valor entregue", "Melhora para quem atende — e para quem retorna")
# left persona
rect(slide, 0.65, 1.92, 5.72, 4.58, WHITE, True)
text(slide, "CLIENTE", 0.95, 2.18, 1.2, 0.22, 11, GOLD, True)
text(slide, "Menos atrito", 0.95, 2.72, 3.7, 0.45, 23, NAVY, True)
rich_text(slide, ["Agenda em minutos", "Entende o que foi feito", "Mantém o resultado", "Compra apenas o que faz sentido"], 0.95, 3.42, 4.90, 2.20, 17, NAVY, True)
# right persona
rect(slide, 6.62, 1.92, 6.06, 4.58, NAVY, True)
text(slide, "GESTOR & PROFISSIONAL", 6.92, 2.18, 3.1, 0.22, 11, GOLD, True)
text(slide, "Mais eficiência", 6.92, 2.72, 3.8, 0.45, 23, WHITE, True)
rich_text(slide, ["Agenda organizada", "Menos horários ociosos", "Atendimento registrado", "Pedidos e estoque conectados"], 6.92, 3.42, 5.20, 2.20, 17, WHITE, True)
footer(slide, 5)

# 6 — MVP / convite
slide = prs.slides.add_slide(prs.slide_layouts[6])
rect(slide, 0, 0, 13.333, 7.5, WHITE)
title(slide, "Escopo inicial", "O MVP integra todo o ciclo")
# central loop
modules = ["Agendamento", "Atendimento", "Journey", "Care", "Store"]
for i, mod in enumerate(modules):
    x = 0.68 + i*2.45
    rect(slide, x, 2.05, 2.08, 1.12, NAVY if i in (0,4) else ICE, True)
    text(slide, mod, x+0.12, 2.42, 1.84, 0.30, 16, WHITE if i in (0,4) else NAVY, True, align=PP_ALIGN.CENTER)
    if i < 4:
        text(slide, "→", x+2.09, 2.36, 0.36, 0.34, 18, GOLD, True, align=PP_ALIGN.CENTER)
rect(slide, 0.68, 3.75, 12.00, 1.15, ICE, True)
text(slide, "Segurança, gestão, estoque e pagamentos simulados desde o primeiro ciclo.", 0.95, 4.12, 11.45, 0.38, 20, NAVY, True, align=PP_ALIGN.CENTER)
text(slide, "NORDEN", 0.68, 5.48, 3.0, 0.55, 38, NAVY, True)
text(slide, "Seu estilo. Sua evolução.", 0.68, 6.08, 6.2, 0.38, 18, GOLD, True)
text(slide, "Projeto preparado para validação, demonstração e implementação.", 7.18, 5.68, 5.28, 0.58, 16, SLATE, False, align=PP_ALIGN.RIGHT)
footer(slide, 6)

prs.save(OUT)
print(OUT)
