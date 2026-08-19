from pathlib import Path
from textwrap import wrap

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "pdf" / "GLP_Tutorial_Lanzamiento.pdf"
LOGO = ROOT / "assets" / "branding" / "logo_blanco.png"
ISOTYPE = ROOT / "assets" / "branding" / "logo_isotipo.png"

PAGE_W, PAGE_H = A4
MARGIN = 42

GREEN = HexColor("#6B7B38")
DARK_GREEN = HexColor("#33492D")
BLUE = HexColor("#1B4B6C")
PALE_BLUE = HexColor("#DAE0E5")
LIME = HexColor("#E2FF87")
CREAM = HexColor("#FFEAC6")
INK = HexColor("#1D2B24")
MUTED = HexColor("#5D6A63")
PAPER = HexColor("#F7F8F5")
WHITE = white


def register_fonts():
    montserrat = Path(r"C:\Windows\Fonts\Montserrat.ttf")
    arial = Path(r"C:\Windows\Fonts\arial.ttf")
    arial_bold = Path(r"C:\Windows\Fonts\arialbd.ttf")
    if montserrat.exists():
        pdfmetrics.registerFont(TTFont("Montserrat", str(montserrat)))
    if arial.exists():
        pdfmetrics.registerFont(TTFont("Inter", str(arial)))
    if arial_bold.exists():
        pdfmetrics.registerFont(TTFont("Inter-Bold", str(arial_bold)))


def rounded(c, x, y, w, h, fill, radius=12, stroke=None, width=1):
    c.setLineWidth(width)
    c.setFillColor(fill)
    c.setStrokeColor(stroke or fill)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=1 if stroke else 0)


def text(c, value, x, y, size=10, color=INK, font="Inter", max_width=None, leading=None):
    c.setFont(font, size)
    c.setFillColor(color)
    if max_width is None:
        c.drawString(x, y, value)
        return y
    leading = leading or size * 1.35
    words = value.split()
    lines = []
    line = ""
    for word in words:
        candidate = f"{line} {word}".strip()
        if pdfmetrics.stringWidth(candidate, font, size) <= max_width:
            line = candidate
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    for index, line_value in enumerate(lines):
        c.drawString(x, y - index * leading, line_value)
    return y - len(lines) * leading


def page_header(c, section, page):
    c.setFillColor(PAPER)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    if ISOTYPE.exists():
        c.drawImage(str(ISOTYPE), MARGIN, PAGE_H - 58, 25, 24, preserveAspectRatio=True, mask="auto")
    text(c, "GLP", MARGIN + 33, PAGE_H - 48, 12, DARK_GREEN, "Inter-Bold")
    c.setFont("Inter-Bold", 8)
    c.setFillColor(MUTED)
    c.drawRightString(PAGE_W - MARGIN, PAGE_H - 48, section.upper())
    c.setStrokeColor(PALE_BLUE)
    c.line(MARGIN, PAGE_H - 68, PAGE_W - MARGIN, PAGE_H - 68)
    text(c, f"Guía para usuarios de Excel  |  v0.2  |  {page}/6", MARGIN, 25, 8, MUTED)


def title(c, eyebrow, heading, description=None):
    text(c, eyebrow.upper(), MARGIN, PAGE_H - 102, 8, GREEN, "Inter-Bold")
    text(c, heading, MARGIN, PAGE_H - 133, 23, DARK_GREEN, "Montserrat")
    if description:
        text(c, description, MARGIN, PAGE_H - 158, 10.2, MUTED, max_width=PAGE_W - 2 * MARGIN, leading=14)


def number_circle(c, number, x, y, fill=GREEN):
    c.setFillColor(fill)
    c.circle(x, y, 13, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Inter-Bold", 10)
    c.drawCentredString(x, y - 3.5, str(number))


def bullet(c, value, x, y, width, color=INK, dot=GREEN, size=9.5):
    c.setFillColor(dot)
    c.circle(x + 4, y + 3, 2.2, fill=1, stroke=0)
    return text(c, value, x + 15, y + 7, size, color, max_width=width - 15, leading=13)


def code_box(c, lines, x, y, w, h):
    rounded(c, x, y, w, h, DARK_GREEN, radius=10)
    line_y = y + h - 21
    for line in lines:
        text(c, line, x + 16, line_y, 9.2, LIME, "Inter")
        line_y -= 17


def step_card(c, number, heading, body, x, y, w, h, accent=GREEN):
    rounded(c, x, y, w, h, WHITE, radius=12, stroke=PALE_BLUE)
    number_circle(c, number, x + 26, y + h - 27, accent)
    text(c, heading, x + 49, y + h - 23, 11, DARK_GREEN, "Inter-Bold")
    text(c, body, x + 18, y + h - 52, 9.1, MUTED, max_width=w - 36, leading=12.5)


def cover(c):
    c.setFillColor(DARK_GREEN)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFillColor(GREEN)
    c.circle(PAGE_W + 35, PAGE_H - 95, 170, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.circle(-40, 60, 155, fill=1, stroke=0)
    rounded(c, MARGIN, PAGE_H - 118, 132, 25, LIME, radius=12)
    text(c, "GUÍA OPERATIVA  |  v0.2", MARGIN + 13, PAGE_H - 109, 8.5, DARK_GREEN, "Inter-Bold")
    text(c, "GLP", MARGIN, PAGE_H - 235, 47, WHITE, "Montserrat")
    text(c, "Capítulo IV", MARGIN, PAGE_H - 278, 28, LIME, "Montserrat")
    text(c, "Instalación en Excel, configuración y uso", MARGIN, PAGE_H - 316, 15, WHITE, "Inter-Bold")
    text(
        c,
        "Una guía breve para instalar GLP en Excel, generar el modelo por área y resolver los problemas más comunes.",
        MARGIN,
        PAGE_H - 350,
        11,
        PALE_BLUE,
        max_width=410,
        leading=16,
    )
    rounded(c, MARGIN, 125, PAGE_W - 2 * MARGIN, 116, HexColor("#40563A"), radius=16)
    text(c, "VALIDACIÓN", MARGIN + 22, 214, 8, LIME, "Inter-Bold")
    text(c, "Instalación simple para usuarios de Excel", MARGIN + 22, 184, 14, WHITE, "Inter-Bold")
    text(c, "Instalación guiada con glp-installer.zip", MARGIN + 22, 158, 10, PALE_BLUE)
    text(c, "Actualizado: 19 de agosto de 2026", MARGIN + 22, 139, 9, PALE_BLUE)
    if LOGO.exists():
        c.drawImage(str(LOGO), PAGE_W - 180, 45, 138, 57, preserveAspectRatio=True, mask="auto")
    text(c, "Uso interno", MARGIN, 57, 8.5, PALE_BLUE)


def launch_page(c):
    page_header(c, "Instalación", 2)
    title(c, "Paso a paso", "Instalar GLP en Excel", "El usuario sólo necesita Excel de escritorio, conexión a internet y el archivo glp-installer.zip.")
    rounded(c, MARGIN, 552, PAGE_W - 2 * MARGIN, 88, CREAM, radius=14)
    text(c, "Antes de empezar", MARGIN + 18, 617, 11, DARK_GREEN, "Inter-Bold")
    bullet(c, "Windows 10/11 y Excel de escritorio de Microsoft 365.", MARGIN + 18, 591, 440)
    bullet(c, "Conexión a internet y Excel completamente cerrado durante la instalación.", MARGIN + 18, 567, 440)
    step_card(c, 1, "Descargar", "Descargá glp-installer.zip desde el enlace interno o repositorio indicado por Quintana Energy.", MARGIN, 426, 238, 96, BLUE)
    step_card(c, 2, "Extraer", "Hacé clic derecho sobre el ZIP, elegí Extraer todo y abrí la carpeta resultante.", MARGIN + 258, 426, 238, 96, GREEN)
    step_card(c, 3, "Instalar", "Con Excel cerrado, ejecutá instalar.bat. Esperá el mensaje LISTO - Instalación completada.", MARGIN, 307, 238, 96, DARK_GREEN)
    step_card(c, 4, "Abrir Excel", "Abrí un libro nuevo y entrá en Inicio > Complementos > Más complementos.", MARGIN + 258, 307, 238, 96, BLUE)
    step_card(c, 5, "Activar GLP", "En Complementos de desarrollador, seleccioná GLP. El panel se abrirá a la derecha.", MARGIN, 188, 238, 96, GREEN)
    step_card(c, 6, "Esperar el catálogo", "La primera carga puede demorar unos segundos. Cuando aparezcan las áreas, ya podés trabajar.", MARGIN + 258, 188, 238, 96, DARK_GREEN)
    rounded(c, MARGIN, 98, PAGE_W - 2 * MARGIN, 64, PALE_BLUE, radius=10)
    text(c, "Para desinstalar", MARGIN + 14, 138, 9, BLUE, "Inter-Bold")
    text(c, "Cerrá Excel y ejecutá desinstalar.bat desde la misma carpeta extraída.", MARGIN + 14, 116, 9.3, BLUE)


def configure_page(c):
    page_header(c, "Configuración", 3)
    title(c, "Dentro del panel", "Elegir áreas y supuestos", "El panel está organizado en tres pasos. Los cambios globales se aplican a todas las áreas salvo que se defina una excepción.")
    step_card(c, 1, "Filtrar y seleccionar", "Usá provincia, texto libre o empresa. Revisá la lista y elegí una o varias áreas. La selección masiva toma sólo el resultado filtrado.", MARGIN, 535, PAGE_W - 2 * MARGIN, 90, BLUE)
    step_card(c, 2, "Definir el horizonte", "Indicá año inicial y cantidad de años. Para una prueba rápida conviene usar el año corriente y un horizonte corto.", MARGIN, 425, PAGE_W - 2 * MARGIN, 90, GREEN)
    step_card(c, 3, "Elegir métodos", "Configurá producción bruta, petróleo, gas y pozos. Los parámetros quedan visibles y editables en las hojas de pronóstico.", MARGIN, 315, PAGE_W - 2 * MARGIN, 90, DARK_GREEN)
    rounded(c, MARGIN, 158, PAGE_W - 2 * MARGIN, 132, WHITE, radius=14, stroke=PALE_BLUE)
    text(c, "AJUSTES POR ÁREA", MARGIN + 18, 264, 8, GREEN, "Inter-Bold")
    text(c, "Sobrescribir sólo cuando haga falta", MARGIN + 18, 238, 13, DARK_GREEN, "Inter-Bold")
    bullet(c, "Año inicial diferente para una concesión concreta.", MARGIN + 18, 210, 450)
    bullet(c, "Método de pronóstico específico por corriente.", MARGIN + 18, 185, 450)
    bullet(c, "Producción inicial tomada del histórico o cargada manualmente.", MARGIN + 18, 160, 450)
    rounded(c, MARGIN, 88, PAGE_W - 2 * MARGIN, 62, LIME, radius=12)
    text(c, "Consejo", MARGIN + 16, 126, 9, DARK_GREEN, "Inter-Bold")
    text(c, "Si no activás una excepción, el área sigue cualquier cambio posterior hecho en los valores globales.", MARGIN + 16, 105, 9.4, DARK_GREEN, max_width=460, leading=13)


def generate_page(c):
    page_header(c, "Actualización", 4)
    title(c, "Mes a mes", "Actualizar un libro existente", "No hace falta volver a elegir áreas ni recordar la configuración usada anteriormente.")
    rounded(c, MARGIN, 520, PAGE_W - 2 * MARGIN, 112, LIME, radius=14)
    text(c, "ACTUALIZAR LIBRO", MARGIN + 17, 606, 8, DARK_GREEN, "Inter-Bold")
    text(c, "Traé automáticamente los meses nuevos", MARGIN + 17, 579, 14, DARK_GREEN, "Inter-Bold")
    text(c, "GLP lee el estado guardado dentro del Excel, detecta las áreas existentes y refresca la información oficial sin borrar tus supuestos.", MARGIN + 17, 552, 9.3, DARK_GREEN, max_width=455, leading=13)
    text(c, "Cómo hacerlo", MARGIN, 485, 12, DARK_GREEN, "Inter-Bold")
    checks = [
        ("Abrí el libro correcto", "Usá el mismo archivo de Excel donde GLP generó las áreas anteriormente."),
        ("Pulsá Actualizar libro", "El botón aparece al comienzo del panel, antes de la selección de áreas."),
        ("Esperá la descarga", "GLP revisa toda la serie para incorporar meses nuevos y correcciones oficiales, sin duplicar datos."),
        ("Confirmá el resultado", "El panel informa cuántas áreas se actualizaron y vuelve a calcular resumen y gráficos."),
    ]
    y = 445
    for index, (heading, body) in enumerate(checks, 1):
        number_circle(c, index, MARGIN + 14, y + 5, BLUE if index % 2 else GREEN)
        text(c, heading, MARGIN + 38, y + 10, 10.2, DARK_GREEN, "Inter-Bold")
        text(c, body, MARGIN + 38, y - 8, 9.1, MUTED, max_width=440, leading=12)
        y -= 66
    rounded(c, MARGIN, 120, PAGE_W - 2 * MARGIN, 63, PALE_BLUE, radius=10)
    text(c, "Tus ajustes se conservan", MARGIN + 15, 158, 9.1, BLUE, "Inter-Bold")
    text(c, "Los métodos y valores editados en las hojas Prono y Pozos permanecen. Regenerar áreas, en cambio, reconstruye todo desde cero.", MARGIN + 15, 137, 9, BLUE, max_width=465, leading=12.5)
    text(c, "Si el libro aún no tiene áreas creadas por GLP, usá el flujo normal de selección y generación.", MARGIN, 88, 8.8, MUTED)


def outputs_page(c):
    page_header(c, "Resultados", 5)
    title(c, "Libro generado", "Qué contiene cada hoja", "Las hojas por área usan un nombre corto y estable. El consolidado referencia sus pronósticos mediante fórmulas de Excel.")
    rows = [
        ("{AREA}_HDP", "Histórico mensual oficial agregado por área."),
        ("{AREA}_Prono", "Pronóstico de producción y supuestos editables."),
        ("{AREA}_Pozos", "Actividad y proyección de pozos."),
        ("{AREA}_Graficos", "Visuales de histórico y pronóstico."),
        ("{AREA}_Detalle", "Detalle de producción por pozo y mes."),
        ("Resumen_Areas", "Consolidado dinámico de las áreas seleccionadas."),
        ("CapIV_Descarga", "Recursos, fechas, estado y cantidad de filas."),
        ("CapIV_Debug", "Log de diagnóstico visible."),
        ("_CapIV_State", "Configuración interna; permanece oculta."),
    ]
    y = 615
    for index, (sheet, purpose) in enumerate(rows):
        fill = WHITE if index % 2 == 0 else HexColor("#F0F3ED")
        rounded(c, MARGIN, y - 42, PAGE_W - 2 * MARGIN, 42, fill, radius=7)
        text(c, sheet, MARGIN + 13, y - 26, 9.2, DARK_GREEN, "Inter-Bold")
        text(c, purpose, MARGIN + 164, y - 26, 9, MUTED, max_width=325)
        y -= 47
    rounded(c, MARGIN, 111, PAGE_W - 2 * MARGIN, 68, LIME, radius=12)
    text(c, "Edición segura", MARGIN + 16, 155, 9, DARK_GREEN, "Inter-Bold")
    text(c, "Modificá supuestos en las celdas destacadas de _Prono y _Pozos. Evitá cambiar nombres de hojas o columnas estructurales.", MARGIN + 16, 133, 9.4, DARK_GREEN, max_width=460, leading=13)


def troubleshooting_page(c):
    page_header(c, "Soporte", 6)
    title(c, "Antes de pedir ayuda", "Problemas frecuentes", "Estos controles resuelven la mayoría de los problemas de instalación y uso sin herramientas técnicas.")
    items = [
        ("GLP no aparece", "Cerrá Excel completamente, ejecutá instalar.bat otra vez y volvé a abrir un libro nuevo."),
        ("Windows bloquea el archivo", "Extraé el ZIP antes de ejecutar. Si proviene del enlace oficial interno, usá Propiedades > Desbloquear."),
        ("Falla una descarga", "Confirmá la conexión a internet y repetí con una sola área y un período corto. Revisá CapIV_Debug."),
        ("El resultado no cambia", "Verificá que el cálculo de Excel esté en Automático y que hayas editado una celda de supuesto."),
    ]
    y = 622
    for index, (heading, body) in enumerate(items, 1):
        number_circle(c, index, MARGIN + 14, y - 10, GREEN)
        text(c, heading, MARGIN + 38, y - 4, 10.5, DARK_GREEN, "Inter-Bold")
        text(c, body, MARGIN + 38, y - 24, 9, MUTED, max_width=440, leading=12.5)
        y -= 82
    rounded(c, MARGIN, 212, PAGE_W - 2 * MARGIN, 109, CREAM, radius=14)
    text(c, "IMPORTANTE", MARGIN + 16, 294, 8, GREEN, "Inter-Bold")
    text(c, "Instalación directa en Excel", MARGIN + 16, 268, 13, DARK_GREEN, "Inter-Bold")
    text(c, "La instalación se realiza con instalar.bat y la activación desde el menú Complementos de Excel.", MARGIN + 16, 245, 9.2, MUTED, max_width=460, leading=13)
    rounded(c, MARGIN, 108, PAGE_W - 2 * MARGIN, 79, DARK_GREEN, radius=13)
    text(c, "LISTO PARA USAR", MARGIN + 16, 162, 8, LIME, "Inter-Bold")
    text(c, "Excel de escritorio  |  Internet  |  GLP visible", MARGIN + 16, 139, 10.5, WHITE, "Inter-Bold")
    text(c, "Si las áreas aparecen en el panel, la instalación terminó correctamente.", MARGIN + 16, 120, 9.3, PALE_BLUE)
    text(c, "Ante un error persistente, enviá una captura del panel y de CapIV_Debug al equipo de soporte.", MARGIN, 83, 8.5, MUTED)


def build():
    register_fonts()
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUTPUT), pagesize=A4, pageCompression=1)
    c.setTitle("GLP - Guía de instalación para usuarios de Excel")
    c.setAuthor("Quintana Energy")
    c.setSubject("Tutorial para instalar y utilizar el add-in GLP de Capítulo IV en Excel")
    for draw_page in [cover, launch_page, configure_page, generate_page, outputs_page, troubleshooting_page]:
        draw_page(c)
        c.showPage()
    c.save()
    print(OUTPUT)


if __name__ == "__main__":
    build()
