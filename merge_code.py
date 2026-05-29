import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from pygments import lex
from pygments.lexers import get_lexer_by_name, CSharpLexer
from pygments.token import Token
from pygments.util import ClassNotFound

# =========================
# הגדרות
# =========================

project_path = r"C:\Users\aradl\source\repos\TheWorldOfHorses🐎"
output_file = "all_pages_documentation.docx"
extensions_to_find = [".aspx", ".cs"]


# =========================
# קיבוץ דפים
# =========================

pages = {}


# =========================
# צבעים
# =========================

COLOR_MAP = {
    Token.Keyword: RGBColor(0, 0, 255),
    Token.Keyword.Type: RGBColor(43, 145, 175),
    Token.Name.Class: RGBColor(43, 145, 175),
    Token.Name.Function: RGBColor(124, 91, 0),
    Token.Name.Variable: RGBColor(100, 40, 200),
    Token.String: RGBColor(163, 21, 21),
    Token.Number: RGBColor(9, 134, 115),
    Token.Operator: RGBColor(255, 69, 0),
    Token.Punctuation: RGBColor(139, 0, 139),
    Token.Comment: RGBColor(0, 128, 0),
}


# =========================
# סריקת קבצים + קיבוץ
# =========================

print("סורק קבצים ומקבץ דפים...")

for root, dirs, files in os.walk(project_path):

    if any(x in root for x in [".git", ".vs", "bin", "obj", "Properties"]):
        continue

    for file in files:

        if file.endswith((".aspx", ".cs")):

            file_path = os.path.join(root, file)

            base_name = (
                file.split(".aspx")[0] if ".aspx" in file else file.split(".cs")[0]
            )

            if base_name not in pages:
                pages[base_name] = {"aspx": "", "cs": "", "designer": ""}

            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
            except:
                with open(file_path, "r", encoding="windows-1255") as f:
                    content = f.read()

            if file.endswith(".aspx") and not file.endswith(".aspx.cs"):
                pages[base_name]["aspx"] = content

            elif file.endswith(".aspx.cs") and "designer" not in file:
                pages[base_name]["cs"] = content

            elif "designer" in file:
                pages[base_name]["designer"] = content


# =========================
# יצירת מסמך
# =========================

doc = Document()

for section in doc.sections:
    section.top_margin = Pt(72)
    section.bottom_margin = Pt(72)
    section.left_margin = Pt(72)
    section.right_margin = Pt(72)


# =========================
# כותרת ראשית
# =========================

title = doc.add_heading("תיעוד קוד הפרויקט - TheWorldOfHorses", level=0)
title.alignment = WD_ALIGN_PARAGRAPH.RIGHT


# =========================
# תוכן עניינים (TOC)
# =========================

doc.add_paragraph("תוכן עניינים")
toc = doc.add_paragraph()

fld = OxmlElement("w:fldSimple")
fld.set(qn("w:instr"), 'TOC \\o "1-2" \\h \\z \\u')
toc._element.append(fld)

doc.add_page_break()


# =========================
# פונקציית צביעת קוד
# =========================


def add_code(text, file_type):

    if file_type == "cs":
        lexer = CSharpLexer()
    else:
        try:
            lexer = get_lexer_by_name("html+aspx")
        except ClassNotFound:
            lexer = get_lexer_by_name("html")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    tokens = lex(text, lexer)

    for token_type, value in tokens:

        run = p.add_run(value)
        run.font.name = "Courier New"
        run.font.size = Pt(9)

        color = None
        for base, c in COLOR_MAP.items():
            if token_type in base:
                color = c
                break

        run.font.color.rgb = color if color else RGBColor(40, 40, 40)


# =========================
# כתיבת דפים (בלי כפילויות!)
# =========================

for page_name, data in pages.items():

    doc.add_heading(f"דף: {page_name}", level=1)

    # ASPX
    if data["aspx"]:
        doc.add_heading("ASPX / HTML", level=2)
        add_code(data["aspx"], "aspx")

    # C#
    if data["cs"]:
        doc.add_heading("C# Code Behind", level=2)
        add_code(data["cs"], "cs")

    # Designer
    if data["designer"]:
        doc.add_heading("Designer", level=2)
        add_code(data["designer"], "cs")

    doc.add_page_break()


# =========================
# שמירה
# =========================

doc.save(output_file)

print("סיום! נוצר קובץ מסודר בלי כפילויות:", output_file)
