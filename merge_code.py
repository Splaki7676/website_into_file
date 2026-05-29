import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pygments import lex
from pygments.lexers import get_lexer_by_name, CSharpLexer
from pygments.token import Token

# =========================
# טעינת מילון תיאורים
# =========================

try:
    from site_descriptions import PAGES_DESCRIPTION

    print("✔ מילון תיאורים נטען")
except:
    print("❌ לא נמצא מילון תיאורים")
    PAGES_DESCRIPTION = {}

# =========================
# הגדרות
# =========================

project_path = r"C:\Users\aradl\source\repos\TheWorldOfHorses🐎"
output_file = "all_pages_documentation.docx"

pages = {}

# =========================
# סריקה
# =========================

print("🔍 סורק קבצים...")

for root, _, files in os.walk(project_path):

    if any(skip in root for skip in [".git", ".vs", "bin", "obj"]):
        continue

    for file in files:

        if not (file.endswith(".aspx") or file.endswith(".cs")):
            continue

        path = os.path.join(root, file)

        # =========================
        # שם דף אחיד (הכי חשוב!)
        # =========================

        page_name = file
        page_name = page_name.replace(".aspx.cs", "")
        page_name = page_name.replace(".designer.cs", "")
        page_name = page_name.replace(".aspx", "")
        page_name = page_name.replace(".cs", "")

        if page_name not in pages:
            pages[page_name] = {"aspx": "", "cs": "", "designer": ""}

        # =========================
        # קריאת קובץ
        # =========================

        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
        except:
            with open(path, "r", encoding="windows-1255") as f:
                content = f.read()

        # =========================
        # שמירה לפי סוג
        # =========================

        if file.endswith(".aspx") and not file.endswith(".aspx.cs"):
            pages[page_name]["aspx"] = content

        elif file.endswith(".aspx.cs"):
            pages[page_name]["cs"] = content

        elif file.endswith(".designer.cs"):
            pages[page_name]["designer"] = content

# =========================
# יצירת Word
# =========================

doc = Document()

doc.add_heading("תיעוד פרויקט - TheWorldOfHorses", 0)

doc.add_page_break()

# =========================
# פונקציית קוד
# =========================


def add_code(text, lang):
    lexer = CSharpLexer() if lang == "cs" else get_lexer_by_name("html")

    p = doc.add_paragraph()

    for t, v in lex(text, lexer):

        run = p.add_run(v)
        run.font.name = "Courier New"
        run.font.size = Pt(9)

        if t in Token.Keyword:
            run.font.color.rgb = RGBColor(0, 0, 255)
        elif t in Token.String:
            run.font.color.rgb = RGBColor(163, 21, 21)
        elif t in Token.Comment:
            run.font.color.rgb = RGBColor(0, 128, 0)
        else:
            run.font.color.rgb = RGBColor(40, 40, 40)


# =========================
# כתיבה למסמך
# =========================

print("📄 כותב למסמך...")

for page_name, data in sorted(pages.items()):

    doc.add_heading(f"דף: {page_name}", 1)

    # =========================
    # 🔥 פתרון חכם למילון (זה החלק החשוב)
    # =========================

    desc = (
        PAGES_DESCRIPTION.get(page_name)
        or PAGES_DESCRIPTION.get(page_name + ".aspx")
        or PAGES_DESCRIPTION.get(page_name + ".cs")
        or "⚠ אין תיאור בדאטהבייס"
    )

    p = doc.add_paragraph(desc)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # =========================
    # ASPX
    # =========================

    if data["aspx"]:
        doc.add_heading("ASPX", 2)
        add_code(data["aspx"], "html")

    # =========================
    # C#
    # =========================

    if data["cs"]:
        doc.add_heading("C#", 2)
        add_code(data["cs"], "cs")

    # =========================
    # Designer
    # =========================

    if data["designer"]:
        doc.add_heading("Designer", 2)
        add_code(data["designer"], "cs")

    doc.add_page_break()

# =========================
# שמירה
# =========================

doc.save(output_file)

print("✅ סיום! הקובץ נוצר בהצלחה:", output_file)
