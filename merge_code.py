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

# הגדרות נתיבים
project_path = r"C:\Users\aradl\source\repos\TheWorldOfHorses🐎"
output_file = "all_pages_documentation.docx"
extensions_to_find = [".aspx", ".cs"]

# מפת צבעים מורחבת ומלאה - סגנון VS Code Dark / Light משולב למקסימום צבעים!
COLOR_MAP = {
    # מילים שמורות וסוגי נתונים
    Token.Keyword: RGBColor(0, 0, 255),  # כחול מודגש (if, public, class, return)
    Token.Keyword.Type: RGBColor(
        43, 145, 175
    ),  # כחול-טורקיז לסוגי נתונים (int, string, void)
    # שמות, מחלקות ופונקציות
    Token.Name.Class: RGBColor(43, 145, 175),  # טורקיז לשמות של מחלקות (Page, Response)
    Token.Name.Function: RGBColor(
        124, 91, 0
    ),  # חום-זהב לפונקציות ומתודות (Page_Load, Button_Click)
    Token.Name.Namespace: RGBColor(0, 100, 0),  # ירוק כהה ל-using ושמות מרחבי שמות
    Token.Name.Variable: RGBColor(100, 40, 200),  # סגול לשמות משתנים ייחודיים
    # טקסט, מספרים וסימנים
    Token.String: RGBColor(163, 21, 21),  # אדום כהה למחרוזות טקסט בגרשיים
    Token.Number: RGBColor(9, 134, 115),  # ירוק-כחול זוהר למספרים
    Token.Operator: RGBColor(
        255, 69, 0
    ),  # כתום/אדום לאופרטורים וסימנים (=, +, ==, &&, !)
    Token.Punctuation: RGBColor(
        139, 0, 139
    ),  # סגול כהה לסוגריים ונקודה-פסיק ({, }, [, ], ;, .)
    Token.Comment: RGBColor(0, 128, 0),  # ירוק בהיר להערות קוד (// או ו-/* */)
    # תגיות HTML ו-ASP.NET (צד לקוח)
    Token.Name.Tag: RGBColor(
        163, 21, 21
    ),  # אדום/בורדו לתגיות HTML רגילות (div, table, h1)
    Token.Name.Attribute: RGBColor(
        255, 0, 0
    ),  # אדום מאורר למאפיינים (id, class, style)
    Token.Comment.Preproc: RGBColor(
        128, 0, 128
    ),  # סגול עמוק לתגיות אחוז של ASPX (<%@ Page ... %>)
}


def set_cell_background(cell, color_hex):
    """צביעת רקע התיבה באפור-קוד מקצועי ונקסטרני"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def add_colored_code_to_cell(cell, code_text, file_ext):
    """מפרקת את הקוד לצבעים מפורטים ושומרת על עימוד"""
    if file_ext == ".cs":
        lexer = CSharpLexer()
    else:
        try:
            lexer = get_lexer_by_name("html+aspx")
        except ClassNotFound:
            try:
                lexer = get_lexer_by_name("aspx")
            except ClassNotFound:
                lexer = get_lexer_by_name("html")

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(0)

    tokens = list(lex(code_text, lexer))

    line_number = 1
    num_run = p.add_run(f"{line_number:3}  ")
    num_run.font.name = "Courier New"
    num_run.font.size = Pt(9)
    num_run.font.color.rgb = RGBColor(150, 150, 150)

    for token_type, value in tokens:
        if "\n" in value:
            parts = value.split("\n")
            for i, part in enumerate(parts):
                if i > 0:
                    line_number += 1
                    p = cell.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(0)

                    num_run = p.add_run(f"{line_number:3}  ")
                    num_run.font.name = "Courier New"
                    num_run.font.size = Pt(9)
                    num_run.font.color.rgb = RGBColor(150, 150, 150)

                if part:
                    run = p.add_run(part)
                    run.font.name = "Courier New"
                    run.font.size = Pt(9.5)

                    matched_color = None
                    for base_type, color in COLOR_MAP.items():
                        if token_type in base_type:
                            matched_color = color
                            break
                    run.font.color.rgb = (
                        matched_color if matched_color else RGBColor(40, 40, 40)
                    )
        else:
            if value:
                run = p.add_run(value)
                run.font.name = "Courier New"
                run.font.size = Pt(9.5)

                matched_color = None
                for base_type, color in COLOR_MAP.items():
                    if token_type in base_type:
                        matched_color = color
                        break
                run.font.color.rgb = (
                    matched_color if matched_color else RGBColor(40, 40, 40)
                )


print("מפעיל את ה-Extension הפיקסלי לצביעת קוד... אנא המתן...")

doc = Document()

# הגדרת שוליים
for section in doc.sections:
    section.top_margin = Pt(72)
    section.bottom_margin = Pt(72)
    section.left_margin = Pt(72)
    section.right_margin = Pt(72)

# כותרת ראשית
main_title = doc.add_heading("תיעוד קוד הפרויקט - TheWorldOfHorses", level=0)
main_title.alignment = WD_ALIGN_PARAGRAPH.RIGHT

# סריקת הקבצים
for root, dirs, files in os.walk(project_path):
    if any(part in root for part in [".git", ".vs", "bin", "obj", "Properties"]):
        continue

    for file in files:
        if any(file.endswith(ext) for ext in extensions_to_find):
            file_path = os.path.join(root, file)
            _, ext = os.path.splitext(file)

            # 1. כותרת הדף
            heading = doc.add_heading(f"שם הדף: {file}", level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # 2. מבנה עברית ואפיון
            if file.endswith(".aspx") and not file.endswith(".aspx.cs"):
                p1 = doc.add_paragraph("[כאן יש להדביק צילום מסך של הדף מהדפדפן]")
                p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                p2 = doc.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p2.add_run("פירוט תפקיד הדף (למה הוא משמש, מה מוצג בו):\n").bold = True
                p2.add_run("תשובה: ____________\n")

                p3 = doc.add_paragraph()
                p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p3.add_run("קוד צד לקוח (HTML / JavaScript):").bold = True

            elif file.endswith(".aspx.cs"):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.add_run("קוד מאחור (Code Behind - C#):").bold = True

            # 3. קריאת תוכן הקובץ
            file_content = ""
            try:
                with open(file_path, "r", encoding="utf-8") as infile:
                    file_content = infile.read()
            except UnicodeDecodeError:
                try:
                    with open(file_path, "r", encoding="windows-1255") as infile:
                        file_content = infile.read()
                except Exception as e:
                    file_content = f"[שגיאה בקריאת הקובץ: {e}]"

            # 4. יצירת הבלוק המעוצב
            table = doc.add_table(rows=1, cols=1)
            table.autofit = False
            cell = table.cell(0, 0)

            # רקע אפור בהיר ונקי (F8F9FA) כדי שהצבעים יבלטו בצורה מקסימלית
            set_cell_background(cell, "F8F9FA")

            add_colored_code_to_cell(cell, file_content, ext)

            # מעבר עמוד נקי
            doc.add_page_break()

# שמירה
doc.save(output_file)
print(f"\nהסתיים בהצלחה! הקוד נצבע בצורה מלאה ומחכה לך: {output_file}")
