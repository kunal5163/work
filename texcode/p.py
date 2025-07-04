import os
import re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.shared import RGBColor

def escape_latex(text):
    return text.replace('\\', r'\textbackslash{}').replace('&', r'\&').replace('%', r'\%') \
               .replace('$', r'\$').replace('#', r'\#').replace('_', r'\_').replace('{', r'\{') \
               .replace('}', r'\}').replace('~', r'\textasciitilde{}').replace('^', r'\^{}')

def get_alignment_env(alignment):
    if alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return 'center'
    elif alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        return 'flushright'
    elif alignment == WD_ALIGN_PARAGRAPH.LEFT:
        return 'flushleft'
    return 'justify'

# Mapping Word highlight color to LaTeX-compatible colors
HIGHLIGHT_MAP = {
    WD_COLOR_INDEX.YELLOW: "yellow",
    WD_COLOR_INDEX.TURQUOISE: "cyan",
    WD_COLOR_INDEX.PINK: "pink",
    WD_COLOR_INDEX.GREEN: "lime",
    WD_COLOR_INDEX.BLUE: "cyan",
    WD_COLOR_INDEX.RED: "red",
    WD_COLOR_INDEX.GRAY_25: "lightgray",
    WD_COLOR_INDEX.DARK_RED: "red",
    WD_COLOR_INDEX.DARK_BLUE: "blue",
    WD_COLOR_INDEX.DARK_YELLOW: "orange"
    # WD_COLOR_INDEX.DARK_GREEN does not exist
}

def docx_to_latex(docx_path, tex_path="output.tex"):
    doc = Document(docx_path)

    latex_lines = [
        r"\documentclass[12pt]{article}",
        r"\usepackage[utf8]{inputenc}",
        r"\usepackage{xcolor}",
        r"\usepackage{soul}",  # for background highlight
        r"\usepackage{geometry}",
        r"\usepackage{setspace}",
        r"\usepackage{ulem}",  # for underline
        r"\usepackage{hyperref}",
        r"\usepackage{ragged2e}",
        r"\geometry{margin=1in}",
        r"\title{Converted Document}",
        r"\author{}",
        r"\date{}",
        r"\begin{document}",
        r"\maketitle",
        r"\noindent"
    ]

    for para in doc.paragraphs:
        alignment = get_alignment_env(para.alignment)
        line_parts = []

        for run in para.runs:
            text = escape_latex(run.text)
            if not text:
                continue

            # Font styling
            if run.bold:
                text = f"\\textbf{{{text}}}"
            if run.italic:
                text = f"\\textit{{{text}}}"
            if run.underline:
                text = f"\\uline{{{text}}}"

            # Font color
            if run.font.color and isinstance(run.font.color.rgb, RGBColor):
                hex_color = run.font.color.rgb.__str__()
                text = f"\\textcolor[HTML]{{{hex_color}}}{{{text}}}"

            # Highlight color (background)
            highlight = run.font.highlight_color
            if highlight is not None and highlight in HIGHLIGHT_MAP:
                latex_color = HIGHLIGHT_MAP[highlight]
                text = f"\\sethlcolor{{{latex_color}}}\\hl{{{text}}}"

            line_parts.append(text)

        full_line = ' '.join(line_parts)
        if full_line.strip():
            latex_lines.append(f"\\begin{{{alignment}}}\n{full_line}\n\\end{{{alignment}}}")
        else:
            latex_lines.append(r"\vspace{1em}")  # Blank paragraph

    latex_lines.append(r"\end{document}")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("\n".join(latex_lines))

    print(f"LaTeX file written to: {tex_path}")

# === Example Usage ===
if __name__ == "__main__":
    docx_to_latex("small.docx", "small_highlighted_output.tex")
