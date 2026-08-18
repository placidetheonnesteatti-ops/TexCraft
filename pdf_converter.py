# -*- coding: utf-8 -*-
"""
Convertisseur PDF -> LaTeX (.tex)

Le PDF n'a pas de structure logique (contrairement au .docx), donc la
conversion est basée sur des heuristiques :
 - la taille de police la plus fréquente = corps de texte
 - les lignes en police nettement plus grande = titres (section/subsection)
 - les images sont extraites page par page et insérées à leur position
   approximative dans le flux
"""

import os
import statistics
from collections import defaultdict

import pdfplumber

from text_utils import escape_latex

LATEX_PREAMBLE = r"""\documentclass[a4paper,11pt]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[french]{babel}
\usepackage{graphicx}
\usepackage{geometry}
\geometry{margin=2.5cm}

\title{%s}
\author{}
\date{}

\begin{document}
\maketitle

"""

LATEX_END = "\n\\end{document}\n"


def _group_lines_by_position(words):
    """Regroupe les mots pdfplumber en lignes selon leur position verticale (top)."""
    lines = defaultdict(list)
    for w in words:
        key = round(w["top"] / 2) * 2  # tolérance de 2pt
        lines[key].append(w)
    ordered_keys = sorted(lines.keys())
    result = []
    for k in ordered_keys:
        ws = sorted(lines[k], key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in ws)
        avg_size = statistics.mean(w["size"] for w in ws) if ws else 0
        result.append((text, avg_size, k))
    return result


def convert_pdf_to_latex(input_path: str, output_dir: str, title: str = "",
                          checker=None, progress_cb=None) -> str:
    """
    Convertit un fichier .pdf en .tex (best-effort : mise en page approximative).
    Retourne le chemin du fichier .tex généré.
    """
    os.makedirs(output_dir, exist_ok=True)
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    if not title:
        title = os.path.splitext(os.path.basename(input_path))[0]

    body_parts = [LATEX_PREAMBLE % escape_latex(title)]
    img_counter = 0

    with pdfplumber.open(input_path) as pdf:
        total = len(pdf.pages)

        # 1ère passe : déterminer la taille de police "normale" (la plus fréquente)
        all_sizes = []
        for page in pdf.pages:
            for w in page.extract_words(extra_attrs=["size"]):
                all_sizes.append(round(w["size"]))
        body_size = statistics.mode(all_sizes) if all_sizes else 11

        for page_idx, page in enumerate(pdf.pages):
            if progress_cb:
                progress_cb(page_idx + 1, total, f"Page {page_idx + 1}/{total}...")

            words = page.extract_words(extra_attrs=["size"])
            lines = _group_lines_by_position(words)

            for text, size, _top in lines:
                text = text.strip()
                if not text:
                    continue
                escaped = escape_latex(text)
                if checker is not None:
                    escaped = checker.correct(escaped)

                if size >= body_size + 6:
                    body_parts.append(f"\\section{{{escaped}}}\n")
                elif size >= body_size + 3:
                    body_parts.append(f"\\subsection{{{escaped}}}\n")
                else:
                    body_parts.append(escaped + "\n")

            # Images de la page
            for img in page.images:
                try:
                    cropped = page.crop((img["x0"], img["top"], img["x1"], img["bottom"]))
                    im = cropped.to_image(resolution=150)
                    img_counter += 1
                    filename = f"image_{img_counter}.png"
                    im.save(os.path.join(images_dir, filename))
                    body_parts.append(
                        "\\begin{figure}[h]\n\\centering\n"
                        f"\\includegraphics[width=0.7\\textwidth]{{images/{filename}}}\n"
                        "\\end{figure}\n"
                    )
                except Exception:
                    # certaines images vectorielles/masquées ne se recadrent pas proprement
                    continue

            body_parts.append("\n")  # séparation entre pages

    body_parts.append(LATEX_END)

    tex_content = "\n".join(body_parts)
    out_path = os.path.join(output_dir, os.path.splitext(os.path.basename(input_path))[0] + ".tex")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(tex_content)

    return out_path
