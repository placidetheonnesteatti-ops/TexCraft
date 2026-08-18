# -*- coding: utf-8 -*-
"""
Convertisseur Word/PDF -> LaTeX — Application de bureau (hors ligne)

Lancement : python main.py
"""

import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx_converter import convert_docx_to_latex
from pdf_converter import convert_pdf_to_latex
from text_utils import SpellGrammarChecker


class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Convertisseur Word/PDF → LaTeX")
        self.geometry("620x420")
        self.resizable(False, False)

        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Documents", "LaTeX_export"))
        self.use_spellcheck = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        frm_in = tk.LabelFrame(self, text="1. Fichier à convertir (.docx ou .pdf)")
        frm_in.pack(fill="x", **pad)
        tk.Entry(frm_in, textvariable=self.input_path, width=60).pack(side="left", padx=8, pady=8)
        tk.Button(frm_in, text="Parcourir...", command=self._choose_input).pack(side="left", padx=8)

        frm_out = tk.LabelFrame(self, text="2. Dossier de sortie")
        frm_out.pack(fill="x", **pad)
        tk.Entry(frm_out, textvariable=self.output_dir, width=60).pack(side="left", padx=8, pady=8)
        tk.Button(frm_out, text="Choisir...", command=self._choose_output).pack(side="left", padx=8)

        frm_opts = tk.LabelFrame(self, text="3. Options")
        frm_opts.pack(fill="x", **pad)
        tk.Checkbutton(
            frm_opts,
            text="Corriger l'orthographe et la grammaire (français, via LanguageTool)",
            variable=self.use_spellcheck,
        ).pack(anchor="w", padx=8, pady=6)

        self.btn_convert = tk.Button(
            self, text="Convertir en LaTeX", command=self._start_conversion,
            bg="#2e7d32", fg="white", font=("Segoe UI", 11, "bold"), height=2
        )
        self.btn_convert.pack(fill="x", **pad)

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=12, pady=(0, 6))

        self.status_var = tk.StringVar(value="Prêt.")
        tk.Label(self, textvariable=self.status_var, anchor="w").pack(fill="x", padx=12)

        frm_log = tk.LabelFrame(self, text="Journal")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frm_log, height=8, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=4, pady=4)

    def _choose_input(self):
        path = filedialog.askopenfilename(
            filetypes=[("Documents Word/PDF", "*.docx *.pdf"), ("Tous les fichiers", "*.*")]
        )
        if path:
            self.input_path.set(path)

    def _choose_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def _log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_progress(self, current, total, msg):
        self.progress["maximum"] = total
        self.progress["value"] = current
        self.status_var.set(msg)
        self.update_idletasks()

    def _start_conversion(self):
        in_path = self.input_path.get().strip()
        out_dir = self.output_dir.get().strip()

        if not in_path or not os.path.isfile(in_path):
            messagebox.showerror("Erreur", "Sélectionne un fichier .docx ou .pdf valide.")
            return
        if not out_dir:
            messagebox.showerror("Erreur", "Choisis un dossier de sortie.")
            return

        self.btn_convert.configure(state="disabled", text="Conversion en cours...")
        self._log(f"Démarrage : {os.path.basename(in_path)}")
        threading.Thread(target=self._run_conversion, args=(in_path, out_dir), daemon=True).start()

    def _run_conversion(self, in_path, out_dir):
        checker = None
        try:
            if self.use_spellcheck.get():
                self._log("Chargement du correcteur orthographique/grammatical (français)...")
                checker = SpellGrammarChecker(lang="fr")
                if not checker.available:
                    self._log("⚠ Correcteur indisponible (LanguageTool/Java non installé) — conversion sans correction.")
                    checker = None

            ext = os.path.splitext(in_path)[1].lower()
            if ext == ".docx":
                out_path = convert_docx_to_latex(
                    in_path, out_dir, checker=checker, progress_cb=self._set_progress
                )
            elif ext == ".pdf":
                out_path = convert_pdf_to_latex(
                    in_path, out_dir, checker=checker, progress_cb=self._set_progress
                )
            else:
                raise ValueError("Format non supporté (utilise .docx ou .pdf)")

            self._log(f"✔ Terminé : {out_path}")
            self.status_var.set("Conversion terminée.")
            messagebox.showinfo("Succès", f"Fichier LaTeX généré :\n{out_path}")

        except Exception as e:
            self._log("✘ Erreur : " + str(e))
            self._log(traceback.format_exc())
            messagebox.showerror("Erreur", str(e))
        finally:
            if checker:
                checker.close()
            self.btn_convert.configure(state="normal", text="Convertir en LaTeX")


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
