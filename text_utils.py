# -*- coding: utf-8 -*-
"""
Utilitaires de traitement de texte :
 - échappement des caractères spéciaux LaTeX
 - correction orthographique et grammaticale (via LanguageTool, hors ligne
   après le premier téléchargement du serveur local)
"""

import re

LATEX_SPECIAL_CHARS = {
    '&': r'\&',
    '%': r'\%',
    '$': r'\$',
    '#': r'\#',
    '_': r'\_',
    '{': r'\{',
    '}': r'\}',
    '~': r'\textasciitilde{}',
    '^': r'\textasciicircum{}',
    '\\': r'\textbackslash{}',
}


def escape_latex(text: str) -> str:
    """Échappe les caractères spéciaux LaTeX dans un texte brut."""
    if not text:
        return ""
    pattern = re.compile('|'.join(re.escape(k) for k in LATEX_SPECIAL_CHARS))
    return pattern.sub(lambda m: LATEX_SPECIAL_CHARS[m.group()], text)


class SpellGrammarChecker:
    """
    Wrapper autour de LanguageTool (language_tool_python).

    Nécessite Java installé sur la machine et le package
    `language_tool_python`. Au premier lancement, LanguageTool télécharge
    son serveur local (~200-250 Mo) ; ensuite il fonctionne entièrement
    hors ligne (le serveur tourne en local, pas d'appel réseau).
    """

    def __init__(self, lang: str = "fr"):
        self.lang = lang
        self.tool = None
        self.available = False
        try:
            import language_tool_python
            self.tool = language_tool_python.LanguageTool(lang)
            self.available = True
        except Exception as e:
            # LanguageTool non installé / Java absent / pas encore téléchargé
            self.available = False
            self._error = str(e)

    def correct(self, text: str) -> str:
        """Corrige les fautes évidentes (orthographe/grammaire) dans un texte."""
        if not self.available or not text.strip():
            return text
        try:
            matches = self.tool.check(text)
            # On applique uniquement les corrections avec une suggestion claire
            import language_tool_python
            return language_tool_python.utils.correct(text, matches)
        except Exception:
            return text

    def close(self):
        if self.tool is not None:
            try:
                self.tool.close()
            except Exception:
                pass
