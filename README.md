# Convertisseur Word/PDF → LaTeX

Application de bureau Windows, **hors ligne**, qui convertit des fichiers
`.docx` ou `.pdf` en documents `.tex` prêts à compiler : titres, gras/italique,
listes, images, et correction orthographique/grammaticale en français.

## 1. Installation (sur ta machine, une seule fois)

Il te faut Python 3.10+ installé (https://www.python.org/downloads/,
coche "Add Python to PATH" pendant l'installation).

Ouvre un terminal (PowerShell) dans le dossier du projet :

```
pip install -r requirements.txt
```

### Correction orthographique/grammaticale (optionnelle mais recommandée)
`language_tool_python` s'appuie sur **LanguageTool**, qui a besoin de **Java**
(Java 17+ recommandé : https://adoptium.net/). Au tout premier lancement de
l'appli avec la case "correction" cochée, LanguageTool télécharge son moteur
local (~200 Mo, une seule fois — connexion internet nécessaire ce jour-là).
Ensuite, il tourne 100% en local, donc l'appli reste utilisable hors ligne.

Si tu ne veux pas de cette dépendance (pas de Java, pas envie du
téléchargement), tu peux simplement laisser la case décochée dans l'appli :
la conversion fonctionne quand même, juste sans correction.

## 2. Lancer l'application

```
python main.py
```

Une fenêtre s'ouvre : choisis ton fichier `.docx`/`.pdf`, le dossier de
sortie, coche ou non la correction, puis clique sur "Convertir en LaTeX".
Le fichier `.tex` et un sous-dossier `images/` sont générés dans le dossier
choisi.

## 3. Compiler le `.tex` généré en PDF

Il te faut une distribution LaTeX installée (MiKTeX pour Windows est la plus
simple : https://miktex.org/, elle s'installe une fois et fonctionne ensuite
hors ligne). Ensuite, dans le dossier de sortie :

```
pdflatex fichier.tex
```

## 4. Créer un .exe autonome

### Option A — en local
Une fois que tout fonctionne avec `python main.py`, tu peux packager
l'appli en un seul `.exe` avec PyInstaller (déjà dans requirements.txt) :

```
pyinstaller --noconsole --onefile --name "ConvertisseurLatex" main.py
```

L'exécutable sera dans `dist/ConvertisseurLatex.exe`.

### Option B — via GitHub Actions (comme Gestion Scolaire Congo)
Le workflow `.github/workflows/build.yml` est déjà inclus dans ce projet.
Il suffit de :

1. Créer un dépôt GitHub et y pousser ce dossier (`git init`, `git add .`,
   `git commit`, `git push`)
2. GitHub construit automatiquement le `.exe` sur une machine Windows à
   chaque `push` sur `main` (ou manuellement via l'onglet "Actions" →
   "Run workflow")
3. Une fois le build terminé, l'exe est téléchargeable dans l'onglet
   "Actions" → le run correspondant → section "Artifacts" →
   `ConvertisseurLatex-windows`

⚠️ Cette méthode vérifie que le code s'importe et que PyInstaller construit
sans erreur, mais ne teste pas visuellement l'interface (pas d'écran sur
les machines GitHub). Un premier test manuel avec `python main.py` reste
recommandé avant de packager.

⚠️ Important : `--onefile` n'embarque pas Java ni le moteur LanguageTool.
Si tu distribues l'exe à d'autres personnes, elles devront aussi avoir Java
installé pour que la correction orthographique fonctionne (ou tu distribues
l'appli sans cette fonctionnalité activée par défaut).

## 5. Limites connues (première version)

- PDF : la mise en page est reconstruite par heuristique (taille de police
  = titres), donc moins fidèle qu'un `.docx` natif, qui a une vraie structure.
- Tableaux Word : convertis mais insérés en fin de document (pas encore à
  leur position exacte dans le texte).
- Notes de bas de page, en-têtes/pieds de page, et mise en forme très
  avancée (colonnes, zones de texte) ne sont pas encore gérés.

Dis-moi lesquelles de ces limites te gênent le plus dans ton usage réel
(cours SVT, dossiers de candidature...) et je les corrige en priorité.
