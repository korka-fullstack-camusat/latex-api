# Guide de test — Word to LaTeX API

## Sommaire
1. [Démarrage rapide](#1-démarrage-rapide)
2. [Endpoints](#2-endpoints)
3. [Tests cURL](#3-tests-curl)
4. [Script de test automatisé](#4-script-de-test-automatisé)
5. [Cas d'erreur à tester](#5-cas-derreur-à-tester)
6. [Valider la réponse ZIP](#6-valider-la-réponse-zip)
7. [Checklist de validation](#7-checklist-de-validation)

---

## 1. Démarrage rapide

### Lancer le backend
```bash
# Installation
pip install fastapi uvicorn python-docx anthropic python-multipart

# Clé API
export ANTHROPIC_API_KEY="sk-ant-..."

# Démarrage
uvicorn main:app --reload --port 8000

# Vérification
curl http://localhost:8000/health
# → {"status":"ok"}
```

### Swagger UI intégré
```
http://localhost:8000/docs        ← interface graphique (recommandé)
http://localhost:8000/redoc       ← documentation alternative
http://localhost:8000/openapi.json
```

---

## 2. Endpoints

### `GET /health`
Vérification que le serveur est opérationnel.

**Réponse attendue**
```json
{"status": "ok"}
```

---

### `POST /convert`
Convertit un fichier `.docx` en archive ZIP contenant :
- `document.tex` — code LaTeX complet et compilable
- `references.bib` — entrées BibTeX (si des références sont détectées)
- `figure_1.png`, `figure_2.jpg`, … — images extraites du document

**Requête**
| Champ | Type | Obligatoire | Description |
|-------|------|-------------|-------------|
| `file` | `multipart/form-data` | ✅ | Fichier `.docx` |

**Réponse HTTP 200**
| Header | Valeurs | Signification |
|--------|---------|---------------|
| `Content-Type` | `application/zip` | Archive ZIP |
| `Content-Disposition` | `attachment; filename="nom_latex.zip"` | Nom du fichier |
| `X-Has-Bibliography` | `true` / `false` | Présence de références BibTeX |
| `X-Image-Count` | entier ≥ 0 | Nombre d'images extraites |

**Codes d'erreur**
| Code | Cause |
|------|-------|
| 400 | Pas de fichier, pas de nom, mauvaise extension |
| 422 | Document corrompu ou vide |
| 500 | Échec de conversion LaTeX |
| 502 | Erreur API Claude |

---

## 3. Tests cURL

### 3.1 Health check
```bash
curl -s http://localhost:8000/health | python3 -m json.tool
```

### 3.2 Conversion basique (fichier simple)
```bash
curl -s -X POST http://localhost:8000/convert \
  -F "file=@mon_document.docx" \
  -o resultat.zip \
  -D - | head -20
```

### 3.3 Vérifier les headers de réponse sans télécharger le corps
```bash
curl -sI -X POST http://localhost:8000/convert \
  -F "file=@mon_document.docx"
```

### 3.4 Voir X-Has-Bibliography et X-Image-Count
```bash
curl -s -X POST http://localhost:8000/convert \
  -F "file=@mon_document.docx" \
  -o resultat.zip \
  -w "\n--- STATS ---\nCode HTTP : %{http_code}\nTaille ZIP : %{size_download} octets\n"
```

### 3.5 Tester l'erreur 400 — mauvaise extension
```bash
curl -s -X POST http://localhost:8000/convert \
  -F "file=@rapport.pdf" | python3 -m json.tool
# Attendu: {"detail": "Only .docx files are supported."}
```

### 3.6 Tester l'erreur 400 — fichier vide
```bash
touch vide.docx
curl -s -X POST http://localhost:8000/convert \
  -F "file=@vide.docx" | python3 -m json.tool
# Attendu: {"detail": "The uploaded file is empty."}
```

### 3.7 Décompresser et lire le LaTeX
```bash
curl -s -X POST http://localhost:8000/convert \
  -F "file=@mon_document.docx" \
  -o /tmp/latex.zip && \
  mkdir -p /tmp/latex_out && \
  unzip -o /tmp/latex.zip -d /tmp/latex_out && \
  cat /tmp/latex_out/document.tex
```

### 3.8 Lister le contenu du ZIP
```bash
curl -s -X POST http://localhost:8000/convert \
  -F "file=@mon_document.docx" \
  -o /tmp/latex.zip && \
  unzip -l /tmp/latex.zip
```

---

## 4. Script de test automatisé

Exécutez ce script pour générer des fichiers `.docx` de test et valider l'API :

```bash
pip install python-docx requests
python3 test_api.py
```

**→ Voir le fichier `test_api.py` joint.**

---

## 5. Cas d'erreur à tester

### Cas valides (HTTP 200 attendu)

| # | Description | Vérification |
|---|-------------|--------------|
| V1 | Document texte simple | `document.tex` non vide |
| V2 | Document avec titres H1/H2/H3 | `\section`, `\subsection` présents |
| V3 | Document avec liste à puces | `\begin{itemize}` présent |
| V4 | Document avec liste numérotée | `\begin{enumerate}` présent |
| V5 | Document avec tableau | `\begin{tabular}` et `\toprule` présents |
| V6 | Document avec texte **gras** et *italique* | `\textbf{}`, `\textit{}` présents |
| V7 | Document avec exposant/indice | `\textsuperscript{}`, `\textsubscript{}` présents |
| V8 | Document avec images PNG/JPG | `X-Image-Count > 0`, figures dans ZIP |
| V9 | Document avec bibliographie | `X-Has-Bibliography: true`, `references.bib` dans ZIP |
| V10 | Document avec caractères spéciaux (é, ç, ü, …) | Pas de `?` dans le .tex |
| V11 | Document avec citation en corps de texte | `\cite{...}` présent dans .tex |
| V12 | Document long (> 20 pages) | Réponse complète, pas de timeout |

### Cas invalides (erreur attendue)

| # | Requête | Code | Message attendu |
|---|---------|------|-----------------|
| E1 | Aucun fichier joint | 422 | Unprocessable Entity |
| E2 | Fichier `.txt` | 400 | Only .docx files are supported |
| E3 | Fichier `.pdf` | 400 | Only .docx files are supported |
| E4 | Fichier `.docx` vide (0 octets) | 400 | The uploaded file is empty |
| E5 | Fichier `.docx` corrompu | 422 | Failed to parse Word document |
| E6 | Fichier `.docx` sans texte ni image | 422 | The document is empty |

---

## 6. Valider la réponse ZIP

### Contenu minimal attendu
```
Archive:  resultat.zip
  Length   Name
--------   ----
   12345   document.tex       ← OBLIGATOIRE
    2345   references.bib     ← si X-Has-Bibliography: true
    8900   figure_1.png       ← si X-Image-Count > 0
    5600   figure_2.jpg
```

### Vérifier que le LaTeX est compilable
```bash
# Installer pdflatex (TeX Live)
sudo apt-get install texlive-full     # Linux
brew install --cask mactex            # macOS

# Compiler
unzip -o resultat.zip -d /tmp/latex_out
cd /tmp/latex_out
pdflatex document.tex
bibtex document       # si references.bib présent
pdflatex document.tex
pdflatex document.tex
# → document.pdf généré sans erreur critique
```

### Éléments LaTeX à vérifier manuellement

```bash
# Classe de document choisie
grep -E "\\\\documentclass" /tmp/latex_out/document.tex

# Packages inclus (doit contenir graphicx, booktabs, amsmath...)
grep -E "\\\\usepackage" /tmp/latex_out/document.tex

# Figures (si images)
grep -E "\\\\includegraphics" /tmp/latex_out/document.tex

# Citations (si biblio)
grep -E "\\\\cite\{" /tmp/latex_out/document.tex

# Fin de document
grep "\\\\end{document}" /tmp/latex_out/document.tex
```

---

## 7. Checklist de validation

```
[ ] GET /health  →  {"status": "ok"}
[ ] Swagger UI accessible sur /docs
[ ] Conversion d'un .docx simple  →  ZIP téléchargeable
[ ] document.tex  →  \documentclass présent
[ ] document.tex  →  \begin{document} ... \end{document}
[ ] document.tex  →  packages: graphicx, booktabs, amsmath, hyperref
[ ] Headers ZIP  →  X-Has-Bibliography et X-Image-Count présents
[ ] Document avec titres  →  \section / \subsection générés
[ ] Document avec tableau  →  \begin{tabular} avec booktabs
[ ] Document avec images  →  figures dans ZIP + \includegraphics dans .tex
[ ] Document avec biblio  →  references.bib dans ZIP + \cite{} dans .tex
[ ] Erreur .pdf  →  HTTP 400
[ ] Fichier vide  →  HTTP 400
[ ] pdflatex document.tex  →  PDF généré sans erreur bloquante
```
