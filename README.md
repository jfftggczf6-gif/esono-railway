# ESONO Document Parser — Micro-service Python

Micro-service de parsing de documents pour la plateforme ESONO.
Parse PDF, Excel, Word, PowerPoint, images, CSV avec les meilleures librairies Python.

## Déploiement sur Railway (recommandé)

### 1. Créer un repo GitHub

```bash
cd esono-parser
git init
git add .
git commit -m "Initial commit — ESONO document parser"
git remote add origin https://github.com/TON-USER/esono-parser.git
git push -u origin main
```

### 2. Déployer sur Railway

1. Aller sur https://railway.app
2. "New Project" → "Deploy from GitHub repo"
3. Sélectionner le repo `esono-parser`
4. Railway détecte automatiquement le Dockerfile
5. Ajouter la variable d'environnement : `PARSER_API_KEY=ton-secret-key-ici`
6. Attendre le déploiement (~2-3 minutes)
7. Copier l'URL publique (genre `https://esono-parser-production.up.railway.app`)

### 3. Configurer dans Lovable

Dans Lovable → Settings → Environment Variables :
```
VITE_PARSER_URL=https://esono-parser-production.up.railway.app
VITE_PARSER_API_KEY=ton-secret-key-ici
```

### 4. Tester

```bash
# Health check
curl https://esono-parser-production.up.railway.app/health

# Parse a PDF
curl -X POST https://esono-parser-production.up.railway.app/parse \
  -H "Authorization: Bearer ton-secret-key-ici" \
  -F "file=@test.pdf"
```

## Déploiement sur Render (alternative)

1. Aller sur https://render.com
2. "New" → "Web Service" → connecter le repo GitHub
3. Runtime: Docker
4. Plan: Starter ($7/mois)
5. Ajouter `PARSER_API_KEY` dans les env vars
6. Deploy

## Développement local

```bash
# Installer les dépendances
pip install -r requirements.txt

# Installer Tesseract (macOS)
brew install tesseract tesseract-lang

# Installer Tesseract (Ubuntu)
sudo apt-get install tesseract-ocr tesseract-ocr-fra

# Lancer le serveur
python main.py
# → http://localhost:8000

# Tester
curl -X POST http://localhost:8000/parse \
  -H "Authorization: Bearer esono-parser-dev-key" \
  -F "file=@test.pdf"
```

## API

### POST /parse

Parse un document et retourne le texte structuré.

**Headers:**
- `Authorization: Bearer <PARSER_API_KEY>`

**Body:** `multipart/form-data` avec un champ `file`

**Réponse:**
```json
{
  "success": true,
  "fileName": "etats_financiers_2024.pdf",
  "content": "--- Page 1/8 ---\nBILAN AU 31/12/2024\n...",
  "method": "pymupdf+pdfplumber",
  "category": "etats_financiers",
  "quality": "high",
  "summary": "pymupdf+pdfplumber — 8 pages — 3 tableaux — 12,450 caractères",
  "pages": 8,
  "tables_found": 3,
  "chars_extracted": 12450
}
```

### GET /health

```json
{"status": "ok", "version": "1.0.0"}
```

## Formats supportés

| Format | Librairie | Capacité |
|--------|-----------|----------|
| PDF (texte) | pymupdf | 50 pages max, extraction texte + tableaux |
| PDF (scanné) | pymupdf + Tesseract OCR | 20 pages max, OCR français + anglais |
| Excel (.xlsx, .xls) | openpyxl | Cellules mergées, formules, feuilles cachées, 500 lignes/feuille |
| Word (.docx) | python-docx | Texte + tableaux + en-têtes + styles |
| PowerPoint (.pptx) | python-pptx | Texte + tableaux + notes présentateur |
| Images (.jpg, .png, etc.) | Tesseract OCR | OCR français + anglais |
| CSV/TSV/TXT | chardet + python | Détection encodage auto (UTF-8, Windows-1252, ISO-8859-1) |

## Coût

- Railway Starter: ~$5/mois
- Toutes les librairies: open-source ($0)
- OCR: Tesseract gratuit ($0 — pas d'API Vision)
- Total: **~$5/mois**
