# DocSwitch: PDF and PPT converter

Python + Gradio web app with six tools:

| Tool | What it does | Library |
|---|---|---|
| PDF → PPT | Image slides (exact look) or editable text boxes | PyMuPDF, python-pptx |
| PPT → PDF | Converts .pptx / .ppt | LibreOffice (headless) |
| Images → PDF | Photos of notes / lab reports into one PDF | Pillow, PyMuPDF |
| Merge PDFs | Combine PDFs in upload order | pypdf |
| Extract pages | `1-3, 5, 8-` style page selection | pypdf |
| PDF → Text | Selectable text to .txt (no OCR) | PyMuPDF |

Limits (edit at the top of `app.py`): 50 MB per file, 120 pages per conversion.

## Run locally

```bash
pip install -r requirements.txt
# PPT -> PDF also needs LibreOffice:  sudo apt install libreoffice-impress
python app.py
# open http://localhost:7860
```

## Go live free on Render

1. Push this folder to a new public GitHub repository (files at the repo root:
   `app.py`, `Dockerfile`, `requirements.txt`, `README.md`).
2. On render.com sign in with GitHub, then **New + → Web Service**.
3. Pick the repository. Render detects the `Dockerfile` (Language: **Docker**).
4. Instance type: **Free**. Leave the other settings as they are and click **Deploy**.
5. Wait for the first build (several minutes, LibreOffice is large). The site goes live at
   `https://<service-name>.onrender.com`.

Every `git push` to the repo redeploys automatically.

## Notes

- Uploaded files live only in the server's temp folder and are not stored permanently.
- Render's free instances spin down after 15 minutes without traffic; the next visit takes about a minute to wake the app.
- Scanned PDFs need OCR for text extraction (possible next feature: `pytesseract`).
