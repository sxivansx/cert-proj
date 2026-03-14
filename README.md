# Certificate Generator

Web app to bulk-generate certificates from a template image and a CSV/XLSX data file. Upload, point-and-click to place fields, customize fonts/colors/sizes per field, and download a ZIP.

**Live:** https://cert-generator-production.up.railway.app

## Features

- Drag-and-drop upload for template (PNG/JPG/WEBP) and data (CSV/XLSX)
- Visual field placement — click on the template to position text
- Per-field font, size, and color controls
- Custom font upload (TTF/OTF)
- ZIP download of all generated certificates

## Quick Start

```bash
git clone https://github.com/<your-username>/cert-proj.git
cd cert-proj
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open http://localhost:8080

## Deploy

```bash
railway login
railway init
railway up
railway domain
```

## Contributing

1. Fork the repo
2. Create a branch (`git checkout -b feature/my-feature`)
3. Make your changes
4. Run the app locally and verify everything works
5. Commit (`git commit -m "Add my feature"`)
6. Push (`git push origin feature/my-feature`)
7. Open a Pull Request

Keep PRs focused on a single change. No frameworks — vanilla HTML/CSS/JS only for the frontend.

## License

Free to use.
