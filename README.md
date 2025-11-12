
# Packlisten Converter (Web)

Dieses Repo macht deine bestehende `packliste_converter.py` (Tkinter-Desktop) im Web verfügbar.

## Lokal starten

```bash
python -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Öffne http://localhost:5000

## Deployment auf Render

1. Neues GitHub-Repo anlegen, Inhalt dieses Ordners pushen.
2. Auf https://render.com → **New +** → **Web Service** → Repository verbinden.
3. Environment: **Python** · Build: `pip install -r requirements.txt` · Start: `gunicorn app:app`
4. Optional **render.yaml** verwenden (Region Frankfurt, Free Plan).

## Zoho CRM (Web-Register)

- In Zoho CRM → **Einstellungen** → **Developer Space → Web-Tabs** (Web-Register).
- URL deines Render-Services eintragen, z. B. `https://packlisten-converter.onrender.com/`.
- Sichtbarkeit/Teambereiche festlegen.

### Hinweis

Die Web-App ruft `convert_file(input_path, output_path, user_dichtungen, show_message=False)` aus deiner vorhandenen Logik auf. 
Optional kannst du im Formular JSON für `user_dichtungen` übergeben.

