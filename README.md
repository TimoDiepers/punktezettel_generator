# Punktezettel Generator

Streamlit-App zur automatischen Erstellung von Punktezetteln (Klausur-Bewertungsbögen) als Excel-Dateien.

> **Hinweis:** In öffentlich gehosteten Instanzen sollten **keine vertraulichen oder personenbezogenen Daten** (z.B. echte Matrikelnummern, Namen)  hochgeladen werden.

## Voraussetzungen

[uv](https://docs.astral.sh/uv/) wird als Python-Paketmanager benötigt. Installation:

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

## Installation & Start

```bash
# Repository klonen
git clone <repo-url>
cd punktezettel

# App starten (uv installiert Abhängigkeiten automatisch)
uv run streamlit run app.py
```

Die App öffnet sich unter [http://localhost:8501](http://localhost:8501).

## Nutzung

1. **Vorlage herunterladen** — Beispiel-Excel mit dem erwarteten Format (Matr-Nr, Nachname, Vorname)
2. **Studierendenliste hochladen** — eigene Excel-Datei mit den drei Spalten
3. **Semester, Datum, Studis pro Mappe** konfigurieren
4. **Aufgaben anlegen** — Teilaufgaben, Punkte und optionale Beschreibungen pro Teilpunkt
5. **Punktezettel erstellen** und als `.xlsx` herunterladen
