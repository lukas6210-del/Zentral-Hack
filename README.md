# ClickTracker

**Windows-App, die beim Arbeiten automatisch Schritt-für-Schritt-Anleitungen erstellt.**
Aufnahme starten, normal arbeiten, jeder Klick wird mit Screenshot und fertig formuliertem Text festgehalten. Am Ende entsteht per Knopfdruck ein Word- oder PDF-Dokument im Layout der eigenen Firma.

**Alles bleibt lokal:** keine Cloud, kein Internet, keine Web-App. Passwörter werden nie gespeichert.

| | |
|---|---|
| Version | 0.4.0 |
| Plattform | Windows 10 / 11 |
| Sprache der Anleitungen | Deutsch (Sie-Form, Schweizer Schreibweise) |
| Früherer Name | Klickspur (bis Version 0.1) |

---

## Inhalt

1. [Ausgangslage und Ziel](#1-ausgangslage-und-ziel)
2. [Funktionen](#2-funktionen)
3. [Installation und Start](#3-installation-und-start)
4. [Bedienung](#4-bedienung)
5. [Technische Umsetzung](#5-technische-umsetzung)
6. [Datenschutz und Sicherheit](#6-datenschutz-und-sicherheit)
7. [Projektverlauf](#7-projektverlauf)
8. [Analysen](#8-analysen)
9. [Bekannte Grenzen](#9-bekannte-grenzen)
10. [Nächste Schritte](#10-nächste-schritte)

---

## 1. Ausgangslage und Ziel

Im IT-Support und in der Applikationsbetreuung werden laufend Anleitungen geschrieben, in den meisten Firmen von Hand mit Word und dem Snipping Tool. Das führt zu vier Problemen:

| Problem | Folge |
|---------|-------|
| Anleitungen schreiben ist mühsam (Screenshot, einfügen, zuschneiden, Text) | Es werden zu wenige geschrieben, Wissen bleibt in den Köpfen |
| Nach jedem Software-Update veralten die Screenshots | Niemand vertraut der Doku, der Support wird angerufen |
| Jede Person formuliert anders | Uneinheitliche, schwer lesbare Dokumente |
| Auf Screenshots sind Passwörter, Namen oder Patientendaten sichtbar | Datenschutzrisiko beim Weitergeben |

Die bekannten automatischen Tools (Scribe, Tango) sind Cloud-Dienste. Für Firmen mit sensiblen Daten, zum Beispiel Spitäler oder Verwaltungen, ist das oft nicht zulässig. Ausserdem schafft Microsoft die Windows-Schrittaufzeichnung (psr.exe) ab Windows 11 24H2 ab.

**Ziel von ClickTracker:** Eine Anleitung soll nebenbei beim normalen Arbeiten entstehen, einheitlich formuliert, im Firmenlayout und ohne dass ein Bild den Computer verlässt.

---

## 2. Funktionen

### Aufnahme
- Globale Aufnahme von Links-, Rechts- und Doppelklicks
- Screenshot des richtigen Monitors, auch bei mehreren Monitoren und hoher DPI-Skalierung
- Tippen wird zu einem Schritt zusammengefasst
- Tastenkombinationen werden erkannt und erklärt
- Schwebende Aufnahme-Leiste mit *Pause* und *Stopp*, die auf den Screenshots nicht zu sehen ist

### Automatische Texte
Das angeklickte Bedienelement wird über *Windows UI Automation* erkannt. Daraus entstehen fertige Sätze, zum Beispiel:

- *Klicken Sie auf die Schaltfläche „Speichern“ im Fenster „Editor“.*
- *Geben Sie im Feld „Benutzername“ den Text „peter.muster“ ein und drücken Sie die Tabulatortaste.*
- *Drücken Sie die Tastenkombination Strg + S (Speichern).*

### Bearbeitung
- Texte anpassen, Schritte per Drag & Drop umsortieren und mit `Entf` löschen
- **Schwärzen:** sensible Bereiche im Screenshot von Hand abdecken (neu in 0.4)
- Klick-Markierung (roter Kreis) und Feld-Rahmen werden erst beim Anzeigen bzw. Export gezeichnet und lassen sich deshalb nachträglich ändern

### Export
- Word (.docx) und PDF, A4, mit Titel, Datum, Einleitung und Seitenzahlen
- **Firmenvorlagen:** eigenes Logo, Schrift, Kopf- und Fusszeile (neu in 0.4)

### Speichern
- Automatisches Speichern jedes Schritts
- Projekte liegen lokal unter `%LOCALAPPDATA%` und bewusst **nicht** in OneDrive oder unter *Dokumente*

---

## 3. Installation und Start

| Variante | Vorgehen |
|----------|----------|
| **Installationsprogramm** | `ClickTracker-Setup-0.4.0.exe` ausführen. Zielrechner brauchen kein Python. |
| **Entwicklung** | Doppelklick auf die Start-Batchdatei. Beim ersten Start wird automatisch eine Python-Umgebung (`.venv`) eingerichtet und `requirements.txt` installiert. |

Voraussetzung für die Entwicklung: **Python 3.11 oder neuer**.

### App selbst bauen

1. `build_exe.ps1` ausführen. PyInstaller erzeugt einen Ordner mit der fertigen `.exe`.
2. Das Installationsprogramm wird anschliessend mit **Inno Setup** aus diesem Ordner erstellt.

---

## 4. Bedienung

1. **Aufnahme starten** (oder `Strg+R`). Das Hauptfenster minimiert sich, oben erscheint die Aufnahme-Leiste.
2. Ganz normal arbeiten. Jeder Klick, jede Texteingabe und jede Tastenkombination wird zu einem Schritt.
3. **Stopp:** Die Schritte erscheinen links in der Liste.
4. Titel und Einleitung eintragen, Texte bei Bedarf anpassen, Schritte umsortieren oder löschen.
5. Sensible Bereiche schwärzen.
6. Firmenvorlage wählen und **als Word** oder **als PDF exportieren**.

Ein Doppelklick auf die Vorschau öffnet den Screenshot in Originalgrösse.

> **Wichtig:** Die Anleitung vor dem Weitergeben immer kurz durchsehen. Die automatische Passworterkennung funktioniert nur, wenn das Programm das Feld korrekt als Passwortfeld meldet.

---

## 5. Technische Umsetzung

### Technologien

| Bereich | Bibliothek |
|---|---|
| Oberfläche | PySide6 (Qt) |
| Maus- und Tastatur-Hooks | pynput |
| Screenshots | mss, Pillow |
| Erkennung der Bedienelemente | uiautomation (Windows UI Automation) |
| Word-Export | python-docx |
| PDF-Export | fpdf2 |
| Paketierung | PyInstaller, Inno Setup |

### Aufbau des Codes

```
main.pyw         Einstieg (setzt DPI-Modus, startet die App)
app.py           Hauptfenster, Aufnahme-Leiste, Dialoge
recorder.py      Maus-/Tastatur-Hooks -> Worker -> Schritte
uia.py           Bedienelement unter dem Klick bzw. mit Fokus (UI Automation)
texts.py         Deutsche Schritt-Texte (Sie-Form, Schweizer Schreibweise)
capture.py       Screenshots, Schwärzen, Markierung, Zuschnitt
exporters.py     Word- und PDF-Export
model.py         Projekt und Schritte, Speichern als Ordner mit project.json + PNGs
winutil.py       Win32-Hilfsfunktionen (ctypes)
```

### Ablauf einer Aufnahme

Die Aufnahme ist bewusst auf drei Ebenen verteilt, damit Windows die Hooks nicht wegen Zeitüberschreitung entfernt und die Reihenfolge der Schritte stimmt:

```
Hook-Callbacks (pynput)   legen nur das Ereignis in eine Warteschlange
        |
Worker-Thread             Screenshot (sofort), UI Automation, Text erzeugen
        |
Speicher-Thread           Schwärzen, PNG schreiben, Signal an die Oberfläche
```

1. **Hook:** Ein Mausklick wird registriert und sofort in eine Queue gelegt. Mehr passiert hier nicht, weil Windows langsame Hooks automatisch entfernt.
2. **Worker:** Der Screenshot wird gemacht, solange der Zustand noch dem Klick entspricht. Danach fragt UI Automation ab, welches Element angeklickt wurde (Typ, Name, Fenster, Passwortfeld ja/nein). `texts.py` erzeugt daraus den Satz.
3. **Speicher-Thread:** Passwortfelder werden im Bild übermalt, danach wird das PNG im Projektordner gespeichert. Es gibt bewusst nur einen Speicher-Thread, damit die Schritte in der richtigen Reihenfolge ankommen.

### Datenmodell

Jede Anleitung ist ein Ordner mit einer `project.json` und einem Unterordner `images/`. Pro Schritt werden gespeichert: Aktion (Klick, Rechtsklick, Doppelklick, Eingabe, Taste), Text, Bildpfad, Klickposition, Rahmen des Elements, Fenster und Programm.

Weil Markierung und Rahmen als Koordinaten gespeichert und erst beim Anzeigen gezeichnet werden, bleibt das Originalbild unverändert. Farbe, Zuschnitt und Markierung lassen sich daher jederzeit ändern.

### Export

- Screenshots werden für den Export auf maximal 1800 px Breite verkleinert, damit die Dokumente klein bleiben.
- Pro Schritt: Überschrift „Schritt X“, Text, Screenshot. Überschrift und Text bleiben mit dem Bild auf derselben Seite.
- Der PDF-Export nutzt die Windows-Schrift Segoe UI (Ersatz: Arial) und setzt eine Fusszeile mit Titel und Seitenzahl.
- Mit Firmenvorlage werden Logo, Schrift sowie Kopf- und Fusszeile aus der Vorlage übernommen.

---

## 6. Datenschutz und Sicherheit

| Massnahme | Umsetzung |
|---|---|
| Keine Cloud | Kein Netzwerkzugriff, alle Daten bleiben auf dem Rechner |
| Passwortfelder | Inhalt wird **nie** gespeichert, auch nicht im Arbeitsspeicher. Das Feld wird im Screenshot übermalt, der Text lautet „Ihr Passwort“. |
| Zusätzliche Erkennung | Ein Feld gilt auch als geheim, wenn sein Name „Passwort“, „Kennwort“, „PIN“ o. ä. enthält |
| Schwärzen von Hand | Weitere sensible Bereiche (Namen, Patientendaten, E-Mail-Adressen) können abgedeckt werden |
| Speicherort | `%LOCALAPPDATA%`, bewusst nicht synchronisiert (kein OneDrive) |
| Aufnahme-Leiste | Ist auf den Screenshots nicht sichtbar |

**Hinweis für die IT:** Die App nutzt während der Aufnahme globale Tastatur- und Maus-Hooks. Virenscanner oder EDR-Software können das als Keylogger melden. Vor einem Einsatz im Betrieb muss die App deshalb von der IT-Sicherheit freigegeben werden.

---

## 7. Projektverlauf

| Version | Name | Inhalt |
|---|---|---|
| 0.1 | Klickspur | Erster Prototyp: Aufnahme, automatische Texte, Passwortschutz, Export nach Word und PDF, portable App über PyInstaller |
| 0.4 | ClickTracker | Neuer Name, Schwärzen von Hand, Firmenvorlagen, Installationsprogramm mit Inno Setup |

### Vorgehen

1. **Problem aus dem Alltag:** Ausgangspunkt war der Aufwand für Anleitungen im IT-Support.
2. **Prototyp:** Zuerst nur der Kern: aufnehmen, Text erzeugen, exportieren.
3. **Offene Fragen gesammelt:** In `FRAGEN.md` wurden Fragen zu Zielgruppe, Texten, Datenschutz, Aufnahme, Dokument und Verteilung festgehalten.
4. **Marktvergleich:** Nutzwertanalyse gegen bestehende Tools (siehe unten).
5. **Gezielte Weiterentwicklung:** Die Funktionen mit dem grössten Nutzen für Firmen wurden zuerst umgesetzt: Schwärzen und Firmenvorlagen.

---

## 8. Analysen

### Nutzwertanalyse

Vergleich mit den gängigen Alternativen, Skala 1–5, gewichtet (max. 500 Punkte):

| Kriterium | Gewicht | ClickTracker | Scribe | Tango | Snagit | Word manuell |
|---|---|---|---|---|---|---|
| Datenschutz / lokal | 20 % | 5 | 2 | 2 | 5 | 5 |
| Automatische Erfassung | 20 % | 4 | 5 | 5 | 2 | 1 |
| Kosten | 15 % | 4 | 2 | 2 | 4 | 5 |
| Export / Firmenvorlage | 10 % | 5 | 3 | 3 | 3 | 5 |
| Bedienbarkeit | 10 % | 3 | 5 | 5 | 4 | 3 |
| Nachbearbeitung | 10 % | 3 | 4 | 4 | 5 | 2 |
| Reifegrad / Support | 10 % | 1 | 5 | 4 | 5 | 4 |
| Desktop-Programme | 5 % | 4 | 4 | 4 | 5 | 5 |
| **Total** | | **380** | 360 | 350 | **395** | 360 |

ClickTracker liegt auf Rang 2. Der Rückstand auf Snagit kommt fast nur vom Reifegrad (Prototyp, keine Code-Signatur). Mit einer signierten, stabilen Version 1.0 läge ClickTracker bei rund 400 Punkten und damit auf Rang 1.

### Positionierung

| | Manuelle Erfassung | Automatische Erfassung |
|---|---|---|
| **Cloud** | – | Scribe, Tango |
| **Lokal** | Word, Snagit | **ClickTracker**, Folge |

Im Feld „lokal und automatisch“ gibt es kaum Konkurrenz. Der direkteste Mitbewerber ist Folge.

### Zeitersparnis (Schätzung)

| Anleitung mit 10 Schritten | Zeit |
|---|---|
| Manuell mit Word | ca. 25–30 min |
| Mit ClickTracker | ca. 8–10 min |
| **Ersparnis** | **ca. 2/3, rund 0,3 h pro Anleitung** |

Faustregel: **Je 1'000 Anleitungen pro Jahr spart ein Unternehmen rund 300 Stunden bzw. CHF 21'000** (bei CHF 70/h). Aktualisierungen bestehender Anleitungen sind darin nicht enthalten. Die Werte sind Schätzungen und sollen in einem Pilot gemessen werden.

### MVP

Version 0.4 bildet das MVP: Aufnahme, automatische Texte, Passwortschutz, Schwärzen, Firmenvorlagen und Export nach Word/PDF. Die Kernannahme, die damit getestet werden soll:

> Mitarbeitende erstellen mit ClickTracker Anleitungen deutlich schneller als mit Word, und die Firma darf es einsetzen, weil alles lokal und sicher bleibt.

---

## 9. Bekannte Grenzen

- Die Passworterkennung hängt davon ab, dass das Programm das Feld als Passwortfeld meldet. Windows, Office, Edge, Chrome und Qt tun das.
- Programme, die *als Administrator* laufen, und UAC-Dialoge sind nur eingeschränkt erfassbar.
- Scrollen und Drag & Drop werden noch nicht aufgezeichnet.
- Der Screenshot entsteht beim Herunterdrücken der Maustaste. Bei sehr schnellen Programmen kann die Wirkung des Klicks schon sichtbar sein.
- Noch keine Code-Signatur, Windows SmartScreen warnt beim ersten Start.
- Virenscanner bzw. EDR können die globalen Hooks melden.
- Oberfläche und Texte nur auf Deutsch, nur helles Design.

---

## 10. Nächste Schritte

1. **Freigabe durch die IT-Sicherheit** (Tastatur-Hooks, EDR)
2. **Code-Signatur**, damit Windows nicht mehr warnt
3. **Pilot** mit 5–10 Personen während 4 Wochen, mit Zeitmessung gegenüber Word
4. Je nach Feedback: Pfeile und Text im Bild, Zuschneiden, Scrollen und Drag & Drop, zentrale Vorlage per GPO/Intune

### Erfolgskriterien für den Pilot

| Kriterium | Ziel |
|---|---|
| Zeitersparnis (gemessen) | mindestens 50 % gegenüber Word |
| Nutzung | mindestens 3 Anleitungen pro Person in 4 Wochen |
| Textqualität | höchstens 2 von 10 Schritten müssen umgeschrieben werden |
| Sicherheit | keine EDR-Meldung, keine sichtbaren Passwörter |
| Zufriedenheit | Mehrheit will weiter damit arbeiten |
