# Maiershirts – Präsentationsmaster

PowerPoint-Vorlage im Maiershirts-Look nach dem Brand Book (Mai 2026):
Schwarz `#0F0F0F`, Sage `#819E72` (Pantone 7492 C), Sage Dark `#677E5B`,
Warm Sand `#E8DFD0`, Sand Dark `#D9D2C5`, Weiß. Schrift: Inter.

## Dateien

| Datei | Zweck |
|---|---|
| `dist/Maiershirts_Master.potx` | **Die Vorlage.** Doppelklick öffnet eine neue Präsentation mit allen Layouts und Platzhaltern. |
| `dist/Maiershirts_Master.pptx` | Beispieldeck: jedes Layout einmal mit Musterinhalt (15 Folien). Zum Abschauen und Kopieren. |
| `assets/logo.svg` | Original-Logo (Vektor) |
| `assets/logo-dark.png`, `assets/logo-light.png` | Logo schwarz (helle Folien) und weiß (dunkle Folien), erzeugt aus `logo.svg` |
| `build.js` | Erzeugt POTX und PPTX (`npm run build`) |
| `make-logo.js` | Erzeugt die beiden Logo-PNGs aus `logo.svg` (`npm run logo`) |

## So wird gearbeitet

1. `Maiershirts_Master.potx` doppelklicken. PowerPoint öffnet eine neue
   Präsentation mit den Maiershirts-Layouts und den 15 Musterfolien.
   Nicht benötigte Musterfolien löschen, die Vorlage selbst bleibt unverändert.
2. Folien über **Start → Neue Folie** anlegen und das passende Layout wählen.
   Jede Folie hat Platzhalter mit Hinweistext (z. B. „Foto einfügen“,
   „Was die Zahl bedeutet“). Platzhalter anklicken und ausfüllen.
3. Bilder: auf das Symbol im Bildplatzhalter klicken und Datei wählen. Das Bild
   wird auf den Rahmen zugeschnitten. Passt der Ausschnitt nicht:
   **Bildformat → Zuschneiden → Anpassen** (Logos) bzw. **Ausfüllen** (Fotos).
4. Farben und Schrift sind im Design hinterlegt: Neue Formen, Tabellen und
   Diagramme bekommen automatisch die Markenfarben, neue Textfelder Inter.
5. Musterfolien aus `Maiershirts_Master.pptx` können per Kopieren/Einfügen
   übernommen werden (Zieldesign verwenden).

**Schrift:** Inter muss auf jedem Rechner installiert sein, auf dem die
Präsentation bearbeitet oder gezeigt wird (kostenlos bei Google Fonts, die
statischen Schnitte „Inter“ installieren, nicht nur „Inter Variable“). Fehlt
sie, ersetzt PowerPoint sie durch Arial/Calibri. Für den Versand als PDF
exportieren, dann ist die Schrift eingebettet.

## Layouts

| Layout | Verwendung | Platzhalter |
|---|---|---|
| `MS_TITEL` | Titelfolie, dunkel | Titel, Untertitel, Kunde/Projekt, Datum/Angebotsnummer |
| `MS_ABSCHNITT` | Abschnittstrenner, dunkel | Nummer, Titel, Kurzbeschreibung |
| `MS_INHALT` | Standardfolie, weiß | Titel, Textkörper |
| `MS_INHALT_SAND` | Standardfolie, Warm Sand | Titel, Textkörper |
| `MS_FREI` / `MS_FREI_SAND` | Nur Titel, Fläche frei für Karten, Tabellen, Diagramme | Titel |
| `MS_BILD` | Foto links, Text rechts | Titel, Foto, Textkörper |
| `MS_BILD_VOLL` | Foto über volle Breite, Bildunterschrift auf Schwarz | Foto, Titel, Ergänzung |
| `MS_KENNZAHLEN` | Drei Kennzahl-Karten auf Sand | 3 × Zahl, Bedeutung, Quelle; Hinweis |
| `MS_PROZESS` | Vier nummerierte Schritte | 4 × Überschrift, Text |
| `MS_TEXTIL` | Drei Textilien mit Foto, Details und Preis ab | 3 × Foto, Name, Details, Preis |
| `MS_PREISSTAFFEL` | Vier Stufen Menge gegen Stückpreis | 4 × Stufe, Preis, Zusatz; Hinweis |
| `MS_ANGEBOT` | Angebotsübersicht mit Konditionen | Titel, Gültig bis, Lieferzeit, Zahlung, Hinweis (Tabelle: Musterfolie kopieren) |
| `MS_REFERENZEN` | Logo-Raster 4 × 3 für breite Firmenlogos | 12 × Logo |
| `MS_REFERENZEN_WAPPEN` | Logo-Raster 4 × 2 mit hohen Feldern für Vereinswappen | 8 × Logo |
| `MS_ABSCHLUSS` | Handlungsaufforderung, Ansprechpartner, nächster Schritt, dunkel | Titel, Text, Porträt, Name, Rolle, Kontakt, nächster Schritt, QR-Code, QR-Text |

Seitenrhythmus: helle und Sand-Folien wechseln sich ab (Kennzahlen,
Preisstaffel und Referenzen liegen auf Sand), dunkle Folien rahmen das Deck
(Titel, Abschnitt, Vollbild, Abschluss). Das Motiv der dunklen Folien ist eine
Bergflanke aus dem Logo in Sage Dark.

Typografie: Titel 28 pt, Text 16 pt, Nebentext 12 pt, Beschriftungen und
Fußzeile 10 pt. Nichts darunter, damit es am Beamer lesbar bleibt.

## Vorlage ändern (Entwickler)

```bash
npm install        # einmalig
npm run logo       # Logo-PNGs aus assets/logo.svg neu erzeugen
npm run build      # dist/*.pptx und dist/*.potx neu erzeugen
```

- Farben stehen am Anfang von `build.js` im Objekt `C`, Schriftgrößen in `FS`.
- Layouts sind `defineSlideMaster`-Blöcke, Beispielfolien folgen darunter.
- Flächen der Layouts (Karten, Kreise, Konditionenband) sind gerenderte
  Hintergrundbilder; Platzhalter und Logo liegen darüber.
- Nach dem Build patcht das Skript im Paket das Design (Farbpalette, Inter)
  und markiert Bildplatzhalter als Bildplatzhalter, weil pptxgenjs beides
  nicht kann.

Empfehlung: Vorlage per Skript fertigstellen, dann die POTX ans Team geben
und dort nur noch in PowerPoint arbeiten. Das Skript ist zum Erzeugen der
Vorlage gedacht, nicht zur Pflege einzelner Präsentationen.

## Beispielinhalte

Alle Texte, Zahlen, Preise, Kontaktdaten und Bilder im Beispieldeck sind
Platzhalter (siehe Notizen der jeweiligen Folie). Preise im Deck sind eine
Zusammenfassung; verbindlich ist immer das Angebot aus Lexware.
