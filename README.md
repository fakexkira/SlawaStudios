# SlawaStudios

## Präsentation: Strukturwandel in der US-amerikanischen Landwirtschaft

`praesentation/` enthält die Quelldateien des Foliensatzes (24 Folien, 16:9) zum Thema
„Bewässerungslandwirtschaft als Auslöser tiefgreifender agrarstruktureller Prozesse – die Great Plains (USA)“.

- `praesentation/project/deck.json` – Reihenfolge, Abschnitte, Schriften
- `praesentation/project/slides/*.html` – eine Datei pro Folie (inkl. Sprechernotizen in `<aside>`)
- `praesentation/bilder/` – verwendete Bilder
  - NASA/GSFC/METI/ERSDAC/JAROS, U.S./Japan ASTER Science Team: Finney County, Kansas, 24.06.2001 (gemeinfrei)
  - U.S. Geological Survey, SIR 2017–5040: Wasserspiegeländerung im High-Plains-Aquifer (gemeinfrei)
  - Billy Hathorn: Feedlot bei Rocky Ford, Colorado (CC BY-SA 3.0, Wikimedia Commons)

## Präsentation: Lachs aus dem Käfig (Intensivierung auch im Meer?)

`praesentation-lachs/` enthält eine 15-Folien-Präsentation (16:9) zu Aquakultur und Lachszucht in Norwegen,
ohne Bilder und ohne Quellenzeilen auf den Folien. Diagramme sind aus den Werten der Buchseite gezeichnet.

- `Lachs-aus-dem-Kaefig.html` – eigenständige Datei, läuft offline (← → blättern, `N` Sprechernotizen, `F` Vollbild)
- `Lachs-aus-dem-Kaefig.pdf` – PDF-Export
- `Lachs-aus-dem-Kaefig.pptx` – PowerPoint-Fassung (bearbeitbare Formen, Sprechernotizen; Schriften Impact, Georgia, Courier New; `python3 build_pptx.py`)
- `Lachs-billig.pptx` – bewusst „billige“ Variante (Comic Sans/Times/Arial, knallige Farben, Rahmen und Stern; `CHEAP=1 python3 build_pptx.py`)
- `src/deck.html`, `src/fonts/`, `build.py` – Quelle und Build (`python3 build.py`, braucht Node + Playwright für das PDF)
