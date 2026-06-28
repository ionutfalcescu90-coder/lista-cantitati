# lista-cantitati

Claude Code skill — generează o **Listă de cantități** (antemăsurătoare) structurală în
Excel pornind de la modele **IFC** exportate din Revit: beton, cofraj și armătură pe nivel
și pe categorie de elemente (pereți, stâlpi, grinzi, plăci, radier, egalizare), plus un
centralizator cu €/mp și o copie pentru client cu formulele ascunse.

Cantitățile se calculează **geometric** din mesh-ul IFC — nu e nevoie de parametri de
cantitate în modelul Revit.

## Conținut

- `SKILL.md` — declanșare + workflow + convenții cheie
- `scripts/cantitati_ifc.py` — scriptul generator (cu valori-exemplu de configurare)
- `references/conventii.md` — ghid detaliat: recon, convenții, configurare per corp, capcane

## Instalare pe alt PC

1. Clonează repo-ul în directorul de skill-uri Claude Code:
   ```
   git clone <url-repo> "%USERPROFILE%\.claude\skills\lista-cantitati"
   ```
   (sau clonează oriunde și copiază folderul în `~/.claude/skills/`)
2. Dependințe:
   ```
   pip install ifcopenshell openpyxl
   ```
3. În Claude Code, skill-ul se declanșează automat când ceri o listă de cantități dintr-un
   IFC. Pentru un proiect nou: copiază `scripts/cantitati_ifc.py` lângă fișierele `.ifc` și
   editează blocul CONFIG din capul scriptului.

## Utilizare rapidă

```
python cantitati_ifc.py output.xlsx
```

Procesează IFC-urile din `SHEETS` și scrie workbook-ul master + copia `_CLIENT`.
