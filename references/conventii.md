# Conventii si configurare — Lista de cantitati din IFC

Detailed reference for `scripts/cantitati_ifc.py`. Read this before changing extraction
logic or configuring a new project. Sections:

1. Recon checklist (run first, always)
2. Element conventions (the "why")
3. Geometry & formwork
4. Per-corp configuration recipe
5. Reinforcement indices
6. Surface area, uprates, inheritance
7. Cantitativ (incinta / consolidare teren)
8. Centralizator & costs
9. Notes
10. Client copy
11. Common pitfalls

---

## 1. Recon checklist (run first, ALWAYS)

Never generate the workbook from a new IFC blind. Revit exports carry quirks (steel vs
concrete columns, slab-less base levels, partial slabs, retras top levels, egalizare
modeled as slabs) that silently corrupt a takeoff. Scan first:

```python
import ifcopenshell
from collections import defaultdict
ifc = ifcopenshell.open(PATH)
print("schema:", ifc.schema)
storeys = sorted(ifc.by_type("IfcBuildingStorey"),
                 key=lambda s: float(s.Elevation) if s.Elevation else 0.0)
smap = {}
for rel in ifc.by_type("IfcRelContainedInSpatialStructure"):
    if rel.RelatingStructure.is_a("IfcBuildingStorey"):
        for el in rel.RelatedElements: smap[el.id()] = rel.RelatingStructure.Name
tc = {}
for rel in ifc.by_type("IfcRelDefinesByType"):
    for el in rel.RelatedObjects: tc[el.id()] = rel.RelatingType.Name or ""
print("Wall", len(ifc.by_type("IfcWall")), "Beam", len(ifc.by_type("IfcBeam")),
      "Slab", len(ifc.by_type("IfcSlab")), "Column", len(ifc.by_type("IfcColumn")))
print("col types:", {tc.get(c.id(),"") for c in ifc.by_type("IfcColumn")})
print("slab types:", {tc.get(s.id(),"") for s in ifc.by_type("IfcSlab")})
for s in storeys:
    print(s.Name, s.Elevation)
```

What to look for:
- **Storey names + elevations** → defines the level rows and feeds the "assign by top"
  rule. Note any `Subsol`/base level with no slabs, any `Parter dx`, any retras (`Etaj Ndx`)
  top level.
- **Column types** → `Ortbeton`/`C##/##`/`BA` = concrete (counted); `RO …`/S235 = steel
  (excluded). A frame corp may be column-dominated (e.g. 60+ concrete columns).
- **Slab types** → `Egalizare`, `Radier`, `Placa` are different categories.
- **Counts** → sanity for the final reconciliation.

To diagnose a "missing walls on the top floor" situation, also bucket walls by their
base-Z vs top-Z (geometry) and compare to the storey elevations — walls whose TOP sits at
a storey's elevation belong to that storey.

---

## 2. Element conventions (the "why")

| Rule | Why |
|---|---|
| `by_type("IfcWall")` only (it already includes `IfcWallStandardCase`) | A second `by_type("IfcWallStandardCase")` call doubles every wall. |
| Vertical elements (walls + concrete columns) → level of the slab they SUPPORT (top elevation) | Revit tags verticals to their base level. Structurally a wall from floor N to floor N+1 supports the N+1 slab and belongs to that storey. Fixes empty top floor, floating basement walls, and absorbs a slab-less base level automatically. |
| Beams & slabs → their own Revit level | They sit at one elevation; that elevation is their level. |
| Concrete columns → `pereti` ("stalpi inclusi"); steel columns excluded | Steel tubular columns are not concrete/cofraj. Concrete = `Ortbeton`/`C##/##`/`BA` in the type name. |
| The `pereti` row is **labelled by content**: only walls → "Pereti BA"; only columns → "Stalpi BA"; both → "Pereti BA (stalpi inclusi)" | A level with no walls (e.g. a frame parter/etaj) reads as "Stalpi BA", not the misleading "Pereti". Tracked via per-row `wall`/`col` counts. |
| Beams whose type name contains `BS` (e.g. `GF BS` = grinda de fundatie beton simplu) → own category `fundatii_bs` | Continuous plain-concrete foundations are a different material: unreinforced and cast in the trench. Class via `CLASA_FUND_BS`, **no reinforcement, cofraj = 0** (like egalizare). Without this they fall into `grinzi` and wrongly pick up a BA reinforcement index. |
| Atic walls → `placi`, excluded from built area | Parapets are not floor slabs; their footprint is negligible and must not inflate suprafata. |
| Egalizare slabs → own category, class `C12/15`, cofraj 0, excluded from built area | Blinding is unreinforced, cast on ground, and shares the radier footprint — counting its area double-counts the footprint. |
| Radier slabs → own category, own reinforcement index, cofraj = lateral edges only | The raft is cast on blinding (no soffit formwork) and is reinforced differently from suspended slabs. |

The "assign by top" rule is implemented via `nivel_dupa_cota(maxz)` (nearest storey
elevation to the element's top Z). It changes EVERY vertical's level by +1 relative to
Revit, but because mid-floors all look similar this is only visible (and correcting) at the
extremes — which is exactly where Revit's base-level tagging is wrong.

---

## 3. Geometry & formwork

- **Volume**: signed tetrahedron sum over mesh faces (divergence theorem), `abs()` at the
  end — sign-stable, subtracts openings/voids correctly.
- **Area**: half-cross-product per triangle. Horizontal faces detected by `|nz|/mag > 0.9`.
- **Formwork (`cofraj_val`)**:
  - walls / beams / columns → `area_tot - area_h` (lateral faces incl. ends — a small,
    deliberate over-estimate that a QS accepts; absorbed by the 1.05 factor).
  - slabs → `area_h / 2` (soffit only; the top is finished/cast against nothing).
  - radier → `area_tot - area_h` (perimeter edges; soffit is on the blinding).
  - egalizare → 0.
- **Units**: these Revit exports report `LENGTHUNIT = METRE`, so geometry is taken directly
  in metres. If a new model's volumes look 10³ off, check the unit and adjust.
- The slab-soffit-only convention slightly under-counts slab edges; compensated by the flat
  `+20 mp/level` cofraj allowance on `placi`.

---

## 4. Per-corp configuration recipe

Each entry in `SHEETS` maps one IFC to one sheet. Fields:

```python
{"name": "Suprastructura T1.E",            # sheet tab name
 "titlu": "SUPRASTRUCTURA — Corp Est (T1.E)",
 "ifc": "corp-est.ifc",                    # filename in the same folder; None => empty sheet
 "clase": CLASE_SUPRA,                       # forced concrete classes per category
 "indici": INDICI_SUPRA,                     # reinforcement kg/mc per category
 "supr_inherit": {"4 - Parter": "5 - Parter dx"},  # partial-slab level borrows area
 "majorari": {"Etaj 1": 1.5},                # uprate a level's quantities x factor
 "scari": False,                             # disable the +3mc/+20mp stair allowance on placi
 "placi_labels": {"_default": "Placa"}},     # per-level label for the placi row
```

- Use the EXACT storey names from the recon (with the `"N - "` prefix) as keys in
  `supr_inherit`; `indici`/`majorari`/`placi_labels` keys match the label after the prefix
  (`"Parter dx"`); `placi_labels` also takes a `"_default"`.
- A corp without `Parter dx` should use an index set without the `Parter dx`/`Etaj 1`-slab
  specials (see `INDICI_SUPRA_NODX`).
- `scari: False` for a corp that is only ground floor (no stairs) so placi rows don't get the
  stair allowance; pair with `placi_labels` to drop "scari/reborduri" from the label.
- A `tip: "cantitativ"` entry (no IFC) routes to the incinta/consolidare layout instead.

## 5. Reinforcement — extracts first, indices only as fallback

If the project has rebar extracts (`extras de armare`), **use the real kg, do not estimate**.
Configure `ARMATURA_EXTRAS[sheet][(nivel, categorie)] = [components]`:

- Parse each `*E_EXTRAS*.pdf` for `Greutate totala` (the recapitulatie total kg). PyMuPDF
  (`fitz`) + regex `r'Greutate totala\s*\n\s*([\d.,]+)'` reads it directly; the per-marca
  rows sum to that total.
- Each cell value is a **list of components**. With >1 component the H cell is written as a
  formula `=a+b+c` so the contributors stay identifiable in Excel; with one it's a plain value.
- Group `mustati` (starter bars) with the element they reinforce. Foundation starters
  (mustati pereti, mustati stalpi) usually belong to the **radier** (or the foundation beam
  they are cast in) — the extract for "armare fundatii" often already covers the foundation
  beams' reinforcement.
- An **empty list `[]`** means "counted elsewhere" → the H cell is left blank (e.g. grinzi de
  fundare whose steel is inside the radier extract). This avoids double counting.
- **A sheet listed in `ARMATURA_EXTRAS` never uses indices**: any (nivel, categorie) without
  an entry is left blank, not estimated. So the sheet total equals the sum of the extracts
  exactly (plus any explicit manual additions you add as extra list components, e.g. a stair).
- Verify: sum of all `Greutate totala` across the extract PDFs must equal the workbook total.

Indices below are the **fallback** when there are no extracts. Armatura is then estimated as
`beton * indice (kg/mc)`, configured per category with optional per-level overrides and a
`_default`:

```python
INDICI_SUPRA = {
    "pereti": {"Parter": 180, "Parter dx": 180, "_default": 130},  # stalpi inclusi
    "grinzi": {"_default": 135},
    "placi":  {"Etaj 1": 110, "_default": 105},   # e.g. transfer slab over parter dx
}
INDICI_INFRA = {"radier":{"_default":135}, "pereti":{"_default":150},
                "grinzi":{"_default":200}, "placi":{"_default":110}}
```

The H column formula is `=IF(index>0, beton*index, "")`, so a category with no index
(egalizare) yields blank. The takeoff carries a standard note that armatura is from indices.
A quick cage check (e.g. 4Ø14 + Ø8/200, 6 m ≈ 43 kg) is a good way to sanity-test a
proposed kg/buc for piles/inclusions.

## 6. Surface area, uprates, inheritance

- **Built area (`Suprafata construita`)** per level = sum of real slab footprints
  (`area_h/2`), excluding egalizare and atic.
- **`supr_inherit`**: a level with only partial slabs (e.g. a Parter whose floor is mostly
  open to the Parter dx above) borrows the built area of the named level so its mc/mp and
  mp/mp indices are meaningful.
- **`majorari`**: multiply a level's beton/cofraj/armatura by a factor, shown live in the
  formula (e.g. `=(M*N+3)*1.5`), to allow for an unmodeled level (a mansarda you have no
  data for). Keep it transparent and note it — or omit the note if the client copy shouldn't
  show it.
- **Placi allowance**: `+3 mc` beton and `+20 mp` cofraj per placi row for unmodeled stairs;
  visible in the formula. Does not apply to radier/egalizare.

## 7. Cantitativ (incinta / consolidare teren)

A different sheet type for site works, listed as `tip: "cantitativ"`. Layout:
Nr · Denumire · UM · Cantitate · **Mod de calcul** (internal note column, kept out of the
print area). Quantities are Excel formulas so the derivation is auditable, e.g.:

- Excavation + taluz: `=<area>*<depth> + <perim>*<proj>*<height>/2` (flat volume +
  triangular-section taluz over a length; taluz volume = length × proj × height / 2).
- Backfill from taluz: reference the taluz portions of the excavation rows, e.g.
  `=(D6-<area1>*<depth1>)+(D7-<area2>*<depth2>)`, so it auto-updates.
- Bored piles: foraj `=n*L`; beton `=n*PI()*r^2*L`; armatura `=beton*index` or `=n*kg/buc`.

Articles are tuples `(denumire, UM, lambda r: "=formula", "mod de calcul")`; the lambda gets
the row number so a row can reference the one above (`=D{r-1}*160`). Move non-quantity
articles (torcret, spraituri, ancore) into the notes instead of table rows when they are
recommendations/exclusions rather than firm quantities.

## 8. Centralizator & costs

Pulls each corp's totals via cross-sheet formulas (`='Sheet'!D{tot}` etc.), applies editable
unit prices (`PRET_BETON`/`PRET_COFRAJ`/`PRET_ARM`, €/mc, €/mp, €/kg — material+manopera),
and computes cost and **€/mp** (cost / built area, all levels). `write_sheet` returns the
`tot_row` so the centralizator knows where each sheet's total is. Note that infrastructure
€/mp is naturally high (massive raft + thick subsol walls over the subsol footprint).

## 9. Notes

`write_sheet(notes=[...])` writes wrapped note rows under the table, inside the print area,
with auto height. The script can pull existing project observations from a template
workbook (`citeste_note_template`) — adapt the cell coordinates to the template at hand.
Standard notes: "cantitati teoretice, fara pierderi", "armatura estimata pe baza de indici
de consum", exclusions (acoperis/confectii metalice/structura lemn, amenajari, bordaje,
confinare), dewatering wells / hydro study caveats.

## 10. Client copy

`make_client_copy(master, client)` loads the master and: hides calculation columns (L+ on
structural sheets, E+ on the cantitativ), sets `Protection(hidden=True)` on formula cells +
`ws.protection.sheet = True` so formulas don't show in the bar, hides the Centralizator
sheet, and locks workbook structure. No password by default (a deterrent, not a lock — add
one if true locking is needed, or delete the Centralizator entirely for a hard cut). Values
display correctly because Excel recalculates openpyxl formulas on open.

## 11. Common pitfalls

- **Doubled walls** → a stray `IfcWallStandardCase` query. Reconcile total element count.
- **Empty top floor / floating basement** → not using assign-by-top for verticals.
- **Missing concrete columns** → only querying Wall/Beam/Slab; a frame corp loses most of
  its volume. Include `IfcColumn` and filter steel.
- **Doubled radier area** → egalizare counted in built area or merged into the radier line.
- **`#VALUE!`** → divide-by-zero on a level with no built area; guard ratios with `IFERROR`.
- **Excel "repair" prompt on open** → overlapping merged cells; never merge two ranges that
  share a cell.
- **File locked / PermissionError on save** → the xlsx is open in Excel; save to a new name
  or ask the user to close it.
- **Per-corp index mismatch** → applying a `Parter dx`/transfer-slab index to a corp that
  has no such level; give that corp its own index set.
