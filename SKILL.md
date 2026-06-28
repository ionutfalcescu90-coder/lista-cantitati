---
name: lista-cantitati
description: >
  Generate a structural Bill of Quantities (Romanian "Lista de cantitati" /
  "antemasuratoare") as a formatted Excel workbook from Revit-exported IFC models —
  concrete (beton), formwork (cofraj) and reinforcement (armatura) per level and per
  element category (pereti, stalpi, grinzi, placi, radier, egalizare). Use this skill
  whenever the user has an IFC file (exported from Revit or similar BIM tools) and wants
  quantities extracted: "lista de cantitati", "antemasuratoare", "cantitativ de lucrari",
  "deviz de cantitati", "cubaje beton/cofraj/armatura", "extras de cantitati pe nivel",
  centralizator de cantitati, or a €/mp (euro pe mp construit) cost estimate from a BIM
  model. Trigger it even if the user only says "scoate cantitatile din IFC", "cat beton
  iese", "fa-mi cubajele", or shares a .ifc and asks what is in it quantity-wise. Also use
  to add a client-facing copy with hidden formulas/calculation columns.
---

# Lista de cantitati din IFC

Extracts concrete + formwork + reinforcement quantities from one or more IFC models and
writes a polished multi-sheet Excel workbook (one sheet per building corp/section, plus a
cost centralizator). Quantities are computed **geometrically** from the IFC mesh, so no
quantity parameters need to exist in the Revit model.

The bundled script `scripts/cantitati_ifc.py` ships with **example configuration values**
(a multi-corp residential layout). For a new project you copy it next to the IFC files and
edit the CONFIG block at the top (project title, IFC filenames, classes, indices, prices,
articles). The script already encodes hard-won engineering conventions — read
`references/conventii.md` before changing any extraction logic.

## Dependencies

```
pip install ifcopenshell openpyxl
```

`ifcopenshell` does the IFC parsing + geometry tessellation; `openpyxl` writes the xlsx.

## Workflow

1. **Locate the IFC(s).** Ask the user where they are (or look in the working dir for
   `*.ifc`). One IFC per building corp/section is typical.
2. **Reconnaissance first — do NOT generate blind.** Before producing the workbook, run a
   quick scan of each IFC to learn its real structure. This catches the issues that make a
   takeoff wrong. Print, per IFC: storey names + elevations; counts of
   IfcWall/IfcColumn/IfcBeam/IfcSlab; column type names (to tell concrete from steel); and
   slab type names (to spot egalizare/radier). See `references/conventii.md` →
   "Recon checklist" for a ready script.
3. **Configure** the CONFIG block of `cantitati_ifc.py` for this project: project title &
   address, the `SHEETS` list (map each IFC to a sheet), concrete classes, reinforcement
   indices, surface overrides, level uprates, and the incinta/consolidare articles.
4. **Run** it: `python cantitati_ifc.py <output.xlsx>`. It processes every configured IFC
   and writes the master workbook + a `_CLIENT` copy.
5. **Verify** the per-level distribution looks sane (walls present on every floor, top
   "retras" level has walls, no phantom basement). Spot-check one element by hand. Report
   the totals and any assumptions the user must confirm.

## The element model (why the numbers come out right)

These conventions are the difference between a believable takeoff and a wrong one. They are
implemented in the script; understand them so you can explain and adjust:

- **Walls are collected once.** `by_type("IfcWall")` already includes
  `IfcWallStandardCase` — never add a separate `by_type("IfcWallStandardCase")` or every
  wall doubles.
- **Vertical elements are assigned to the level whose slab they support (their TOP
  elevation), not the level they stand on.** Revit tags a wall/column to its base level;
  structurally it belongs to the storey above. This single rule fixes the classic "top
  floor has no walls / basement walls float" problem and naturally folds a slab-less base
  level (e.g. a Subsol that is really the ground-floor walls) into the right storey. Beams
  and slabs stay on their own Revit level.
- **Concrete columns count as `pereti` ("stalpi inclusi"); steel columns are excluded.**
  Detect concrete by type name (`Ortbeton`, a `C##/##` class, or `BA`); tubular steel
  (`RO …`, S235) is not concrete and carries no beton/cofraj.
- **Atic (parapet) walls go to `placi`**, not `pereti`, and do not count toward built area.
- **Egalizare** (blinding, `C12/15`) and **radier** (raft) slabs are split into their own
  categories so they get the right class and reinforcement index and the radier footprint
  is counted once (egalizare shares the same footprint — never sum both into the area).

## Geometry

Volume = signed mesh integral (divergence theorem); area = triangle sum. Formwork (cofraj)
heuristics: walls/beams/columns = total area − horizontal faces (lateral); slabs = soffit
(horizontal/2); radier = lateral edges only (cast on blinding, no soffit form); egalizare =
0. IFC length unit is read as **metre** for these Revit exports — verify with the recon if a
new model looks off by 10³.

## Output structure

- One **per-corp sheet**: columns A–K printable (Nivel · Element · Clasa · Beton mc · mc/mp ·
  Cofraj mp · mp/mp · Armatura kg · kg/mc · kg/mp · Suprafata), with the live results as
  **Excel formulas**; columns M–Q hold the raw IFC values, loss coefficients (1.05) and the
  reinforcement index (kg/mc) — these feed the formulas and sit outside the print area.
- Reinforcement is **estimated from consumption indices** (kg/mc) the user provides per
  category/level — the model is not assumed to contain rebar.
- A **cantitativ** sheet (different layout: Nr · Denumire · UM · Cantitate + a "Mod de
  calcul" column) for site works like incinta (shoring) and consolidare teren (rigid
  inclusions), with quantities as formulas so the derivation is visible.
- A **Centralizator** with editable unit prices (€/mc, €/mp, €/kg) and €/mp construit.
- Standard **notes** (cantitati teoretice, armatura din indici, exclusions) — and project
  notes can be pulled from an existing template workbook if present.
- A **client copy** (`*_CLIENT.xlsx`): calculation columns hidden, formulas hidden
  (Protection), Centralizator hidden — only the printable part remains.

## Configuration reference

All project-specific values live in clearly-marked constants near the top of the script:
`PROIECT`, `ADRESA`, `FAZA`, `PROIECTANT`; `CLASE_INFRA` / `CLASE_SUPRA` (forced concrete
classes, overriding the IFC); `INDICI_*` (reinforcement kg/mc per category, with per-level
overrides and a `_default`); `SHEETS` (the sheet↔IFC map, plus per-sheet `supr_inherit`
for partial-slab levels that borrow built area from the level above, and `majorari` to
uprate a level's quantities by a factor when an unmodeled level must be allowed for);
`ARTICOLE_INCINTA` / `ARTICOLE_CONSOLIDARE` (cantitativ rows, each a formula); unit prices
`PRET_BETON` / `PRET_COFRAJ` / `PRET_ARM`.

For the full rationale, the recon script, the per-corp configuration recipe, and how to add
a new section type, read **`references/conventii.md`**.
