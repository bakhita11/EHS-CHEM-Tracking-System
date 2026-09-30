# EHS Chemical Inventory Tracking System

Code accompanying **"A Hazard-Centric Information Architecture for Chemical Inventory
Management and Analytics"** (Salman, Khasawneh, Yassin & Rodriguez).

The manuscript designs an information architecture that turns an institutional chemical
inventory spreadsheet into something analyzable: a normalized relational core, a dedicated
hazard dimension with a chemical–hazard bridge relation, a dimensional warehouse, and a
role-secured reporting layer. This repository holds the pipeline behind the results, plus a
harness that checks every number the paper reports against the source data.

---

## Why there is a verification harness

A manuscript number and the data behind it drift apart during revision. Rather than assert
the figures, `paper1_pipeline.py` recomputes them and compares each against the value the
paper claims, naming the section it defends:

```
   manuscript location                         claimed   computed
  ----------------------------------------------------------------
ok Abstract: unique substances                     479        479
ok Abstract: hazard links                          515        515
ok Sec. res-construction: storage units            504        504
ok Sec. res-colocation: type-only conflicts         14         14
   ...
  20/20 checks passed
```

A failure tells you which sentence to fix. Two numbers in an early draft were caught this
way — see [Known corrections](#known-corrections).

---

## Requirements

Python 3.9 or later.

```bash
pip install pandas numpy openpyxl pyxlsb
```

`pyxlsb` is needed only for the binary (`.xlsb`) exports.

---

## Quickstart

```bash
# Put the institutional exports in data/, then:
python paper1_pipeline.py                        # single-site, with verification
python paper1_pipeline.py --multisite            # add the two-building analyses
python paper1_pipeline.py --export tables.xlsx   # write the derived tables
```

Pasting the file into a Jupyter or Colab cell runs it directly — it locates the workbooks
itself and prints the report, with no arguments to supply.

To point it somewhere other than `data/`:

```bash
DATA_DIR=/path/to/exports python paper1_pipeline.py --multisite
# or
python paper1_pipeline.py --single /path/to/inventory.xlsb
```

---

## What the pipeline does

**Preparation** (nine stages, Section 2.4 of the paper). Reads whichever sheet and header
row actually carries the data — the raw institutional exports put a multi-row reporting
banner above the real column labels. Removes template residue, normalizes placeholder
tokens, reconstructs CAS registry numbers and validates them against the check-digit
scheme, parses storage locations, and resolves container records to unique substances.

Two decisions in that stage matter more than they look:

- **Substances are grouped by validated CAS number alone.** Grouping on a key that
  concatenates CAS with a free-text name splits one substance across several groups — zinc
  appears in this inventory under eight name variants.
- **Container records are not de-duplicated.** One chemical is frequently held in several
  containers, and row-level de-duplication would discard real holdings. Records and
  physical containers are reported separately throughout.

**The hazard cross-walk** (Section 2.6). This is deterministic, not learned. A storage-group
code maps through the institution's own compatible-storage-group legend, and each
special-hazard flag contributes its own category. A substance carrying several categories
yields several rows in the bridge relation — which is the point, since a single-attribute
field would force a choice and lose the rest.

**Analyses** (Section 3). Hazard distribution, the high-consequence severity tier, NFPA 704
profiling, redundant holdings across locations, and co-location screening against the
segregation legend.

---

## Files

| File | Purpose |
|---|---|
| `paper1_pipeline.py` | The pipeline and verification harness. Self-contained; this is the file to read first. |
| `ehs_review_checks.py` | Operational checks requested by the EHS department — expiry scheduling, peroxide retention clocks, purchase cross-checking, special-handling reports. Not used for any result in the paper; these appear there as future work. |
| `build_combined_inventory.py` | Earlier standalone merge of the three survey workbooks. Superseded by `paper1_pipeline.py`, which does this internally; kept for reference. |

---

## Data

The chemical inventory analyzed in the paper is **not included here.** It is an institutional
Environmental Health and Safety record giving the identity, quantity, and physical storage
location of hazardous materials in a working university laboratory. Publishing
container-level locations of pyrophoric, peroxide-forming, and acutely toxic materials in a
named building is a security consideration the institution manages through the
role-based access model described in Section 2.8 of the paper.

The derived aggregate results are complete as presented in the manuscript. For access to the
underlying records, contact the corresponding author; requests are considered in
coordination with the institution's EHS office.

### Expected input format

The pipeline reads the standard institutional export. It scans every sheet for a CAS column
rather than assuming a layout, so both raw and cleaned exports work. The columns it uses:

| Column | Used for |
|---|---|
| `CAS No.` | Substance identity, after check-digit validation |
| `Proper Chemical Name` | Naming; fallback identity where no CAS is present |
| `A&M System Storage Group` | Hazard cross-walk |
| `Peroxide Forming?`, `Potentially Pyrophoric?`, `Potentially Explosive Chemical (PEC)?` | Special-hazard categories |
| `Health Hazard`, `Fire Hazard`, `Instability (Reactivity)` | The three NFPA 704 axes |
| `No of Containers` | Physical container counts |
| `Manufacturer Name` | Supplier analysis |
| `Unit No., Name or Description` | Storage location |
| `Room No.` | Location |

Missing columns degrade gracefully — the affected statistics are reported as zero rather
than raising.

---

## Known corrections

Two values in an early draft did not survive verification, both traceable to how the
storage-unit field is read. That field concatenates a cabinet label with a supplier catalog
number (`GS-113 // F/M8`), and two readings are possible:

- **Canonical** (`--unit-mode canonical`, the default) recognizes the unit and normalizes
  it. `GS-113 // F/M8` and `GS 113` become one unit, every spelling of a 2–8 °C
  refrigerator becomes one unit, and a cell holding only a catalog number resolves to no
  location at all.
- **Literal** (`--unit-mode literal`) takes whatever precedes the separator. Simpler, but it
  counts `MERCK/801809` — a catalog number — as a storage location, and treats three
  spellings of one refrigerator as three units.

The published manuscript reports the canonical figures: **504** distinct storage units and
**72** substances held in more than one location. An early draft reported 527 and 51 from the
literal reading. Both modes are retained so the difference is inspectable.

---

## Limitations

- Evaluated on one institution's inventory; hazard distributions are specific to that setting.
- The hazard fields originate in a single institutional reference library, so they cannot
  corroborate one another, and any claim about classification *accuracy* — as distinct from
  classification *structure* — requires an external source.
- The custody, movement, transaction and waste entities are defined in the schema but
  unpopulated; the source inventory holds no operational lifecycle records.
- Co-location screening requires locations recorded at individual-cabinet granularity. Where
  a unit is named only by cabinet type, a conflict cannot be confirmed.

---

## Citation

```bibtex
@article{Salman2026HazardCentric,
  author  = {Salman, Bakhita and Khasawneh, Mahmoud and Yassin, Muneeb and Rodriguez, Cristian},
  title   = {A Hazard-Centric Information Architecture for Chemical Inventory
             Management and Analytics},
  year    = {2026},
  note    = {Manuscript under review}
}
```

---


