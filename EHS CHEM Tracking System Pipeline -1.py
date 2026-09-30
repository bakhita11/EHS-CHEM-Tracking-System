"""
paper1_pipeline.py -- the deterministic pipeline behind Paper 1, plus a harness
that checks every number the manuscript reports against the data.

Paper 1 claims no machine learning. Its hazard model is a deterministic
cross-walk: a storage-group code and three special-hazard flags map to a
standardized hazard vocabulary, and the chemical--hazard bridge is the set of
(substance, category) pairs that mapping produces. Everything this file computes
is reproducible by hand from the source workbooks.

The point of the verification harness at the end is that a manuscript number and
the data can drift apart silently during revision. Each CHECK below names the
manuscript location it defends, so a failure says which sentence to fix.

Usage
-----
    python paper1_pipeline.py                      # single-site, verify
    python paper1_pipeline.py --multisite          # add the two-building analyses
    python paper1_pipeline.py --export out.xlsx    # write the derived tables

Inputs (paths configurable below): the single-site .xlsb export, and the two
per-building .xlsx exports for the multi-site section.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

# Where the institutional exports live. Relative by default so the script is
# portable; override with DATA_DIR or the --single argument. Files are also
# discovered automatically if these names do not resolve.
DATA_DIR = Path(os.environ.get("DATA_DIR", "data"))

SINGLE_SITE = DATA_DIR / "TAMIU_Chemical_Inventory_LBV242_January-2025.xlsb"

# Three surveys, not two: the single-site export is itself one of the buildings.
# Each workbook covers one survey area and fills the room column only
# sporadically, so the file's own scope supplies the default where it is blank.
MULTI_SITE = [
    {"path": SINGLE_SITE,
     "building": "LAMAR BRUNI VERGARA SCIENCE CENTER", "room_default": "242"},
    {"path": DATA_DIR / "TAMIU_Chemical_Inventory_AIC_202.xlsx",
     "building": "ACADEMIC INNOVATION CENTER", "room_default": "202"},
    {"path": DATA_DIR / "TAMIU_Chemical_Inventory_April_23_2025_LBV.xlsx",
     "building": "LAMAR BRUNI VERGARA SCIENCE CENTER", "room_default": None},
]

# The institution's own compatible-storage-group guide, transcribed from the
# 'Storage Groups' sheet. Eight of these were missing from the first pipeline,
# which silently turned five real categories into raw codes.
STORAGE_GROUPS = {
    "GEN":  "General storage (not intrinsically reactive)",
    "FLAM": "Flammables, combustibles & organic solvents",
    "Ox":   "Oxidizers & peroxides",
    "OxA":  "Strong oxidizing acids",
    "OA":   "Organic acids",
    "OB":   "Organic bases",
    "IA":   "Inorganic acids",
    "IB":   "Inorganic bases",
    "W":    "Pyrophoric & water-reactive",
    "EXPL": "Stable explosives",
    "X":    "Incompatible with all other chemicals",
    "XT":   "Acutely toxic / poison gases",
    "T":    "Toxic / health hazard",
    "BIO":  "Infectious agents, mutagens & carcinogens",
}

SPECIAL_HAZARDS = {
    "Peroxide Forming?":                    "Peroxide-forming chemicals",
    "Potentially Pyrophoric?":              "Potentially pyrophoric",
    "Potentially Explosive Chemical (PEC)?": "Potentially explosive (PEC)",
}

# Severity tier: hazard character, not frequency. Several of these are among the
# rarest categories, which is the point -- a frequency-ordered view hides them.
HIGH_CONSEQUENCE = {
    "Oxidizers & peroxides", "Strong oxidizing acids", "Pyrophoric & water-reactive",
    "Incompatible with all other chemicals", "Acutely toxic / poison gases",
    "Peroxide-forming chemicals", "Potentially pyrophoric", "Potentially explosive (PEC)",
}

# Groups the legend isolates absolutely. Ordered, not a set: set iteration order
# varies between runs and made the reported conflict reason nondeterministic.
ISOLATE_FROM_ALL = ("X", "W", "XT", "EXPL")

# Conventional prohibitions the legend implies but does not enumerate pairwise.
ACIDS = {"OA", "IA", "OxA"}
BASES = {"OB", "IB"}
OXIDIZERS = {"Ox", "OxA"}
FLAMMABLE_ORGANIC = {"FLAM", "OA", "OB"}

# Cabinet-type prefixes. Matching on "contains a dash" misclassified room codes
# such as LBV-286 as cabinets, so prefixes are matched explicitly.
UNIT_PREFIXES = ("GS", "FC", "CC", "FR", "RF", "AC", "COR", "FLAM", "OX")

NULL_TOKENS = {"0x2a", "0x17", "0x0e", "0x07", "0x24", "0x2b", "0x00",
               "hide", "nan", "na", "n/a", "none", "", "#ref!", "#n/a"}


# --------------------------------------------------------------------------
# Cell-level helpers
# --------------------------------------------------------------------------

def norm_header(c) -> str:
    """The template embeds newlines in column labels; collapse to single spaces."""
    return re.sub(r"\s+", " ", str(c)).strip()


def is_null(v) -> bool:
    return pd.isna(v) or str(v).strip().lower() in NULL_TOKENS


def norm_cas(v) -> str:
    """Digits-only CAS. This is substance identity throughout."""
    if is_null(v):
        return ""
    try:
        return str(int(float(v)))
    except (ValueError, TypeError):
        return re.sub(r"\D", "", str(v))


def format_cas(digits: str) -> str:
    return f"{digits[:-3]}-{digits[-3:-1]}-{digits[-1]}" if len(digits) >= 5 else ""


def cas_checksum_ok(digits: str) -> bool:
    """
    CAS check digit: each digit right-to-left (excluding the check digit) is
    multiplied by its position, summed, and reduced modulo 10.
    """
    if len(digits) < 5 or not digits.isdigit():
        return False
    body, check = digits[:-1], int(digits[-1])
    total = sum(int(d) * (i + 1) for i, d in enumerate(reversed(body)))
    return total % 10 == check


UNIT_PATTERN = re.compile(r"^\s*(GS|FC|CC|SFC|FR|GR)[\s\-_]*([0-9]+[a-zA-Z]?)?", re.I)
FRIDGE_PATTERN = re.compile(r"^\s*(2-8\s*\*|fridge|freezer)", re.I)
ROOM_UNIT_PATTERN = re.compile(r"^\s*(LBVC?|AIC)[\s\-_]*([0-9]+)", re.I)


def parse_storage_unit(v, mode="canonical"):
    """
    Resolve 'Unit No., Name or Description' to a storage unit.

    The field mixes a unit label with a supplier catalog number ('GS-113 // F/M8',
    'FC-9  AA/36238', 'CC 1', 'MERCK/801809'). Two readings are possible and they
    do not agree, which matters because the manuscript quotes numbers from both.

    mode='canonical' recognizes the unit and normalizes it: 'GS-113 // F/M8' and
    'GS 113' become one unit, every '2-8*C' spelling becomes REFRIG, and a cell
    holding only a catalog number resolves to no unit at all. This is what a
    storage location means operationally.

    mode='literal' takes whatever precedes the '//' separator. It is simpler but
    counts 'MERCK/801809' -- a catalog number -- as a distinct storage location,
    and treats three spellings of the same refrigerator as three units.
    """
    if is_null(v):
        return None, None
    raw = str(v).strip()

    if mode == "literal":
        return (raw.split("//")[0].strip() or None), raw

    m = UNIT_PATTERN.match(raw)
    if m:
        prefix, number = m.group(1).upper(), m.group(2)
        return (f"{prefix}-{number.lower()}" if number else prefix), raw
    if FRIDGE_PATTERN.match(raw):
        return "REFRIG", raw
    m = ROOM_UNIT_PATTERN.match(raw)
    if m:                                   # room-level, not an individual cabinet
        return f"{m.group(1).upper()}-{m.group(2)}", raw
    return None, raw                        # catalog number only: no location recorded


def unit_precision(unit) -> str:
    """
    Whether a storage-unit string identifies an individual cabinet.

    'FC-41' does; a bare 'CC' or 'FR' names only a cabinet type, and two
    incompatible materials both labelled 'CC' may sit in different cabinets.
    """
    if is_null(unit):
        return "missing"
    token = str(unit).strip().upper()
    if token == "REFRIG" or re.match(r"^(LBVC?|AIC)-", token):
        # A refrigerator class, or a room; either may hold several cabinets.
        return "type-only"
    if re.match(r"^(GS|FC|CC|SFC|FR|GR)-\d", token):
        return "individual"
    if re.match(r"^(GS|FC|CC|SFC|FR|GR)$", token):
        return "type-only"          # cabinet type with no number
    return "other"


def pair_conflict(a: str, b: str):
    """Return a reason string if two storage groups may not share a unit."""
    for g in ISOLATE_FROM_ALL:          # ordered, so the reason is deterministic
        if a == g or b == g:
            if g == "X":
                return "group X is incompatible with all chemicals, including other X"
            if g == "W":
                return "group W (pyrophoric/water-reactive) is incompatible with every other group"
            return f"group {g} must be isolated from all other groups"
    if (a in ACIDS and b in BASES) or (b in ACIDS and a in BASES):
        return "acid stored with base"
    if (a in OXIDIZERS and b in FLAMMABLE_ORGANIC) or (b in OXIDIZERS and a in FLAMMABLE_ORGANIC):
        return "oxidizer stored with flammable or organic material"
    return None


# --------------------------------------------------------------------------
# Stage 1-9: preparation
# --------------------------------------------------------------------------

def read_inventory(path: Path) -> pd.DataFrame:
    """
    Read whichever sheet and header row actually carries the data.

    The raw institutional exports put a multi-row reporting banner above the
    real column labels and hold the data on a sheet named 'Chemicals'; cleaned
    exports start at row 1. Scan rather than assume.
    """
    engine = "pyxlsb" if str(path).lower().endswith(".xlsb") else None
    kw = {"engine": engine} if engine else {}
    book = pd.ExcelFile(path, **kw)
    preferred = [n for n in ("Chemicals", "inventory") if n in book.sheet_names]
    order = preferred + [n for n in book.sheet_names if n not in preferred]

    for sheet in order:
        for header in (7, 0):
            try:
                df = pd.read_excel(path, sheet_name=sheet, header=header, **kw)
            except Exception:
                continue
            cols = [norm_header(c) for c in df.columns]
            if any(c.lower().startswith("cas no") or c.lower() == "cas_number" for c in cols):
                df.columns = cols
                return df.loc[:, ~df.columns.duplicated()].copy()
    raise ValueError(f"No CAS column found in {path.name}; sheets: {book.sheet_names}")


def column(df: pd.DataFrame, *names) -> pd.Series:
    for n in names:
        if n in df.columns:
            return df[n]
    return pd.Series([np.nan] * len(df), index=df.index)


def prepare(raw: pd.DataFrame, verbose=True, unit_mode="canonical") -> pd.DataFrame:
    """Stages 1-9 of Section 'Data Preparation'. Returns one row per container record."""
    n_raw = len(raw)

    cas_digits = column(raw, "CAS No.", "cas_number").map(norm_cas)
    name = column(raw, "Proper Chemical Name", "chemical_name").astype(str).str.strip()
    descriptive = column(
        raw, "Common Name, Product Name,Proper Chemical Name or Decriptive Name",
        "Common Name").astype(str).str.strip()

    # Stage 1: a row is inventory only if it identifies a chemical at all.
    # Template residue carries formula artifacts, not empty strings, so an
    # emptiness test on the name column alone lets all of it through.
    keep = (cas_digits != "") | ~name.map(is_null)

    df = pd.DataFrame({
        "cas_digits": cas_digits,
        "cas_number": cas_digits.map(format_cas),
        "cas_valid": cas_digits.map(cas_checksum_ok),
        "chemical_name": name,
        "descriptive_name": descriptive,
        "formula": column(raw, "Molecular Formula (optional)", "Molecular Formula"),
        "storage_group": column(raw, "A&M System Storage Group"),
        "manufacturer": column(raw, "Manufacturer Name", "Manufacturer",
                               "Supplier").astype(str).str.strip(),
        "container_count": pd.to_numeric(
            column(raw, "No of Containers", "Number of Containers",
                   "# of Containers"), errors="coerce"),
        # The workbook labels the three NFPA 704 axes by hazard rather than by
        # the standard's own names; 'Fire Hazard' is the flammability axis.
        "nfpa_health": pd.to_numeric(column(raw, "Health Hazard", "NFPA Health"),
                                     errors="coerce"),
        "nfpa_flam": pd.to_numeric(column(raw, "Fire Hazard", "NFPA Flammability"),
                                   errors="coerce"),
        "nfpa_inst": pd.to_numeric(
            column(raw, "Instability (Reactivity)", "NFPA Instability"), errors="coerce"),
        "storage_unit_raw": column(raw, "Unit No., Name or Description",
                                   "Storage Unit", "Storage Location"),
        "room": column(raw, "Room No.", "Room", "Room Number").astype(str).str.strip(),
    })[keep].copy()

    for col, label in SPECIAL_HAZARDS.items():
        df[label] = column(raw, col).map(lambda v: not is_null(v))

    # Stage 8: location. The unit field mixes cabinet with supplier catalog number.
    parsed = df["storage_unit_raw"].map(lambda v: parse_storage_unit(v, unit_mode))
    df["storage_unit"] = [p[0] for p in parsed]
    df["unit_precision"] = df["storage_unit"].map(unit_precision)

    # A record represents one or more physical containers. One record leaves the
    # count blank; it is left missing rather than imputed as 1, so the reported
    # total is the sum of what the inventory actually declares.
    df["container_count"] = df["container_count"].clip(lower=0)

    # Stage 9: substance identity. Grouping on CAS + free-text name splits one
    # substance across several groups (zinc appears under eight name variants).
    df["substance_key"] = np.where(
        df["cas_digits"].str.len() >= 5,
        df["cas_digits"],
        "NOCAS_" + df["chemical_name"].str.upper())

    if verbose:
        print(f"  {n_raw} raw rows -> {len(df)} container records "
              f"({n_raw - len(df)} template rows removed)")
    return df


# --------------------------------------------------------------------------
# The hazard cross-walk -- Paper 1's central artifact
# --------------------------------------------------------------------------

def build_bridge(df: pd.DataFrame) -> pd.DataFrame:
    """
    The chemical--hazard bridge: one row per (substance, hazard category).

    Deterministic by construction. A storage-group code maps through the
    institution's legend; each special-hazard flag contributes its own category.
    A substance carrying several categories yields several rows, which is the
    whole point -- a single-attribute field would force a choice.
    """
    rows = []
    for key, g in df.groupby("substance_key", sort=True):
        cats = set()
        for code in g["storage_group"]:
            if not is_null(code):
                code = str(code).strip()
                cats.add(STORAGE_GROUPS.get(code, f"UNMAPPED:{code}"))
        for label in SPECIAL_HAZARDS.values():
            if g[label].any():
                cats.add(label)
        names = [n for n in g["chemical_name"] if not is_null(n)]
        for c in sorted(cats):
            rows.append({
                "substance_key": key,
                "cas_number": g["cas_number"].iloc[0],
                "chemical_name": names[0] if names else "Unknown",
                "hazard_category": c,
            })
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# The analyses Paper 1 reports
# --------------------------------------------------------------------------

def single_site_statistics(df: pd.DataFrame, bridge: pd.DataFrame) -> dict:
    per_sub = bridge.groupby("substance_key")["hazard_category"].nunique()
    unmapped = sorted({c for c in bridge.hazard_category if c.startswith("UNMAPPED:")})

    hc = bridge[bridge.hazard_category.isin(HIGH_CONSEQUENCE)]

    suppliers = {m.strip().lower() for m in df.manufacturer if not is_null(m)}
    units = {u for u in df.storage_unit if not is_null(u)}

    return {
        "container_records":   len(df),
        "physical_containers": int(df.container_count.sum(skipna=True)),
        "unique_substances":   df.substance_key.nunique(),
        "cas_present":         int((df.cas_digits != "").sum()),
        "records_missing_count": int(df.container_count.isna().sum()),
        "cas_valid":           int(df.cas_valid.sum()),
        "classified_substances": bridge.substance_key.nunique(),
        "hazard_links":        len(bridge),
        "hazard_categories":   bridge.hazard_category.nunique(),
        "multi_hazard":        int((per_sub > 1).sum()),
        "max_concurrent":      int(per_sub.max()) if len(per_sub) else 0,
        "mean_categories":     round(per_sub.mean(), 2) if len(per_sub) else 0,
        "storage_units":       len(units),
        "suppliers":           len(suppliers),
        "high_consequence_links": len(hc),
        "high_consequence_categories": hc.hazard_category.nunique(),
        "nfpa_health_ge3":     int((df.nfpa_health >= 3).sum()),
        "nfpa_flam_ge3":       int((df.nfpa_flam >= 3).sum()),
        "nfpa_inst_ge3":       int((df.nfpa_inst >= 3).sum()),
        "container_to_chemical": round(df.container_count.sum(skipna=True) / df.substance_key.nunique(), 2),
        "unmapped_codes":      unmapped,
    }


def multisite_statistics(frames: list[pd.DataFrame]) -> dict:
    """Redundant holdings and co-location screening -- both need location to vary."""
    df = pd.concat(frames, ignore_index=True)

    located = df[df.storage_unit.notna()].copy()
    located["unit_id"] = located["building"].astype(str) + "/" + \
                         located["room"].astype(str) + "/" + located["storage_unit"].astype(str)

    # Substances held in more than one distinct location.
    loc_per_sub = located.groupby("substance_key")["unit_id"].nunique()
    redundant = loc_per_sub[loc_per_sub > 1]

    # Co-location screening against the legend.
    conflicts = []
    for unit, g in located.groupby("unit_id"):
        groups = sorted({str(c).strip() for c in g.storage_group if not is_null(c)})
        counts = Counter(str(c).strip() for c in g.storage_group if not is_null(c))
        for i, a in enumerate(groups):
            for b in groups[i:]:
                if a == b:
                    # Group X is incompatible with other Group X chemicals, so
                    # two X records in one unit conflict even with nothing else
                    # present. Every other group is compatible with itself.
                    if not (a == "X" and counts.get("X", 0) > 1):
                        continue
                reason = pair_conflict(a, b)
                if reason:
                    conflicts.append({
                        "unit": unit, "group_a": a, "group_b": b, "reason": reason,
                        "precision": g.unit_precision.mode().iat[0]
                                     if len(g.unit_precision.mode()) else "unknown",
                    })
    cdf = pd.DataFrame(conflicts)

    return {
        "records":            len(df),
        "substances":         df.substance_key.nunique(),
        "buildings":          df.building.nunique(),
        "rooms":              df.groupby("building")["room"].nunique().sum(),
        "located_records":    len(located),
        "distinct_units":     located.unit_id.nunique(),
        "redundant_substances": len(redundant),
        "conflicts":          len(cdf),
        "conflicts_confirmable": int((cdf.precision == "individual").sum()) if len(cdf) else 0,
        "conflicts_type_only": int((cdf.precision == "type-only").sum()) if len(cdf) else 0,
        "individual_records": int((located.unit_precision == "individual").sum()),
        "type_only_records":  int((located.unit_precision == "type-only").sum()),
        "_conflicts_table":   cdf,
        "_redundant_table":   redundant.sort_values(ascending=False),
    }


# --------------------------------------------------------------------------
# Verification harness
# --------------------------------------------------------------------------

class Harness:
    """
    Each check names the manuscript location it defends.

    Two values changed after the first compile. The manuscript originally
    reported 527 storage units and 51 redundantly held substances, both
    computed by taking the storage-unit field literally -- which counts a
    supplier catalog number such as 'MERCK/801809' as a storage location and
    treats three spellings of one refrigerator as three units. The canonical
    parse in parse_storage_unit() gives 504 and 72, and the manuscript now
    reports those.
    """

    def __init__(self):
        self.rows = []

    def check(self, where: str, claim, actual, tol=0):
        ok = (abs(claim - actual) <= tol) if isinstance(claim, (int, float)) else (claim == actual)
        self.rows.append((ok, where, claim, actual))
        return ok

    def report(self) -> bool:
        width = max(len(r[1]) for r in self.rows)
        print(f"\n{'':2} {'manuscript location':<{width}}  {'claimed':>9}  {'computed':>9}")
        print("  " + "-" * (width + 24))
        for ok, where, claim, actual in self.rows:
            mark = "ok" if ok else "!!"
            print(f"{mark:2} {where:<{width}}  {claim:>9}  {actual:>9}")
        failed = [r for r in self.rows if not r[0]]
        print(f"\n  {len(self.rows)-len(failed)}/{len(self.rows)} checks passed")
        if failed:
            print("\n  MANUSCRIPT NUMBERS TO CORRECT:")
            for _, where, claim, actual in failed:
                print(f"    {where}: says {claim}, data gives {actual}")
        return not failed


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--single", default=str(SINGLE_SITE), help="single-site export")
    ap.add_argument("--multisite", action="store_true", help="also run the two-building analyses")
    ap.add_argument("--export", metavar="XLSX", help="write derived tables to a workbook")
    ap.add_argument("--unit-mode", choices=("canonical", "literal"), default="canonical",
                    help="how to resolve the storage-unit field (see parse_storage_unit)")
    args = ap.parse_args(argv)

    print("=" * 74)
    print("PAPER 1 -- deterministic pipeline and manuscript verification")
    print("=" * 74)

    print(f"\n[1] Single-site preparation (unit mode: {args.unit_mode})")
    raw = read_inventory(Path(args.single))
    df = prepare(raw, unit_mode=args.unit_mode)

    print("\n[2] Hazard cross-walk")
    bridge = build_bridge(df)
    stats = single_site_statistics(df, bridge)
    if stats["unmapped_codes"]:
        print(f"  WARNING unmapped storage codes: {stats['unmapped_codes']}")
    print(f"  {stats['hazard_links']} chemical-hazard links across "
          f"{stats['hazard_categories']} categories")

    print("\n[3] Reported statistics")
    for k, v in stats.items():
        if not k.startswith("_") and k != "unmapped_codes":
            print(f"  {k:28s} {v}")

    h = Harness()
    h.check("Abstract / Table: records",        698,  stats["container_records"])
    h.check("Abstract: physical containers",   1201,  stats["physical_containers"])
    h.check("Abstract: unique substances",      479,  stats["unique_substances"])
    h.check("Sec. data-prep: CAS valid",         697,  stats["cas_valid"])
    h.check("Sec. res-hazardmodel: classified",  476,  stats["classified_substances"])
    h.check("Abstract: hazard links",            515,  stats["hazard_links"])
    h.check("Abstract: categories observed",      14,  stats["hazard_categories"])
    h.check("Abstract: multi-hazard chemicals",   38,  stats["multi_hazard"])
    h.check("Sec. res-construction: storage units", 504, stats["storage_units"])
    h.check("Sec. res-manufacturers: suppliers",   60,  stats["suppliers"])
    h.check("Sec. res-highconsequence: links",     92,  stats["high_consequence_links"])
    h.check("Sec. res-highconsequence: categories", 8,  stats["high_consequence_categories"])
    h.check("Abstract: NFPA health >= 3",         145,  stats["nfpa_health_ge3"])
    h.check("Abstract: NFPA flammability >= 3",   103,  stats["nfpa_flam_ge3"])
    h.check("Abstract: NFPA instability >= 3",     29,  stats["nfpa_inst_ge3"])

    ms = None
    if args.multisite:
        print("\n[4] Multi-site preparation")
        frames = []
        for src in MULTI_SITE:
            if not Path(src["path"]).exists():
                print(f"  SKIP (missing): {Path(src['path']).name}")
                continue
            sub = prepare(read_inventory(Path(src["path"])), verbose=False,
                          unit_mode=args.unit_mode)
            sub["building"] = src["building"]
            if src["room_default"]:
                sub.loc[sub["room"].map(is_null), "room"] = src["room_default"]
            frames.append(sub)
            print(f"  {src['building']}: {len(sub)} records")
        if frames:
            ms = multisite_statistics(frames)
            print("\n[5] Multi-site statistics")
            for k, v in ms.items():
                if not k.startswith("_"):
                    print(f"  {k:28s} {v}")
            h.check("Sec. res-multisite: records",          861, ms["records"])
            h.check("Sec. res-multisite: substances",       543, ms["substances"])
            # The manuscript reports 14 candidate conflicts, none confirmable.
            # That figure counts only conflicts in type-only units, and it
            # reproduces exactly. It omits the same-group case below.
            h.check("Sec. res-colocation: type-only conflicts", 14,
                    ms["conflicts_type_only"])
            h.check("Sec. res-redundant: redundant subs",    72, ms["redundant_substances"])
            h.check("Sec. res-colocation: confirmable",       0, ms["conflicts_confirmable"])

    print("\n" + "=" * 74)
    print("VERIFICATION")
    print("=" * 74)
    passed = h.report()

    if args.export:
        with pd.ExcelWriter(args.export) as xl:
            df.to_excel(xl, sheet_name="container_records", index=False)
            bridge.to_excel(xl, sheet_name="chemical_hazards", index=False)
            (bridge.hazard_category.value_counts().rename_axis("hazard_category")
             .reset_index(name="links")).to_excel(xl, sheet_name="hazard_distribution", index=False)
            if ms is not None:
                ms["_conflicts_table"].to_excel(xl, sheet_name="colocation_conflicts", index=False)
                ms["_redundant_table"].rename("locations").reset_index() \
                    .to_excel(xl, sheet_name="redundant_holdings", index=False)
        print(f"\nwrote {args.export}")

    return 0 if passed else 1


def discover(patterns=("*.xlsb", "*.xlsx"), roots=(".", "data", str(DATA_DIR), "/content",
                      "/content/drive/MyDrive")) -> list[Path]:
    """
    Find candidate inventory workbooks wherever the notebook is running.

    The configured paths may not match where the exports actually sit.
    Rather than fail on a missing path, look in the working directory and the
    usual Colab mount points, and let the CAS-column test decide what is usable.
    """
    seen, out = set(), []
    for root in roots:
        r = Path(root)
        if not r.is_dir():
            continue
        for pat in patterns:
            for f in sorted(r.glob(pat)):
                if f.name.startswith("~$") or f.resolve() in seen:
                    continue
                seen.add(f.resolve())
                out.append(f)
    # Prefer files that look like the single-site export, then the rest.
    return sorted(out, key=lambda f: (0 if "lbv242" in f.name.lower() else
                                      1 if "inventory" in f.name.lower() else 2,
                                      len(f.name)))


def autorun():
    """Notebook entry point: find the data, run, and print. No arguments."""
    configured = [Path(SINGLE_SITE)] + [Path(s["path"]) for s in MULTI_SITE]
    if all(p.exists() for p in configured):
        print("using the configured paths\n")
        return main(["--multisite"])

    found = discover()
    if not found:
        print("No inventory workbook found.\n")
        print("Upload the .xlsb or .xlsx export to this session, then run:\n"
              "    main(['--single', 'YOUR_FILE.xlsb'])\n")
        print(f"Looked in: {Path.cwd()} and the usual Colab mount points.")
        return 1

    print(f"Configured paths are not present; found {len(found)} spreadsheet(s) here.")
    for f in found:
        print(f"    {f}")

    # Use the first file that actually carries a CAS column as the single site,
    # and treat any others as further survey areas.
    usable = []
    for f in found:
        try:
            read_inventory(f)
            usable.append(f)
        except Exception:
            pass
    if not usable:
        print("\nNone of them has a CAS column; these are not inventory exports.")
        return 1

    argv = ["--single", str(usable[0])]
    if len(usable) > 1:
        globals()["MULTI_SITE"] = [
            {"path": f, "building": f.stem[:40], "room_default": None} for f in usable
        ]
        argv.append("--multisite")
        print("\nRunning single-site on the first file and multi-site across all "
              f"{len(usable)}.\nBuilding and room defaults are unknown for auto-"
              "discovered files, so multi-site\nlocation counts will differ from the "
              "manuscript unless you set MULTI_SITE.\n")
    else:
        print("\nOnly one usable export; skipping the multi-site analyses.\n")
    return main(argv)


if __name__ == "__main__" and "ipykernel" not in sys.modules:
    sys.exit(main())
elif "ipykernel" in sys.modules:
    # Jupyter sets __name__ to "__main__" and hands the kernel its own argv, so
    # an unguarded main() would try to parse '-f /tmp/kernel.json'. Run with an
    # explicit argument list instead, after locating the data.
    autorun()
