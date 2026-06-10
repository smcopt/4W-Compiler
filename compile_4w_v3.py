"""
Gaza Site Management Cluster — 4W Compilation & Validation Pipeline (v3)
========================================================================
March + April 2026 reporting cycle only (Jan/Feb deliberately excluded).

v3 methodology (locked-in cluster decisions):
  * COD-compliant INTEGER pcodes sourced from the masterlist
        governorate_pcode  <- masterlist adm2  (255/260/265/270/275)
        neighborhood_pcode <- masterlist adm4  (7-8 digit ints)
  * Indicator counting classification (every Indicator_Code mapped):
        PEOPLE_CAPPED (8)      reach capped at masterlist Total Inv (no tolerance)
        PEOPLE_CUMULATIVE (2)  IND_021 Cash-for-Work, IND_026 Training (no cap)
        SITE_CLAMPED (4)       unit "# of sites": total>1 clamped to 1
        NOT_PEOPLE (11)        days/reports/meetings/products/bags/tools/etc.
  * Corrections recorded with total_count_original + was_corrected flag
  * Reading B cluster reach:
        per site x month  -> MAX across overlapping people-indicators
        across months     -> MAX for capped (same population)
                             SUM for cumulative (fresh cohorts each month)
  * Multi-month container files (e.g. 202603_202604_*, or "April" files that
    still contain March rows) -> month taken from the in-file Reporting Month
    column per row; single-month files -> filename month wins (stale-copy guard)
  * Global content de-duplication collapses redundant re-submissions
    (the separate 202604_ACTED is a subset of the combined file; the two IOM
    files are byte-identical).
"""
from __future__ import annotations
import argparse, logging, re, sys
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
import unicodedata
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
C = dict(
    month="reporting_month", org="organization_name", prog_org="program_organization",
    gov="governorate", nbhd="neighborhood", site_text="site_name_raw",
    status="activity_status", primary="primary_activity", sub="sub_activity",
    indicator="indicator", spec="specification", details="activity_details",
    unit="unit", total="total_count", male="male", female="female",
    boys="boys_under18", girls="girls_under18", men="men_18_59", women="women_18_59",
    elder_m="elderly_male_60plus", elder_f="elderly_female_60plus",
    pwd_m="male_pwd", pwd_f="female_pwd", remarks="additional_remarks",
)
DEMO_COLS = [C["boys"], C["girls"], C["men"], C["women"], C["elder_m"], C["elder_f"]]
PWD_COLS = [C["pwd_m"], C["pwd_f"]]
SEX_COLS = [C["male"], C["female"]]
PEOPLE_UNITS = {"# of individuals", "# of participants", "# of people"}

# Indicator counting classification by SMC_IND code -------------------------
PEOPLE_CAPPED = {"SMC_IND_001","SMC_IND_004","SMC_IND_005","SMC_IND_015",
                 "SMC_IND_017","SMC_IND_020","SMC_IND_022","SMC_IND_024"}
PEOPLE_CUMULATIVE = {"SMC_IND_021","SMC_IND_026"}
SITE_CLAMPED = {"SMC_IND_006","SMC_IND_007","SMC_IND_010","SMC_IND_011"}
# everything else -> NOT_PEOPLE

GOV_CANON = {
    "khan yunis":"Khan Younis","khan younis":"Khan Younis","khanyunis":"Khan Younis",
    "deir al balah":"Deir Al-Balah","deir al-balah":"Deir Al-Balah","deir albalah":"Deir Al-Balah",
    "deir-al-balah":"Deir Al-Balah","deir el balah":"Deir Al-Balah",
    "north gaza":"North Gaza","northern gaza":"North Gaza",
    "gaza":"Gaza","gaza city":"Gaza","rafah":"Rafah",
}
PREFIX_TO_GOV = {"DEB":"Deir Al-Balah","GZA":"Gaza","KYS":"Khan Younis",
                 "NGZ":"North Gaza","RFH":"Rafah"}
# Integer ADM1 COD pcodes (mirror masterlist adm2)
GOV_PCODE = {"North Gaza":255,"Gaza":260,"Deir Al-Balah":265,"Khan Younis":270,"Rafah":275}

SITE_CODE_RE = re.compile(r"\b([A-Z]{3}\d{4})\b")
FILENAME_RE = re.compile(r"^((?:\d{6}_)+)Activities_([A-Za-z0-9\-_]+)\.xlsx$", re.IGNORECASE)
MONTH_NAMES = {1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",
               7:"July",8:"August",9:"September",10:"October",11:"November",12:"December"}
KEEP_MONTHS = {"March","April"}          # this cycle only
KEEP_PERIODS = {202603, 202604}

COLUMN_ALIASES = {
    "Reporting Month":C["month"], "Organization Name":C["org"],
    "Program / Coordinating / Donor Organization":C["prog_org"],
    "Governorate":C["gov"], "Governorate\n(Select from Dropdown)":C["gov"],
    "Neighborhood":C["nbhd"], "Neighborhood\n(Select from Dropdown)":C["nbhd"],
    "Site Name":C["site_text"], "Site Name\n(Select from Dropdown)":C["site_text"],
    "Site Name\n(Select from Dropdown, or type name and code if not found)":C["site_text"],
    "Activity Status":C["status"], "Activity\nStatus\n(Select from Dropdown)":C["status"],
    "Primary Activities":C["primary"], "Primary Activities\n(Select From Dropdown)":C["primary"],
    "Sub Activity":C["sub"], "Sub Activity\n(Select from Dropdown)":C["sub"],
    "Indicator":C["indicator"], "Indicator\n(Select from Dropdown)":C["indicator"],
    "Specification":C["spec"], "Specification\n(Select from Dropdown)":C["spec"],
    "Activity Details":C["details"],
    "Activity Details\n(Provide additional brief remarks on the activity, if any)":C["details"],
    "Units":C["unit"], "Units\n(Auto Calculated, Do not modify)":C["unit"],
    "Total":C["total"], "Total\n Planned/Reached\nCount\n":C["total"],
    "Total\nPlanned/Reached\nCount\n":C["total"], "Total Planned/Reached Count":C["total"],
    "Male":C["male"], "Female":C["female"],
    "Boys (<18)":C["boys"], "Girls (<18)":C["girls"],
    "Men (18-59)":C["men"], "Women (18-59)":C["women"],
    "Elderly Male (≥60)":C["elder_m"], "Elderly Female (≥60)":C["elder_f"],
    "Male-PWD":C["pwd_m"], "Female-PWD":C["pwd_f"],
    "Additional Remarks":C["remarks"], "Additional remarks":C["remarks"],
}

# ---------------------------------------------------------------------------
def setup_logger(out: Path):
    out.mkdir(parents=True, exist_ok=True)
    p = out / "pipeline_run_log.txt"; p.write_text("")
    lg = logging.getLogger("smc4w"); lg.setLevel(logging.INFO); lg.handlers.clear()
    fh, sh = logging.FileHandler(p, mode="w"), logging.StreamHandler(sys.stdout)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    for h in (fh, sh): h.setFormatter(fmt); lg.addHandler(h)
    return lg

@dataclass
class Issue:
    file:str; partner_from_filename:str; month_from_filename:str; row:object
    organization:str; site:str; indicator:str; error_type:str; severity:str; description:str

@dataclass
class IssueLog:
    issues:list = field(default_factory=list)
    def add(self, **k):
        for f in ("organization","site","indicator"): k.setdefault(f,"")
        k.setdefault("row","-"); self.issues.append(Issue(**k))
    def to_df(self):
        if not self.issues:
            return pd.DataFrame(columns=[f for f in Issue.__dataclass_fields__])
        return pd.DataFrame([i.__dict__ for i in self.issues])

# ---------------------------------------------------------------------------
def parse_filename(name):
    m = FILENAME_RE.match(name)
    if not m: return None
    yms = re.findall(r"\d{6}", m.group(1)); partner = m.group(2).upper()
    months = []
    for ym in yms:
        y, mo = int(ym[:4]), int(ym[4:])
        if 1 <= mo <= 12: months.append((MONTH_NAMES[mo], y, mo))
    return (partner, months) if months else None

def normalise_gov(v):
    if pd.isna(v): return None
    return GOV_CANON.get(re.sub(r"[-_\s]+"," ",str(v).strip().lower()), str(v).strip())

def extract_site_code(t):
    if pd.isna(t): return None
    m = SITE_CODE_RE.search(str(t).upper()); return m.group(1) if m else None

def _normalise_columns(df):
    ren = {}
    for c in df.columns:
        key = str(c).strip()
        if key in COLUMN_ALIASES: ren[c] = COLUMN_ALIASES[key]
        else:
            coll = re.sub(r"\s+"," ",key)
            matched = False
            for s,d in COLUMN_ALIASES.items():
                if re.sub(r"\s+"," ",s) == coll: ren[c]=d; matched=True; break
            # Robust fallback for the critical month key: some partner templates
            # ship variant headers like "Reporting Month2" (seen in UNRWA's April
            # workbook). Treat any header that begins with "reporting month" as the
            # month column so the file is not silently skipped.
            if not matched and re.sub(r"\s+"," ",key).lower().startswith("reporting month"):
                ren[c] = C["month"]
    return df.rename(columns=ren)

def _trim(df):
    need = [c for c in (C["org"],C["site_text"],C["total"]) if c in df.columns]
    if not need: return df.head(0)
    nb = df[need].apply(lambda col: col.notna() & (col.astype(str).str.strip()!=""))
    return df.loc[nb.any(axis=1)].reset_index(drop=True)

def _read_excel(path, sheet, header):
    """openpyxl first; fall back to calamine for files with malformed styles.xml."""
    try:
        return pd.read_excel(path, sheet_name=sheet, header=header, engine="openpyxl")
    except Exception:
        return pd.read_excel(path, sheet_name=sheet, header=header, engine="calamine")

def read_partner_file(path, log, partner_fn, month_label, lg):
    try:
        try: xl = pd.ExcelFile(path, engine="openpyxl")
        except Exception: xl = pd.ExcelFile(path, engine="calamine")
    except Exception as e:
        log.add(file=path.name,partner_from_filename=partner_fn,month_from_filename=month_label,
                error_type="UNREADABLE_FILE",severity="ERROR",description=f"Cannot open: {e}")
        return None
    sheets = xl.sheet_names
    if "ActivityReporting (Data Entry)" in sheets:
        df = _normalise_columns(_read_excel(path,"ActivityReporting (Data Entry)",2))
        if C["month"] in df.columns and C["org"] in df.columns: return _trim(df)
    for sh in sheets:
        try: df = _normalise_columns(_read_excel(path,sh,0))
        except Exception: continue
        if C["month"] in df.columns and C["org"] in df.columns and C["indicator"] in df.columns:
            return _trim(df)
    log.add(file=path.name,partner_from_filename=partner_fn,month_from_filename=month_label,
            error_type="UNRECOGNISED_FORMAT",severity="ERROR",description="No 4W headers found.")
    return None

def build_indicator_alias(submitted, master, cutoff=0.85):
    ml = master.dropna().astype(str).str.strip().unique().tolist(); ms = set(ml)
    out = {}
    for raw in submitted.dropna().astype(str).str.strip().unique():
        if raw in ms: out[raw]=raw; continue
        best, bm = 0.0, None
        for m in ml:
            s = SequenceMatcher(None, raw.lower(), m.lower()).ratio()
            if s > best: best, bm = s, m
        out[raw] = bm if best >= cutoff else None
    return out

def norm_name(s):
    """Aggressively normalise a site name for matching: strip diacritics,
    drop generic tokens (site/camp/...) and Arabic article prefixes, keep
    only alphanumerics."""
    if pd.isna(s): return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode().lower()
    s = re.sub(r"\b(site|camp|shelter|centre|center|gathering|point|school|area)\b", " ", s)
    s = re.sub(r"\b(al|el|abu|um)\b", " ", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def build_name_index(mst):
    """Two indexes: exact lowercased name -> id, and normalised name -> id."""
    exact, normd = {}, {}
    for _, r in mst.iterrows():
        sid = str(r["Site ID"]).strip().upper()
        for col in ("Site Name","Alternative Name"):
            if col in mst.columns and pd.notna(r.get(col)):
                e = str(r[col]).strip().lower()
                if e and e not in exact: exact[e] = sid
                n = norm_name(r[col])
                if n and n not in normd: normd[n] = sid
    return exact, normd


def norm_nbhd(s):
    """Normalise a neighborhood name for COD matching: strip diacritics,
    apostrophes and parentheses (keeping the word inside), keep all remaining
    words. Direction qualifiers (north/south/east/west) are KEPT — they
    distinguish genuinely different adm4 polygons (e.g. Deir Al Balah east vs
    south), so dropping them would conflate distinct neighborhoods."""
    if pd.isna(s): return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode().lower()
    s = s.replace("\u2019", " ").replace("'", " ")
    s = re.sub(r"[()]", " ", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def build_cod_neighborhood_index(cod_path):
    """Read the COD neighborhood sheet and return a governorate-scoped lookup
    {governorate: {normalised_name: {adm4_pcode, ...}}} plus the key list per
    governorate (for fuzzy fallback). Names collide across governorates AND a
    few collide within one (two 'Ar Rasheed' polygons in Deir Al-Balah), so the
    value is a SET of pcodes; a set with >1 member is treated as ambiguous.
    Returns (index, keys_by_gov) or ({}, {}) if the COD can't be read."""
    try:
        xl = pd.ExcelFile(cod_path)
    except Exception:
        return {}, {}
    sheet = next((s for s in xl.sheet_names if "neighb" in s.lower()), None)
    if sheet is None:
        return {}, {}
    cod = pd.read_excel(cod_path, sheet_name=sheet)
    name_col = next((c for c in cod.columns if str(c).lower().startswith("neighb")), None)
    pcode_col = next((c for c in cod.columns if "pcode_neig" in str(c).lower()), None)
    gov_col = next((c for c in cod.columns if str(c).lower().startswith("governorat")), None)
    if not (name_col and pcode_col and gov_col):
        return {}, {}
    idx = {}
    for _, r in cod.iterrows():
        g = normalise_gov(r[gov_col]); k = norm_nbhd(r[name_col])
        if not (g and k) or pd.isna(r[pcode_col]):
            continue
        idx.setdefault(g, {}).setdefault(k, set()).add(int(r[pcode_col]))
    keys = {g: list(d) for g, d in idx.items()}
    return idx, keys


def build_masterlist_nbhd_index(mst):
    """{governorate: {normalised_name: {adm4_pcode, ...}}} from the masterlist's
    own Neighborhood + adm4 columns. Used only to break COD ambiguities: among
    several COD candidates for one name, prefer the pcode the masterlist
    actually uses operationally."""
    idx = {}
    if not {"Neighborhood","adm 4","Governorate"}.issubset(mst.columns):
        return idx
    for _, r in mst.iterrows():
        g = normalise_gov(r["Governorate"]); k = norm_nbhd(r["Neighborhood"])
        if not (g and k) or pd.isna(r["adm 4"]):
            continue
        idx.setdefault(g, {}).setdefault(k, set()).add(int(r["adm 4"]))
    return idx


def resolve_nbhd_pcode(gov, name, cod_idx, cod_keys, ml_idx=None):
    """Look up a neighborhood's adm4 pcode in the COD by name, scoped to the
    governorate, preserving direction qualifiers. Where one name maps to
    several COD pcodes, break the tie with the masterlist's own usage; if it
    still can't be resolved to a single code, return ambiguous rather than
    guess. Returns (pcode|None, method, score). Methods: cod_exact,
    cod_exact_ml (tie broken via masterlist), cod_fuzzy, cod_fuzzy_ml,
    ambiguous, miss."""
    ml_idx = ml_idx or {}
    gi = cod_idx.get(gov)
    if not gi:
        return None, "miss", 0.0
    q = norm_nbhd(name)
    if not q:
        return None, "miss", 0.0

    def pick(pcs, method, score):
        if len(pcs) == 1:
            return next(iter(pcs)), method, score
        inter = pcs & ml_idx.get(gov, {}).get(q, set())
        if len(inter) == 1:
            return next(iter(inter)), method + "_ml", score
        return None, "ambiguous", score

    if q in gi:
        return pick(gi[q], "cod_exact", 1.0)
    # conservative fuzzy — 0.93 floor keeps direction variants apart
    # ('deir al balah south' vs 'east' scores ~0.86, below the floor)
    best, bm = 0.0, None
    for k in cod_keys.get(gov, ()):
        s = SequenceMatcher(None, q, k).ratio()
        if s > best: best, bm = s, k
    if bm and best >= 0.93:
        return pick(gi[bm], "cod_fuzzy", round(best, 3))
    return None, "miss", round(best, 3)

def resolve_site(text, code, exact_idx, norm_idx, norm_keys, valid):
    """Return (site_id, method, score). Methods: code_match, name_exact,
    name_norm, name_fuzzy, unresolved."""
    if code and code in valid:
        return code, "code_match", 1.0
    if pd.isna(text):
        return None, "unresolved", 0.0
    c = str(text).strip().lower()
    c = re.sub(r"\s*\(([a-z]{3}\d{4})\)\s*$","",c).strip()
    c = re.sub(r"^[a-z]{3}\d{4}\s*[-–]\s*","",c).strip()
    if c in exact_idx:
        return exact_idx[c], "name_exact", 1.0
    q = norm_name(text)
    if not q:
        return None, "unresolved", 0.0
    if q in norm_idx:
        return norm_idx[q], "name_norm", 1.0
    # conservative fuzzy: require >=0.92 ratio and a normalised key of >=6 chars
    if len(q) >= 6:
        best, bm = 0.0, None
        for k in norm_keys:
            s = SequenceMatcher(None, q, k).ratio()
            if s > best: best, bm = s, k
        if bm and best >= 0.92:
            return norm_idx[bm], "name_fuzzy", round(best,3)
    return None, "unresolved", 0.0

def coerce(df, cols):
    for c in cols:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors="coerce")

# ---------------------------------------------------------------------------
def validate_and_clean(raw, file, partner_fn, month_label, alias, ind_master,
                       mst, exact_idx, norm_idx, norm_keys, ind_class, site_pop, log):
    df = raw.copy()
    for col in C.values():
        if col not in df.columns: df[col] = pd.NA

    # exact dup within file
    dm = df.duplicated(keep="first"); n = int(dm.sum())
    if n: log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                  organization=partner_fn,error_type="EXACT_DUPLICATE",severity="WARNING",
                  description=f"{n} fully-duplicated rows removed")
    df = df.loc[~dm].reset_index(drop=True)

    # backfill org — fillna BEFORE string ops; .str.strip() reintroduces NaN
    # for genuine float-NaN cells (seen in UNRWA's April workbook, where 93 of
    # 116 Organization Name cells are blank), so normalise to "" up front.
    org = (df[C["org"]].fillna("").astype(str).str.strip()
                       .replace({"nan":"","NaN":"","None":""}))
    nb = int((org=="").sum())
    if nb:
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=partner_fn,error_type="MISSING_ORG_NAME",severity="WARNING",
                description=f"{nb} row(s) blank Organization Name -> defaulted to '{partner_fn}'")
        df.loc[org=="",C["org"]] = partner_fn

    df["governorate_clean"] = df[C["gov"]].apply(normalise_gov)
    df["indicator_clean"]   = df[C["indicator"]].astype(str).str.strip().map(alias)
    valid_ids = set(mst["Site ID"].dropna().astype(str).str.upper())
    df["site_code_extracted"] = df[C["site_text"]].apply(extract_site_code)
    res = df.apply(lambda r: resolve_site(r[C["site_text"]], r["site_code_extracted"],
                                          exact_idx, norm_idx, norm_keys, valid_ids), axis=1)
    df["site_code_clean"] = [x[0] for x in res]
    df["site_resolution_method"] = [x[1] for x in res]
    df["site_match_score"] = [x[2] for x in res]
    coerce(df,[C["total"]]+DEMO_COLS+PWD_COLS+SEX_COLS)

    # canonical indicator metadata
    look = ind_master.set_index("Indicators")[["Activity_Code","Indicator_Code","Unit",
                                                "Primary_Activity","Sub_Activity","Indicator Purpose"]]
    df = df.merge(look, how="left", left_on="indicator_clean", right_index=True)
    df = df.rename(columns={"Activity_Code":"activity_code","Indicator_Code":"indicator_code",
                            "Unit":"canonical_unit","Primary_Activity":"canonical_primary",
                            "Sub_Activity":"canonical_sub","Indicator Purpose":"indicator_purpose"})
    df["counting_class"] = df["indicator_code"].map(ind_class).fillna("NOT_PEOPLE")
    df["is_people_indicator"] = df["counting_class"].isin({"PEOPLE_CAPPED","PEOPLE_CUMULATIVE"})

    # ----- unknown indicator
    for _, r in df.loc[df["indicator_clean"].isna() & df[C["indicator"]].notna()].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r.get(C["indicator"],""))[:120],error_type="UNKNOWN_INDICATOR",
                severity="ERROR",description="Indicator not in Activity Index (85% cutoff)")
    # ----- status missing
    ns = df[C["status"]].isna() | (df[C["status"]].astype(str).str.strip()=="")
    for _, r in df.loc[ns].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                error_type="ACTIVITY_STATUS_MISSING",severity="WARNING",
                description="Activity Status blank")
    # ----- gov invalid
    vg = set(GOV_CANON.values())
    for _, r in df.loc[df["governorate_clean"].notna() & ~df["governorate_clean"].isin(vg)].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                error_type="INVALID_GOVERNORATE",severity="ERROR",
                description=f"Governorate '{r[C['gov']]}' not recognised")
    # ----- site resolution
    unres = df["site_code_clean"].isna() & df[C["site_text"]].notna()
    for _, r in df.loc[unres].iterrows():
        had = pd.notna(r["site_code_extracted"])
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                error_type=("UNKNOWN_SITE_CODE" if had else "UNRESOLVED_SITE"),severity="ERROR",
                description=("Code not in masterlist & name unmatched" if had
                             else "No site code & name unmatched"))
    for _, r in df.loc[df["site_resolution_method"].isin(["name_exact","name_norm"])].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                error_type="NO_SITE_CODE",severity="WARNING",
                description=f"Resolved by name to '{r['site_code_clean']}'; please include site code")
    for _, r in df.loc[df["site_resolution_method"]=="name_fuzzy"].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                error_type="NAME_FUZZY_MATCH",severity="WARNING",
                description=f"Fuzzy-matched '{r[C['site_text']]}' -> {r['site_code_clean']} "
                            f"(score {r['site_match_score']}); VERIFY before publishing")
    # ----- gov prefix mismatch
    df["prefix"] = df["site_code_clean"].str[:3]
    mis = (df["prefix"].isin(PREFIX_TO_GOV) & df["governorate_clean"].notna()
           & (df["prefix"].map(PREFIX_TO_GOV)!=df["governorate_clean"]))
    for _, r in df.loc[mis].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                error_type="GOV_PREFIX_MISMATCH",severity="WARNING",
                description=f"Prefix '{r['prefix']}'->{PREFIX_TO_GOV[r['prefix']]} != {r['governorate_clean']}")
    # ----- neighborhood blank
    bn = df[C["nbhd"]].isna() | (df[C["nbhd"]].astype(str).str.strip()=="")
    for _, r in df.loc[bn].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                error_type="MISSING_NEIGHBORHOOD",severity="WARNING",description="Neighborhood blank")

    # ===== counting corrections =====================================
    df["total_count_original"] = df[C["total"]]
    df["was_corrected"] = False; df["correction_type"] = ""

    # SITE_COUNT_CLAMPED
    sc = (df["counting_class"]=="SITE_CLAMPED") & df[C["total"]].notna() & (df[C["total"]]>1)
    for _, r in df.loc[sc].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r.get("indicator_clean",""))[:120],error_type="SITE_COUNT_CLAMPED",
                severity="INFO",description=f"# of sites total {int(r[C['total']])} clamped to 1")
    df.loc[sc, C["total"]] = 1
    df.loc[sc,"was_corrected"]=True; df.loc[sc,"correction_type"]="SITE_COUNT_CLAMPED"

    # REACH_CAPPED_AT_POPULATION (people-capped only; no tolerance)
    df["site_population"] = df["site_code_clean"].map(site_pop)
    cap = ((df["counting_class"]=="PEOPLE_CAPPED") & df[C["total"]].notna()
           & df["site_population"].notna() & (df[C["total"]]>df["site_population"]))
    for _, r in df.loc[cap].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r.get("indicator_clean",""))[:120],error_type="REACH_CAPPED_AT_POPULATION",
                severity="WARNING",description=f"Reach {int(r[C['total']])} capped to site pop {int(r['site_population'])}")
    df.loc[cap, C["total"]] = df.loc[cap,"site_population"]
    df.loc[cap,"was_corrected"]=True
    df.loc[cap,"correction_type"]=np.where(df.loc[cap,"correction_type"]=="",
                                           "REACH_CAPPED_AT_POPULATION",
                                           df.loc[cap,"correction_type"]+"+REACH_CAPPED")

    # ===== people / demo checks =====================================
    is_ppl = df["is_people_indicator"]
    has_tot = df[C["total"]].notna() & (df[C["total"]]>0)
    for _, r in df.loc[df[C["total"]].isna() & df["indicator_clean"].notna()].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r["indicator_clean"])[:120],error_type="MISSING_TOTAL",
                severity="ERROR",description="Total Planned/Reached blank")
    df["demo_sum"] = df[DEMO_COLS].sum(axis=1,min_count=1)
    df["demo_missing"] = df[DEMO_COLS].isna().all(axis=1)
    for _, r in df.loc[is_ppl & has_tot & df["demo_missing"]].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r["indicator_clean"])[:120],error_type="MISSING_DISAGGREGATION",
                severity="WARNING",description="People-indicator without age/sex split")
    md = is_ppl & has_tot & ~df["demo_missing"] & (df["demo_sum"].round()!=df["total_count_original"].round())
    for _, r in df.loc[md].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r["indicator_clean"])[:120],error_type="DEMO_SUM_MISMATCH",severity="ERROR",
                description=f"age/sex sum {int(r['demo_sum'])} != total {int(r['total_count_original'])}")
    df["sex_sum"] = df[SEX_COLS].sum(axis=1,min_count=1)
    df["sex_missing"] = df[SEX_COLS].isna().all(axis=1)
    ws = is_ppl & has_tot & ~df["sex_missing"] & (df["sex_sum"].round()!=df["total_count_original"].round())
    for _, r in df.loc[ws].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r["indicator_clean"])[:120],error_type="WRONG_SEX_TOTALS",severity="WARNING",
                description=f"M+F {int(r['sex_sum'])} != total {int(r['total_count_original'])}")
    df["pwd_sum"] = df[PWD_COLS].sum(axis=1,min_count=1)
    po = is_ppl & has_tot & df["pwd_sum"].notna() & (df["pwd_sum"]>df[C["total"]])
    for _, r in df.loc[po].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r["indicator_clean"])[:120],error_type="PWD_EXCEEDS_TOTAL",severity="ERROR",
                description=f"PWD {int(r['pwd_sum'])} > total {int(r[C['total']])}")
    # unit mismatch
    df["unit_clean"] = df[C["unit"]].astype(str).str.strip()
    ub = (df["indicator_clean"].notna() & df["canonical_unit"].notna()
          & (df["unit_clean"].str.replace(r"\s+"," ",regex=True)
             != df["canonical_unit"].astype(str).str.strip().str.replace(r"\s+"," ",regex=True))
          & ~df["unit_clean"].isin(["","nan"]))
    for _, r in df.loc[ub].iterrows():
        log.add(file=file,partner_from_filename=partner_fn,month_from_filename=month_label,
                organization=r.get(C["org"]) or partner_fn,site=r.get(C["site_text"],""),
                indicator=str(r["indicator_clean"])[:120],error_type="UNIT_MISMATCH",severity="WARNING",
                description=f"unit '{r['unit_clean']}' != canonical '{r['canonical_unit']}'")
    return df

# ---------------------------------------------------------------------------
def reading_b_reach(valid):
    """Reading B cluster reach. Returns (headline, per_gov_month_df, per_site_df)."""
    ppl = valid[valid["is_people_indicator"] & valid["total_count"].notna()].copy()
    if ppl.empty:
        return 0, pd.DataFrame(), pd.DataFrame()
    cap = ppl[ppl["counting_class"]=="PEOPLE_CAPPED"]
    cum = ppl[ppl["counting_class"]=="PEOPLE_CUMULATIVE"]
    # per site x month: MAX across overlapping indicators of that class
    cap_sm = cap.groupby(["site_id","governorate","reporting_period_id"],dropna=False)["total_count"].max().reset_index()
    cum_sm = cum.groupby(["site_id","governorate","reporting_period_id"],dropna=False)["total_count"].max().reset_index()
    # across months: MAX (capped) / SUM (cumulative)  -> per site
    cap_site = cap_sm.groupby(["site_id","governorate"],dropna=False)["total_count"].max().reset_index().rename(columns={"total_count":"capped"})
    cum_site = cum_sm.groupby(["site_id","governorate"],dropna=False)["total_count"].sum().reset_index().rename(columns={"total_count":"cumulative"})
    site = cap_site.merge(cum_site,on=["site_id","governorate"],how="outer").fillna(0)
    site["reach"] = site["capped"] + site["cumulative"]
    headline = int(site["reach"].sum())
    # per gov x month MONTHLY reach (MAX across ALL people indicators per site-month, sum sites)
    sm = ppl.groupby(["site_id","governorate","reporting_period_id"],dropna=False)["total_count"].max().reset_index()
    gov_month = sm.groupby(["governorate","reporting_period_id"],dropna=False).agg(
        reach=("total_count","sum"), sites_covered=("site_id","nunique")).reset_index()
    return headline, gov_month, site

def build_star(fact, mst, ind_master, partners_seen, ind_class, lg,
               cod_idx=None, cod_keys=None, log=None, ml_nbhd_idx=None):
    cod_idx = cod_idx or {}; cod_keys = cod_keys or {}; ml_nbhd_idx = ml_nbhd_idx or {}
    # join masterlist geography (integer pcodes)
    geo = mst.set_index("Site ID")[["adm2","adm 4","Neighborhood","Governorate"]]
    fact["governorate_pcode"] = fact["site_code_clean"].map(geo["adm2"])
    fact["neighborhood_pcode"] = fact["site_code_clean"].map(geo["adm 4"])
    fact["neighborhood_master"] = fact["site_code_clean"].map(geo["Neighborhood"])
    fact["gov_master"] = fact["site_code_clean"].map(geo["Governorate"]).apply(normalise_gov)
    # fall back gov pcode from gov name when site unresolved
    fact["governorate_pcode"] = fact["governorate_pcode"].fillna(
        fact["governorate_clean"].map(GOV_PCODE))

    # ------------------------------------------------------------------
    # Neighborhood pcode fallback via the COD.
    # When a row's site could not be resolved (so the masterlist adm4 join
    # gives nothing) but it still carries a neighborhood NAME, look the adm4
    # pcode up in the Common Operational Dataset by name, scoped to the
    # resolved governorate. This completes the geography for unmatched rows
    # without changing their validity or the reach figure.
    # ------------------------------------------------------------------
    if cod_idx:
        nbhd_name = fact["neighborhood_master"].fillna(fact[C["nbhd"]])
        gov_for_lookup = fact["gov_master"].fillna(fact["governorate_clean"])
        need = (fact["neighborhood_pcode"].isna()
                & nbhd_name.notna()
                & (nbhd_name.astype(str).str.strip() != "")
                & (nbhd_name.astype(str).str.strip().str.lower() != "(unspecified)"))
        filled = fuzzy = ambiguous = unresolved = 0
        # resolve once per distinct (gov, name) pair, then broadcast
        pairs = (pd.DataFrame({"g": gov_for_lookup[need], "n": nbhd_name[need]})
                 .drop_duplicates())
        resolved = {}
        for _, pr in pairs.iterrows():
            pc, meth, sc = resolve_nbhd_pcode(pr["g"], pr["n"], cod_idx, cod_keys, ml_nbhd_idx)
            resolved[(pr["g"], str(pr["n"]))] = (pc, meth, sc)
        for i in fact.index[need]:
            g = gov_for_lookup.at[i]; n = str(nbhd_name.at[i])
            pc, meth, sc = resolved.get((g, n), (None, "miss", 0.0))
            if pc is not None:
                fact.at[i, "neighborhood_pcode"] = pc
                filled += 1
                if meth.startswith("cod_fuzzy"):
                    fuzzy += 1
        if log is not None:
            # one representative log entry per distinct pair
            for (g, n), (pc, meth, sc) in resolved.items():
                if pc is not None:
                    is_fuzzy = meth.startswith("cod_fuzzy")
                    tie = meth.endswith("_ml")
                    log.add(file="(COD lookup)", partner_from_filename="—",
                            month_from_filename="Mar/Apr", organization="",
                            site=str(n),
                            error_type=("NEIGHBORHOOD_PCODE_FROM_COD_FUZZY" if is_fuzzy
                                        else "NEIGHBORHOOD_PCODE_FROM_COD"),
                            severity="INFO",
                            description=(f"Neighborhood '{n}' ({g}) had no site match; "
                                         f"adm4 pcode {pc} sourced from COD ({meth}"
                                         f"{'' if meth in ('cod_exact','cod_exact_ml') else f' {sc}'})"
                                         f"{'; tie broken via masterlist usage' if tie else ''}"))
                elif meth == "ambiguous":
                    ambiguous += 1
                    log.add(file="(COD lookup)", partner_from_filename="—",
                            month_from_filename="Mar/Apr", organization="",
                            site=str(n), error_type="NEIGHBORHOOD_PCODE_AMBIGUOUS",
                            severity="WARNING",
                            description=(f"Neighborhood '{n}' ({g}) maps to multiple COD "
                                         f"adm4 polygons and could not be disambiguated; "
                                         f"pcode left blank"))
                else:
                    unresolved += 1
                    log.add(file="(COD lookup)", partner_from_filename="—",
                            month_from_filename="Mar/Apr", organization="",
                            site=str(n), error_type="NEIGHBORHOOD_PCODE_UNRESOLVED",
                            severity="WARNING",
                            description=(f"No COD adm4 pcode for neighborhood '{n}' "
                                         f"in {g} (best score {sc})"))
        lg.info(f"COD neighborhood-pcode fallback: filled {filled} row(s) "
                f"({fuzzy} via fuzzy match) across {len(pairs)} distinct names "
                f"[ambiguous={ambiguous} unresolved={unresolved}]")

    fact_out = pd.DataFrame({
        "fact_id": np.arange(1,len(fact)+1),
        "source_file": fact["__source_file__"],
        "reporting_month": fact["month_final"],
        "reporting_year": fact["reporting_year"],
        "reporting_period_id": fact["reporting_period_id"],
        "organization": fact[C["org"]].fillna("(unknown)"),
        "implementing_via": fact[C["prog_org"]].fillna(""),
        "governorate": fact["gov_master"].fillna(fact["governorate_clean"]).fillna("(unknown)"),
        "governorate_pcode": fact["governorate_pcode"].astype("Int64"),
        "neighborhood": fact["neighborhood_master"].fillna(fact[C["nbhd"]]).fillna("(unspecified)"),
        "neighborhood_pcode": fact["neighborhood_pcode"].astype("Int64"),
        "site_id": fact["site_code_clean"].fillna(""),
        "site_name_raw": fact[C["site_text"]].fillna(""),
        "site_in_masterlist": fact["site_code_clean"].isin(set(mst["Site ID"].astype(str))),
        "activity_status": fact[C["status"]].fillna("(unspecified)"),
        "primary_activity": fact["canonical_primary"].fillna(fact[C["primary"]]).fillna(""),
        "sub_activity": fact["canonical_sub"].fillna(fact[C["sub"]]).fillna(""),
        "indicator": fact["indicator_clean"].fillna(fact[C["indicator"]]).fillna(""),
        "indicator_code": fact["indicator_code"].fillna(""),
        "activity_code": fact["activity_code"].fillna(""),
        "indicator_purpose": fact["indicator_purpose"].fillna("(unmapped)"),
        "framework_category": fact["indicator_purpose"].fillna("").apply(
            lambda s: "Flash Appeal" if "Flash" in str(s) else
            ("SMC Coordination" if "Coordination" in str(s) else "Other / Unmapped")),
        "counting_class": fact["counting_class"],
        "unit": fact["canonical_unit"].fillna(fact[C["unit"]]).fillna(""),
        "is_people_indicator": fact["is_people_indicator"],
        "total_count": fact["total_count"],
        "total_count_original": fact["total_count_original"],
        "was_corrected": fact["was_corrected"],
        "correction_type": fact["correction_type"],
        "male": fact[C["male"]], "female": fact[C["female"]],
        "boys_under18": fact[C["boys"]], "girls_under18": fact[C["girls"]],
        "men_18_59": fact[C["men"]], "women_18_59": fact[C["women"]],
        "elderly_male_60plus": fact[C["elder_m"]], "elderly_female_60plus": fact[C["elder_f"]],
        "male_pwd": fact[C["pwd_m"]], "female_pwd": fact[C["pwd_f"]],
        "pwd_total": fact[PWD_COLS].sum(axis=1,min_count=1),
        "activity_details": fact[C["details"]].fillna(""),
        "site_population_baseline": fact["site_population"],
        "is_valid": fact["__is_valid__"],
    })

    # dim_dates
    dd = (fact_out[["reporting_year","reporting_month","reporting_period_id"]]
          .drop_duplicates().sort_values("reporting_period_id").reset_index(drop=True))
    dd["month_short"] = dd["reporting_month"].str[:3]
    dd["month_year"] = dd["reporting_month"]+" "+dd["reporting_year"].astype(str)
    dd["period_start"] = pd.to_datetime(dd["reporting_period_id"].astype(str)+"01",format="%Y%m%d",errors="coerce")

    # dim_sites — masterlist passthrough + has_reported_activity + integer pcodes
    s = mst.copy()
    s = s.rename(columns={"Site ID":"site_id","Site Name":"site_name","Governorate":"governorate",
        "Neighborhood":"neighborhood","Site Status":"site_status","Managing Agency":"managing_agency",
        "FinalImplementingPartner":"implementing_partner","Total Inv":"population_total",
        "Total HH":"households_total","Displacement Type":"displacement_type",
        "Persons with disabilities":"pwd_baseline","adm2":"governorate_pcode","adm 4":"neighborhood_pcode"})
    s["governorate"] = s["governorate"].apply(normalise_gov)
    reported = set(fact_out.loc[fact_out["site_in_masterlist"] & fact_out["is_valid"],"site_id"])
    s["has_reported_activity"] = s["site_id"].astype(str).isin(reported)
    dim_sites = s[[c for c in ["site_id","site_name","governorate","governorate_pcode","neighborhood",
        "neighborhood_pcode","site_status","managing_agency","implementing_partner","population_total",
        "households_total","displacement_type","pwd_baseline","has_reported_activity",
        "Latitude","Longitude"] if c in s.columns]].copy()

    # dim_indicators — list_indicators_5w passthrough + class + framework
    di = ind_master.rename(columns={"Indicators":"indicator","Activity_Code":"activity_code",
        "Indicator_Code":"indicator_code","Unit":"unit","Primary_Activity":"primary_activity",
        "Sub_Activity":"sub_activity","Indicator Purpose":"indicator_purpose"}).copy()
    di["counting_class"] = di["indicator_code"].map(ind_class).fillna("NOT_PEOPLE")
    di["framework_category"] = di["indicator_purpose"].fillna("").apply(
        lambda s: "Flash Appeal" if "Flash" in str(s) else
        ("SMC Coordination" if "Coordination" in str(s) else "Other"))
    dim_indicators = di[["activity_code","indicator_code","indicator","primary_activity",
                         "sub_activity","unit","counting_class","indicator_purpose","framework_category"]]

    # dim_partners
    pset = set(fact_out["organization"].dropna().astype(str)) | set(partners_seen)
    for col in ("managing_agency","implementing_partner"):
        if col in dim_sites.columns: pset |= set(dim_sites[col].dropna().astype(str))
    for x in ("(unknown)","No SMA","",""): pset.discard(x)
    dim_partners = pd.DataFrame({"partner":sorted(pset)})
    dim_partners["reporting_in_4w"] = dim_partners["partner"].isin(set(fact_out["organization"]))
    dim_partners["managing_in_masterlist"] = dim_partners["partner"].isin(set(mst["Managing Agency"].dropna()))

    # dim_geography (integer pcodes)
    dim_geo = (dim_sites[["governorate","governorate_pcode","neighborhood","neighborhood_pcode"]]
               .dropna(subset=["governorate","neighborhood"]).drop_duplicates().reset_index(drop=True))

    # ===== aggregations (valid rows only) =====
    valid = fact_out[fact_out["is_valid"]].copy()
    headline, gov_month_reach, site_reach = reading_b_reach(valid)

    agg_pm = (valid.groupby(["organization","reporting_period_id","reporting_year","reporting_month"],
              as_index=False).agg(activities=("fact_id","count"),
              unique_sites=("site_id",lambda x:x[x!=""].nunique()),
              governorates=("governorate",pd.Series.nunique),
              total_reported=("total_count","sum")))

    gov_pop = (mst.assign(g=mst["Governorate"].apply(normalise_gov))
               .groupby("g",as_index=False).agg(masterlist_population=("Total Inv","sum"),
                                                 masterlist_sites=("Site ID","nunique")))
    agg_gm = gov_month_reach.merge(gov_pop,left_on="governorate",right_on="g",how="left").drop(columns=["g"])
    agg_gm["governorate_pcode"] = agg_gm["governorate"].map(GOV_PCODE).astype("Int64")
    agg_gm["coverage_rate"] = (agg_gm["sites_covered"]/agg_gm["masterlist_sites"]).round(3)
    agg_gm = agg_gm[["governorate","governorate_pcode","reporting_period_id","reach",
                     "sites_covered","masterlist_population","masterlist_sites","coverage_rate"]]

    agg_im = (valid.groupby(["indicator","indicator_code","activity_code","framework_category",
              "counting_class","reporting_period_id"],as_index=False)
              .agg(activities=("fact_id","count"),total=("total_count","sum"),
                   unique_sites=("site_id",lambda x:x[x!=""].nunique())))

    # agg_neighborhood_month (the one previously missing)
    ppl = valid[valid["is_people_indicator"] & valid["total_count"].notna()]
    nb_sm = ppl.groupby(["governorate","governorate_pcode","neighborhood","neighborhood_pcode",
                         "site_id","reporting_period_id"],dropna=False)["total_count"].max().reset_index()
    agg_nm = (nb_sm.groupby(["governorate","governorate_pcode","neighborhood","neighborhood_pcode",
              "reporting_period_id"],dropna=False)
              .agg(reach=("total_count","sum"),sites_covered=("site_id","nunique")).reset_index())

    # agg_coverage (site x month grid)
    grid = (dim_sites[["site_id","site_name","governorate","governorate_pcode","site_status","population_total"]]
            .merge(dd[["reporting_period_id","reporting_month","reporting_year"]],how="cross"))
    cov = (valid.groupby(["site_id","reporting_period_id"],dropna=False)
           .agg(partners_active=("organization",pd.Series.nunique),
                activities_reported=("fact_id","count")).reset_index())
    agg_cov = grid.merge(cov,on=["site_id","reporting_period_id"],how="left")
    agg_cov["partners_active"] = agg_cov["partners_active"].fillna(0).astype(int)
    agg_cov["activities_reported"] = agg_cov["activities_reported"].fillna(0).astype(int)
    agg_cov["is_covered"] = agg_cov["partners_active"]>0

    return {"fact_activities":fact_out,"dim_sites":dim_sites,"dim_partners":dim_partners,
            "dim_indicators":dim_indicators,"dim_geography":dim_geo,"dim_dates":dd,
            "agg_partner_month":agg_pm,"agg_governorate_month":agg_gm,"agg_indicator_month":agg_im,
            "agg_neighborhood_month":agg_nm,"agg_coverage":agg_cov}, headline, site_reach

# ---------------------------------------------------------------------------
def run(sub, mlp, indp, out, codp=None, keep_months=None):
    keep = set(keep_months) if keep_months else KEEP_MONTHS
    lg = setup_logger(out)
    lg.info("== Gaza SMC 4W pipeline v3 ==  months kept: " + ", ".join(sorted(keep)))
    mst = pd.read_csv(mlp,low_memory=False)
    mst["Site ID"] = mst["Site ID"].astype(str).str.strip().str.upper()
    ind_master = pd.read_csv(indp); ind_master["Indicators"] = ind_master["Indicators"].astype(str).str.strip()
    ind_class = {**{k:"PEOPLE_CAPPED" for k in PEOPLE_CAPPED},
                 **{k:"PEOPLE_CUMULATIVE" for k in PEOPLE_CUMULATIVE},
                 **{k:"SITE_CLAMPED" for k in SITE_CLAMPED}}
    name_idx_exact, name_idx_norm = build_name_index(mst)
    norm_keys = list(name_idx_norm)
    site_pop = mst.set_index("Site ID")["Total Inv"]
    # COD neighborhood index (adm4 pcode fallback for rows with no site match)
    cod_idx, cod_keys = ({}, {})
    if codp and Path(codp).exists():
        cod_idx, cod_keys = build_cod_neighborhood_index(codp)
    ml_nbhd_idx = build_masterlist_nbhd_index(mst)  # tiebreaker for COD ambiguities
    cod_n = sum(len(v) for v in cod_idx.values())
    lg.info(f"masterlist sites={len(mst)} indicators={len(ind_master)} "
            f"name-keys exact={len(name_idx_exact)} norm={len(name_idx_norm)} "
            f"COD neighborhoods={cod_n}")

    log = IssueLog(); frames=[]; partners=set(); proc=skip=0
    for path in sorted(Path(sub).glob("*Activities*.xlsx")):
        parsed = parse_filename(path.name)
        if not parsed:
            lg.warning(f"  [{path.name}] bad filename — skipped"); skip+=1; continue
        partner_fn, fmonths = parsed
        partners.add(partner_fn)
        fname_month_set = {mn for mn,_,_ in fmonths}
        month_label = "/".join(sorted(fname_month_set))
        df = read_partner_file(path, log, partner_fn, month_label, lg)
        if df is None: skip+=1; continue
        if df.empty:
            log.add(file=path.name,partner_from_filename=partner_fn,month_from_filename=month_label,
                    organization=partner_fn,error_type="EMPTY_SUBMISSION",severity="ERROR",
                    description="Template skeleton only — no data"); skip+=1; continue

        # ---- month resolution ----
        infile = df[C["month"]].apply(lambda v: GOV_CANON.get(str(v).strip().lower(), str(v).strip()) if pd.notna(v) else None)
        infile_canon = df[C["month"]].astype(str).str.strip().str.title()
        valid_infile = infile_canon[infile_canon.isin(MONTH_NAMES.values())]
        distinct_infile = set(valid_infile.unique())
        multi = (len(distinct_infile) >= 2) or (len(fname_month_set) >= 2)
        if multi:
            df["month_final"] = infile_canon.where(infile_canon.isin(MONTH_NAMES.values()),
                                                   other=(list(fname_month_set)[0] if len(fname_month_set)==1 else np.nan))
            method="in-file (multi-month container)"
        else:
            fm = list(fname_month_set)[0]
            df["month_final"] = fm
            bad = [v for v in distinct_infile if v != fm]
            if bad:
                log.add(file=path.name,partner_from_filename=partner_fn,month_from_filename=month_label,
                        organization=partner_fn,error_type="MONTH_MISMATCH",severity="WARNING",
                        description=f"in-file month {bad} != filename {fm}; filename used")
            method="filename"
        # keep only target months
        before=len(df)
        df = df[df["month_final"].isin(keep)].reset_index(drop=True)
        if df.empty:
            lg.info(f"  [{path.name}] no in-scope months ({method}) — skipped"); skip+=1; continue

        alias = build_indicator_alias(df[C["indicator"]], ind_master["Indicators"])
        cleaned = validate_and_clean(df, path.name, partner_fn, month_label, alias,
                                     ind_master, mst, name_idx_exact, name_idx_norm,
                                     norm_keys, ind_class, site_pop, log)
        cleaned["__source_file__"]=path.name
        cleaned["reporting_year"]=2026
        cleaned["reporting_period_id"]=cleaned["month_final"].map({"March":202603,"April":202604})
        cleaned["__is_valid__"]=(cleaned["indicator_clean"].notna() & cleaned["total_count"].notna()
            & cleaned["site_code_clean"].isin(set(mst["Site ID"]))
            & cleaned["governorate_clean"].isin(set(GOV_CANON.values())))
        frames.append(cleaned); proc+=1
        lg.info(f"  [{path.name}] {len(cleaned)} rows kept ({method}); valid={int(cleaned['__is_valid__'].sum())}")

    fact = pd.concat(frames,ignore_index=True,sort=False)
    lg.info(f"\nPre-dedup rows: {len(fact):,}")

    # ---- global content de-dup (collapse redundant re-submissions) ----
    dkey = [C["org"],"indicator_clean","site_code_clean","month_final","total_count_original"]+DEMO_COLS
    dkey = [k for k in dkey if k in fact.columns]
    dup = fact.duplicated(subset=dkey, keep="first")
    n_global = int(dup.sum())
    if n_global:
        # log a representative cross-file dup note per partner
        for org in fact.loc[dup, C["org"]].dropna().unique():
            cnt = int((dup & (fact[C["org"]]==org)).sum())
            log.add(file="(cross-file)",partner_from_filename=str(org),month_from_filename="Mar/Apr",
                    organization=str(org),error_type="DUPLICATE_CROSSFILE",severity="WARNING",
                    description=f"{cnt} duplicate row(s) collapsed across overlapping/redundant files")
    fact = fact.loc[~dup].reset_index(drop=True)
    lg.info(f"Removed {n_global:,} cross-file duplicate rows -> {len(fact):,} rows")
    lg.info(f"Valid rows: {int(fact['__is_valid__'].sum()):,}")

    tables, headline, site_reach = build_star(fact, mst, ind_master, partners, ind_class, lg,
                                               cod_idx=cod_idx, cod_keys=cod_keys, log=log,
                                               ml_nbhd_idx=ml_nbhd_idx)

    vlog = log.to_df()
    vsum = (vlog.groupby(["partner_from_filename","month_from_filename","error_type","severity"])
            .size().reset_index(name="count").sort_values(["partner_from_filename","error_type"])
            if not vlog.empty else pd.DataFrame())
    tables["validation_log"]=vlog; tables["validation_summary"]=vsum

    out.mkdir(parents=True,exist_ok=True)
    for name, d in tables.items():
        d.to_csv(out/f"{name}.csv",index=False)
        lg.info(f"  -> {name+'.csv':34s} {len(d):>6,} rows")
    site_reach.to_csv(out/"_site_reach_readingB.csv",index=False)

    lg.info(f"\nFiles processed={proc} skipped={skip} issues={len(vlog)}")
    lg.info(f"READING-B CLUSTER REACH (Mar+Apr, cumulative) = {headline:,}")
    lg.info(f"Sites with reported activity = {int(tables['dim_sites']['has_reported_activity'].sum())}")
    if not vlog.empty:
        lg.info("Top error types:")
        for et,n in vlog["error_type"].value_counts().head(12).items():
            lg.info(f"   {et:28s}{n:>5}")
    return tables, headline

if __name__=="__main__":
    p=argparse.ArgumentParser()
    p.add_argument("--submissions",type=Path,required=True)
    p.add_argument("--masterlist",type=Path,required=True)
    p.add_argument("--indicators",type=Path,required=True)
    p.add_argument("--cod",type=Path,default=None,
                   help="COD_Gaza.xlsx — Common Operational Dataset for adm4 neighborhood pcode fallback")
    p.add_argument("--months",type=str,default=None,
                   help="Comma-separated month names to keep (e.g. 'March,April'). Default: all valid months found.")
    p.add_argument("--out",type=Path,default=Path("./outputs"))
    a=p.parse_args()
    km = [m.strip().title() for m in a.months.split(",")] if a.months else None
    run(a.submissions,a.masterlist,a.indicators,a.out,codp=a.cod,keep_months=km)
