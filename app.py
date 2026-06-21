"""
Gaza SMC — 4W Compiler & QC tool
================================
A small web app for the Site Management Cluster IM team. Upload the monthly
4W partner submissions; the app runs the compilation pipeline
(compile_4w_v3.py) against the bundled reference data, then returns:

  • SMC_4W_<period>_PowerBI.xlsx   — the multi-sheet Power BI workbook
  • fact_activities.csv            — the cleaned, validated fact table
  • a full ZIP of all 13 star-schema CSVs
  • a readable error report: what must be fixed, what to flag to partners,
    and what the pipeline auto-corrected

Reference data (bundled in ./data, can be overridden in the sidebar):
    site_masterlist_*.csv      site population + site→agency mapping
    list_indicators_5w.csv     canonical Activity Index (25 indicators)
    COD_Gaza.xlsx              Common Operational Dataset (adm4 pcode fallback)

Run locally:   streamlit run app.py
Deploy:        push to GitHub → share.streamlit.io  (link from smcopt.org)
"""
from __future__ import annotations
import io
import re
import zipfile
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import compile_4w_v3 as pipeline
from build_powerbi_workbook import build_workbook, derive_period_label
from build_master_summary import build_master_summary

# ---------------------------------------------------------------- branding ---
SAPPHIRE = "#1B657C"; SIENNA = "#EC6B4D"; ECRU = "#F5F3E8"; BALTIC = "#2C2C2C"
APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"

st.set_page_config(page_title="SMC 4W Compiler", page_icon="🗂️", layout="wide")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter+Tight:wght@400;500;600&display=swap');
html, body, [class*="css"] {{ font-family:'Inter Tight',sans-serif; color:{BALTIC}; }}
.stApp {{ background:{ECRU}; }}
#MainMenu, footer {{ visibility:hidden; }}
.hero {{ background:{SAPPHIRE}; color:#eaf1f3; border-left:8px solid {SIENNA};
  border-radius:8px; padding:20px 26px; margin-bottom:8px; }}
.hero h1 {{ font-family:'Fraunces',serif; font-weight:600; font-size:25px; color:#fff; margin:0; }}
.hero p {{ color:#cfe1e6; margin:6px 0 0; font-size:14px; }}
.kpi {{ background:#fff; border-radius:7px; padding:14px 16px; border-top:3px solid {SAPPHIRE};
  box-shadow:0 1px 2px rgba(0,0,0,.04); }}
.kpi.accent {{ border-top-color:{SIENNA}; }}
.kpi .l {{ font-size:10.5px; text-transform:uppercase; letter-spacing:.07em; color:#6f7276; font-weight:600; }}
.kpi .v {{ font-family:'Fraunces',serif; font-size:26px; font-weight:600; color:{SAPPHIRE}; }}
.kpi.accent .v {{ color:{SIENNA}; }}
.stButton>button {{ background:{SIENNA}; color:#fff; border:0; border-radius:6px;
  font-weight:600; padding:8px 22px; }}
.stButton>button:hover {{ background:#d8543a; color:#fff; }}
section[data-testid="stSidebar"] {{ background:{SAPPHIRE}; }}
section[data-testid="stSidebar"] * {{ color:#eaf1f3 !important; }}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>Site Management Cluster — 4W Compiler & QC</h1>
  <p>Upload this cycle's partner 4W files. The tool validates and compiles them, then returns the Power BI workbook, the cleaned fact table, and a report of what needs fixing.</p>
</div>
""", unsafe_allow_html=True)


# --------------------------------------------------- optional password gate --
# If a secret named APP_PASSWORD is set (Streamlit → Settings → Secrets),
# the app asks for it before showing anything. If no secret is set, the app
# runs open — so deployment never locks you out by accident.
def _required_password():
    try:
        return st.secrets["APP_PASSWORD"]
    except Exception:
        return None

def password_gate():
    required = _required_password()
    if not required or st.session_state.get("auth_ok"):
        return
    pw = st.text_input("Access password", type="password",
                       help="Set by the Cluster IM team.")
    if pw == "":
        st.stop()
    if pw == required:
        st.session_state["auth_ok"] = True
        st.rerun()
    else:
        st.error("Incorrect password.")
        st.stop()

password_gate()


# -------------------------------------------------------------- reference ----
def find_reference():
    """Locate bundled reference files; return (masterlist, indicators, cod) Paths or None."""
    ml = next(iter(sorted(DATA_DIR.glob("*asterlist*.csv"))), None)
    ind = next(iter(sorted(DATA_DIR.glob("list_indicators_5w.csv"))), None)
    cod = next(iter(sorted(DATA_DIR.glob("COD*.xlsx"))), None)
    return ml, ind, cod


with st.sidebar:
    st.markdown("### Reference data")
    ml_def, ind_def, cod_def = find_reference()
    st.caption("Bundled with the app. Override only if you have newer versions.")
    ml_up = st.file_uploader("Site masterlist (.csv)", type=["csv"], key="ml")
    cod_up = st.file_uploader("COD_Gaza (.xlsx)", type=["xlsx"], key="cod")
    st.markdown("---")
    st.markdown("### Months to include")
    ALL_MONTHS = ["January","February","March","April","May","June",
                  "July","August","September","October","November","December"]
    month_mode = st.radio("Which reporting months?",
                          ["All months found in the files", "Choose specific months"],
                          index=0,
                          help="The pipeline reads each row's reporting month from the file. "
                               "‘All months found’ keeps everything you upload — safest default. "
                               "Switch to ‘specific’ to restrict to one cycle.")
    if month_mode == "Choose specific months":
        months = st.multiselect("Keep only these reporting months", ALL_MONTHS,
                                 default=["March","April"])
    else:
        months = "auto"  # keep only months present in the filenames (drops typos)
    st.markdown("---")
    st.caption("Reference status:")
    st.caption(f"• masterlist: {'✅ bundled' if ml_def else '⚠️ missing'}"
               f"{' (override loaded)' if ml_up else ''}")
    st.caption(f"• indicators: {'✅ bundled' if ind_def else '⚠️ missing'}")
    st.caption(f"• COD: {'✅ bundled' if cod_def else '⚠️ missing'}"
               f"{' (override loaded)' if cod_up else ''}")


# --------------------------------------------------------------- uploader ----
st.markdown("#### 1 · Upload partner 4W files")
uploads = st.file_uploader(
    "Drop the YYYYMM_Activities_PARTNER.xlsx files (you can select many at once)",
    type=["xlsx"], accept_multiple_files=True)

if uploads:
    st.caption(f"{len(uploads)} file(s) ready: " + ", ".join(u.name for u in uploads))

go = st.button("Compile & validate", disabled=not uploads)


# --------------------------------------------------------------- compile -----
def run_pipeline(upload_files, ml_path, ind_path, cod_path, keep_months):
    """Stage uploads + reference into a temp workspace, run the pipeline, build
    the Power BI workbook and the master_summary workbook."""
    work = Path(tempfile.mkdtemp(prefix="smc4w_"))
    sub = work / "submissions"; out = work / "outputs"
    sub.mkdir(); out.mkdir()
    for uf in upload_files:
        (sub / uf.name).write_bytes(uf.getvalue())
    if ml_path is None:
        raise FileNotFoundError("No site masterlist available (bundle one in ./data or upload it).")
    if ind_path is None:
        raise FileNotFoundError("No indicator list (list_indicators_5w.csv) available in ./data.")
    pipeline.run(sub, ml_path, ind_path, out, codp=cod_path, keep_months=keep_months)
    period = derive_period_label(out)
    safe = period.replace(" ", "").replace("\u2013", "").replace("-", "") or "output"
    xlsx = out / f"SMC_4W_{safe}_PowerBI.xlsx"
    info = build_workbook(out, xlsx, period)
    # master_summary (colleague's format, this cluster's codes + mapping tab).
    # Optional: never let it block the core deliverables.
    summary_xlsx = None; summary_error = None
    coll_codes = next(iter(sorted(DATA_DIR.glob("colleague_indicator_codes.csv"))), None)
    if coll_codes is not None:
        try:
            cand = out / f"master_summary_{safe}.xlsx"
            build_master_summary(out, ml_path, ind_path, coll_codes, cand, period)
            summary_xlsx = cand
        except Exception as e:
            summary_error = str(e)
    return out, xlsx, period, info, summary_xlsx, summary_error


def staged_reference(ml_up, cod_up):
    """Write any sidebar-uploaded reference files to temp, else use bundled."""
    ml_def, ind_def, cod_def = find_reference()
    tmp = Path(tempfile.mkdtemp(prefix="smc4w_ref_"))
    ml = ml_def; cod = cod_def
    if ml_up is not None:
        ml = tmp / ml_up.name; ml.write_bytes(ml_up.getvalue())
    if cod_up is not None:
        cod = tmp / cod_up.name; cod.write_bytes(cod_up.getvalue())
    return ml, ind_def, cod


def error_report(out_dir: Path) -> pd.DataFrame:
    vlog = pd.read_csv(out_dir / "validation_log.csv")
    return vlog


SEVERITY_HELP = {
    "ERROR":   "Excludes the row from the validated dataset — must be fixed and resubmitted.",
    "WARNING": "Row is kept but flagged — worth correcting and reminding the partner.",
    "INFO":    "The pipeline auto-corrected this (e.g. capped reach, clamped site count, COD pcode).",
}


if go and uploads:
    try:
        ml_path, ind_path, cod_path = staged_reference(ml_up, cod_up)
        with st.spinner("Reading files, validating rows, compiling the star schema…"):
            out, xlsx, period, info, summary_xlsx, summary_error = run_pipeline(
                uploads, ml_path, ind_path, cod_path, months)
        st.session_state["result"] = {
            "out": str(out), "xlsx": str(xlsx), "period": period, "info": info,
            "summary_xlsx": str(summary_xlsx) if summary_xlsx else None,
            "summary_error": summary_error,
        }
    except Exception as e:
        st.error(f"Compilation failed: {e}")
        st.stop()


# --------------------------------------------------------------- results -----
res = st.session_state.get("result")
if res:
    out = Path(res["out"]); xlsx = Path(res["xlsx"]); info = res["info"]; period = res["period"]
    summary_xlsx = Path(res["summary_xlsx"]) if res.get("summary_xlsx") else None

    st.markdown(f"#### 2 · Results — {period or 'compiled'}")
    # which months actually landed (so the user can confirm scope)
    try:
        dd = pd.read_csv(out / "dim_dates.csv")
        st.caption("Months compiled: " + ", ".join(
            f"{m} {int(y)}" for m, y in zip(dd["reporting_month"], dd["reporting_year"])))
    except Exception:
        pass
    c1, c2, c3, c4 = st.columns(4)
    for col, label, val, accent in [
        (c1, "Cluster reach", f"{info['reach']:,}" if info['reach'] is not None else "—", True),
        (c2, "Sites reached", f"{info['sites_reported']:,} / {info['total_sites']:,}", False),
        (c3, "Validated rows", f"{info['valid_rows']:,}", False),
        (c4, "Partners", f"{info['partners']}", False),
    ]:
        col.markdown(f"<div class='kpi {'accent' if accent else ''}'><div class='l'>{label}</div>"
                     f"<div class='v'>{val}</div></div>", unsafe_allow_html=True)

    st.markdown("#### 3 · Download")
    d1, d2, d3, d4 = st.columns(4)
    with d1:
        st.download_button("⬇ Power BI workbook (.xlsx)", data=xlsx.read_bytes(),
                           file_name=xlsx.name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        fa = out / "fact_activities.csv"
        st.download_button("⬇ fact_activities.csv", data=fa.read_bytes(),
                           file_name="fact_activities.csv", mime="text/csv")
    with d3:
        if summary_xlsx and summary_xlsx.exists():
            st.download_button("⬇ master_summary (.xlsx)", data=summary_xlsx.read_bytes(),
                               file_name=summary_xlsx.name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        elif res.get("summary_error"):
            st.caption("master_summary skipped (core outputs unaffected)")
        else:
            st.caption("master_summary unavailable (colleague codes not bundled)")
    with d4:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for csv in sorted(out.glob("*.csv")):
                if not csv.name.startswith("_"):
                    z.write(csv, csv.name)
        st.download_button("⬇ All 13 tables (.zip)", data=buf.getvalue(),
                           file_name=f"SMC_4W_{period.replace(' ','')}_star_schema.zip",
                           mime="application/zip")

    # ----- error report -----
    st.markdown("#### 4 · What needs fixing")
    vlog = error_report(out)
    if vlog.empty:
        st.success("No validation issues logged — clean submission set.")
    else:
        counts = (vlog.groupby("severity").size()
                  .reindex(["ERROR", "WARNING", "INFO"]).fillna(0).astype(int))
        s1, s2, s3 = st.columns(3)
        s1.metric("Must fix (errors)", int(counts.get("ERROR", 0)))
        s2.metric("Flag to partners (warnings)", int(counts.get("WARNING", 0)))
        s3.metric("Auto-corrected (info)", int(counts.get("INFO", 0)))

        tabs = st.tabs(["🔴 Must fix", "🟠 Flag to partners", "🔵 Auto-corrected", "By partner"])
        for tab, sev in zip(tabs[:3], ["ERROR", "WARNING", "INFO"]):
            with tab:
                st.caption(SEVERITY_HELP[sev])
                sub = vlog[vlog["severity"] == sev]
                if sub.empty:
                    st.write("None.")
                    continue
                by_type = (sub.groupby("error_type").size()
                           .sort_values(ascending=False).reset_index(name="count"))
                st.dataframe(by_type, hide_index=True, use_container_width=True)
                show = sub[["file", "organization", "site", "indicator",
                            "error_type", "description"]].copy()
                st.dataframe(show, hide_index=True, use_container_width=True, height=320)
        with tabs[3]:
            piv = (vlog.assign(partner=vlog["organization"].replace("", "(unattributed)"))
                   .pivot_table(index="partner", columns="severity", values="error_type",
                                aggfunc="count", fill_value=0))
            for c in ("ERROR", "WARNING", "INFO"):
                if c not in piv.columns:
                    piv[c] = 0
            piv = piv[["ERROR", "WARNING", "INFO"]].rename(
                columns={"ERROR": "Must fix", "WARNING": "Warnings", "INFO": "Auto-fixed"})
            st.dataframe(piv.sort_values("Must fix", ascending=False),
                         use_container_width=True)

        # downloadable error workbook: a Summary tab + one tab per agency
        def _safe_sheet(name):
            s = re.sub(r"[\\/?*\[\]:]", " ", str(name)).strip() or "unattributed"
            return s[:31]

        ebuf = io.BytesIO()
        vlog_a = vlog.copy()
        vlog_a["agency"] = vlog_a["organization"].replace("", pd.NA)
        vlog_a["agency"] = vlog_a["agency"].fillna(vlog_a["partner_from_filename"]).replace("", "(unattributed)")
        with pd.ExcelWriter(ebuf, engine="openpyxl") as xw:
            # Summary: counts per agency x severity
            summ = (vlog_a.pivot_table(index="agency", columns="severity",
                                       values="error_type", aggfunc="count", fill_value=0))
            for c in ("ERROR", "WARNING", "INFO"):
                if c not in summ.columns:
                    summ[c] = 0
            summ = (summ[["ERROR", "WARNING", "INFO"]]
                    .rename(columns={"ERROR": "Must fix", "WARNING": "Warnings", "INFO": "Auto-fixed"})
                    .sort_values("Must fix", ascending=False).reset_index())
            summ.to_excel(xw, sheet_name="Summary", index=False)
            # one tab per agency
            cols = ["file", "row", "organization", "site", "indicator",
                    "error_type", "severity", "description"]
            cols = [c for c in cols if c in vlog_a.columns]
            used = set()
            for agency in summ["agency"]:
                sub = vlog_a[vlog_a["agency"] == agency][cols]
                sheet = _safe_sheet(agency)
                base = sheet; i = 1
                while sheet.lower() in used:
                    sheet = f"{base[:28]}_{i}"; i += 1
                used.add(sheet.lower())
                sub.to_excel(xw, sheet_name=sheet, index=False)
        st.download_button("⬇ Error report — one tab per agency (.xlsx)", data=ebuf.getvalue(),
                           file_name=f"SMC_4W_{period.replace(' ','')}_errors.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload one or more 4W files above, then press **Compile & validate**.")
