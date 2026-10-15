import os, sys, subprocess, time
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

BASE=Path(__file__).resolve().parent
CRAWLER=BASE/"seo_crawler.py"
REPORT=BASE/"seo_audit_report.xlsx"
LOG=BASE/"crawl_log.txt"
PID=BASE/"crawl_process.pid"

st.set_page_config(page_title="Site Audit",page_icon="🔎",layout="wide")
st.markdown("""
<style>
.block-container{padding-top:1.4rem;max-width:1600px}
.hero{background:linear-gradient(120deg,#14213d,#007f8b);padding:24px 28px;border-radius:16px;color:white;margin-bottom:18px}
.hero h1{color:white;margin:0;font-size:2rem}.hero p{color:#e5f6f6;margin:5px 0 0}
div[data-testid="stMetric"]{border:1px solid #dfe5ec;padding:14px;border-radius:12px;background:#fff}
section[data-testid="stSidebar"]{background:#f6f8fb}
</style>
<div class="hero"><h1>Site Audit</h1><p>Technical SEO crawler · Crawl overview · URL inspection · Excel export</p></div>
""",unsafe_allow_html=True)

def running():
    if not PID.exists(): return False
    try:
        pid=int(PID.read_text().strip())
        if os.name=="nt":
            out=subprocess.run(["tasklist","/FI",f"PID eq {pid}","/FO","CSV","/NH"],capture_output=True,text=True,timeout=4).stdout
            return str(pid) in out
        os.kill(pid,0); return True
    except Exception: return False

def get_data():
    if REPORT.exists():
        try: return pd.read_excel(REPORT,sheet_name="1 - All")
        except Exception as e: st.error(f"Report could not be opened: {e}")
    return pd.DataFrame()

with st.sidebar:
    st.markdown("## New site audit")
    url=st.text_input("Website URL",placeholder="https://www.example.com/")
    max_pages=st.select_slider("Maximum URLs",options=[20,50,100,250,500,1000,5000,10000,50000,100000],value=100)
    delay=st.select_slider("Delay between requests",options=[0.5,1.0,1.5,2.0,3.0,5.0],value=1.0)
    sitemap=st.checkbox("Discover URLs from XML sitemaps",value=True)
    st.caption("Crawl only sites you own or are authorized to audit. This version enforces a minimum 0.5-second delay.")
    if st.button("▶ Start audit",type="primary",use_container_width=True):
        if not url.strip(): st.error("Enter a website URL.")
        elif running(): st.warning("A crawl appears to be running. Refresh the page to view progress.")
        else:
            cmd=[sys.executable,str(CRAWLER),url.strip(),"--max-pages",str(max_pages),"--delay",str(delay),"--output",str(REPORT)]
            if not sitemap: cmd.append("--no-sitemap")
            with open(LOG,"w",encoding="utf-8") as f:
                proc=subprocess.Popen(cmd,cwd=str(BASE),stdout=f,stderr=subprocess.STDOUT,
                    creationflags=(subprocess.CREATE_NEW_PROCESS_GROUP if os.name=="nt" else 0))
            PID.write_text(str(proc.pid))
            st.success("Audit started. Use Refresh audit to update.")
    if st.button("↻ Refresh audit",use_container_width=True):
        st.rerun()

is_running=running()
data=get_data()
if is_running: st.info("Crawl is running. Refresh the page to update the latest report.")
else: st.caption("No crawl process currently detected.")

if not data.empty:
    statuses=pd.to_numeric(data.get("Status Code",pd.Series(dtype=float)),errors="coerce")
    errors=int(statuses.ge(400).sum())
    nonindex=int(data.get("Indexability",pd.Series(dtype=str)).astype(str).str.contains("Non-Indexable|Blocked",case=False,regex=True).sum())
    no_title=int(data.get("Title 1",pd.Series(dtype=str)).fillna("").astype(str).str.strip().eq("").sum())
    duplicate_titles=int(data.get("Title 1",pd.Series(dtype=str)).fillna("").astype(str).str.casefold().duplicated(keep=False).sum())
else: errors=nonindex=no_title=duplicate_titles=0

m=st.columns(5)
m[0].metric("URLs crawled",f"{len(data):,}")
m[1].metric("HTTP errors",f"{errors:,}")
m[2].metric("Non-indexable / blocked",f"{nonindex:,}")
m[3].metric("Missing titles",f"{no_title:,}")
m[4].metric("Duplicate titles",f"{duplicate_titles:,}")

left,right=st.columns([1.6,1])
with left:
    st.subheader("Crawl activity")
    if LOG.exists():
        log=LOG.read_text(encoding="utf-8",errors="replace")
        st.code("\n".join(log.splitlines()[-18:]) or "Waiting for crawler output…",language="text")
    else: st.info("Start an audit to see crawl output here.")
with right:
    st.subheader("Export report")
    if REPORT.exists():
        st.write("Your workbook follows the 75-column structure of your example, on the `1 - All` worksheet.")
        with open(REPORT,"rb") as f:
            st.download_button("⬇ Download Excel report",f.read(),file_name=f"seo_audit_{datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    else: st.info("Your Excel export will appear after the crawl starts writing results.")

st.divider()
st.subheader("URL inspection")
if data.empty:
    st.info("No crawl data yet. Start with 20–100 URLs for your first test.")
else:
    q=st.text_input("Search URLs, titles or indexability",placeholder="Type a URL, title, status…")
    view=data
    if q.strip():
        mask=pd.Series(False,index=data.index)
        for col in ["Address","Title 1","Meta Description 1","Indexability","Status","Content Type"]:
            if col in data.columns: mask|=data[col].fillna("").astype(str).str.contains(q,case=False,regex=False)
        view=data[mask]
    cols=["Address","Status Code","Status","Indexability","Title 1","Title 1 Length","Meta Description 1","Word Count","Crawl Depth","Inlinks","Outlinks","Response Time"]
    cols=[c for c in cols if c in view.columns]
    st.caption(f"Showing {len(view):,} of {len(data):,} rows · Excel contains the full 75-column schema.")
    st.dataframe(view[cols],use_container_width=True,hide_index=True)
    with st.expander("View all 75 columns"):
        st.dataframe(view,use_container_width=True,hide_index=True)
st.caption("Independent starter tool inspired by common technical SEO audit workflows; not affiliated with Semrush.")
