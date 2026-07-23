import io, datetime as dt, requests, pandas as pd, streamlit as st
from concurrent.futures import ThreadPoolExecutor

st.set_page_config(page_title="NSE Derivatives", layout="wide")
REF = "https://www.nseindia.com/all-reports-derivatives"
ARCH = "https://nsearchives.nseindia.com/content/nsccl"
PARTS = ["FII", "Pro", "Client", "DII"]


@st.cache_resource
def session():
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
                      "Referer": REF})
    s.get(REF, timeout=15)
    return s


def get(url):
    try:
        r = session().get(url, timeout=30)
        return r.content if r.ok and len(r.content) > 200 else None
    except Exception:
        return None


@st.cache_data(ttl=300, show_spinner=False)
def load_mwpl():
    d = dt.date.today()
    for _ in range(10):
        c = get(f"{ARCH}/mwpl_cli_{d:%d%m%Y}.xls")
        if c and len(c) > 500:
            df = pd.read_excel(io.BytesIO(c), skiprows=1)
            cli = df.columns[2:]
            df[cli] = df[cli].apply(pd.to_numeric, errors="coerce")
            df["Count"] = df[cli].count(axis=1)
            df["Sum"] = df[cli].sum(axis=1).round(2)
            df["Average"] = df[cli].mean(axis=1).round(2)
            return d, df
        d -= dt.timedelta(days=1)
    return None, None


@st.cache_data(ttl=300, show_spinner=False)
def load_oi(days=30):
    today = dt.date.today()
    dates = [d for d in (today - dt.timedelta(days=i) for i in range(days + 1)) if d.weekday() < 5]

    def one(d):
        c = get(f"{ARCH}/fao_participant_oi_{d:%d%m%Y}.csv")
        if not c:
            return None
        df = pd.read_csv(io.BytesIO(c), skiprows=1)
        df.columns = [x.strip() for x in df.columns]
        df["Client Type"] = df["Client Type"].astype(str).str.strip()
        df = df.set_index("Client Type")
        row = {"_d": d}
        for p in PARTS:
            if p not in df.index:
                return None
            lo, sh = float(df.loc[p, "Future Index Long"]), float(df.loc[p, "Future Index Short"])
            t = lo + sh
            row[f"{p} Long %"] = round(lo / t * 100, 1) if t else 0.0
            row[f"{p} Short %"] = round(sh / t * 100, 1) if t else 0.0
            row[f"{p} Total"] = int(t)
        return row

    with ThreadPoolExecutor(max_workers=5) as ex:
        rows = [r for r in ex.map(one, dates) if r]
    if not rows:
        return None
    out = pd.DataFrame(rows).sort_values("_d", ascending=False)
    out["_d"] = out["_d"].map(lambda x: x.strftime("%d-%m-%Y"))
    return out.rename(columns={"_d": "Dates"}).reset_index(drop=True)


def xl(df, sheet):
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name=sheet)
    return buf.getvalue()


if st.button("Refresh"):
    load_mwpl.clear()
    load_oi.clear()

t1, t2 = st.tabs(["MWPL Client Positions", "Participant OI (1M)"])

with t1:
    date, df = load_mwpl()
    if df is None:
        st.error("No MWPL file found in the last 10 days.")
    else:
        st.caption(f"Position date: {date:%d %b %Y}")
        st.dataframe(df, use_container_width=True, hide_index=True,
                     height=min(1200, 35 * len(df) + 40))
        st.download_button("Excel", xl(df, "MWPL"),
                           f"mwpl_cli_{date:%d%m%Y}.xlsx", key="dl_mwpl")

with t2:
    oi = load_oi()
    if oi is None:
        st.error("No participant OI files found.")
    else:
        st.caption(f"{len(oi)} trading days · latest {oi['Dates'].iloc[0]}")
        st.dataframe(oi, use_container_width=True, hide_index=True,
                     height=min(1200, 35 * len(oi) + 40))
        st.download_button("Excel", xl(oi, "Participant OI"),
                           f"participant_oi_{dt.date.today():%d%m%Y}.xlsx", key="dl_oi")
