import io, datetime as dt, requests, pandas as pd, streamlit as st

st.set_page_config(page_title="MWPL Client Positions", layout="wide")
REF = "https://www.nseindia.com/all-reports-derivatives"


@st.cache_data(ttl=300, show_spinner=False)
def load():
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
                      "Referer": REF})
    s.get(REF, timeout=15)
    d = dt.date.today()
    for _ in range(10):
        r = s.get(f"https://nsearchives.nseindia.com/content/nsccl/mwpl_cli_{d:%d%m%Y}.xls", timeout=30)
        if r.ok and len(r.content) > 500:
            df = pd.read_excel(io.BytesIO(r.content), skiprows=1)
            cli = df.columns[2:]
            df[cli] = df[cli].apply(pd.to_numeric, errors="coerce")
            df["Count"] = df[cli].count(axis=1)
            df["Sum"] = df[cli].sum(axis=1).round(2)
            df["Average"] = df[cli].mean(axis=1).round(2)
            return d, df
        d -= dt.timedelta(days=1)
    return None, None


if st.button("Refresh"):
    load.clear()

date, df = load()
if df is None:
    st.error("No MWPL file found in the last 10 days.")
else:
    st.caption(f"Position date: {date:%d %b %Y}")
    st.dataframe(df, use_container_width=True, hide_index=True,
                 height=min(1200, 35 * len(df) + 40))
    st.download_button("CSV", df.to_csv(index=False).encode(),
                       f"mwpl_cli_{date:%d%m%Y}.csv", "text/csv")