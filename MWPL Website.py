import io, os, datetime as dt, requests, pandas as pd, streamlit as st, altair as alt
from concurrent.futures import ThreadPoolExecutor

st.set_page_config(page_title="NSE Derivatives", page_icon="📊", layout="wide")
REF = "https://www.nseindia.com/all-reports-derivatives"
ARCH = "https://nsearchives.nseindia.com/content/nsccl"
PARTS = ["FII", "Pro", "Client", "DII"]
LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Ashika_logo-removebg-preview.png")


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
    return df.to_csv(index=False).encode()


st.markdown("""
<style>
  div[data-testid="stButton"] button {padding:2px 12px; font-size:12px; min-height:0;}
  div[data-testid="stButton"] {display:flex; justify-content:flex-end;}
  h1 {padding-top:0 !important; margin-top:0 !important;}
</style>
""", unsafe_allow_html=True)

a1, a2 = st.columns([5, 1], vertical_alignment="center")
with a1:
    st.title("NSE Derivatives Monitor")
with a2:
    if os.path.exists(LOGO):
        st.image(LOGO, width=130)

b1, b2 = st.columns([5, 1], vertical_alignment="center")
with b1:
    st.caption("MWPL client positions · participant-wise open interest · live from NSE")
with b2:
    if st.button("Refresh"):
        load_mwpl.clear()
        load_oi.clear()

t0, t1, t2 = st.tabs(["About", "MWPL Client Positions", "Participant OI (1M)"])

with t0:
    st.markdown("""
**MWPL Client Positions**  
NSE's "F&O Clients Position % greater than or equal to 3% of Stock MWPL" file.
Each row is an underlying stock; each Client column is one client holding at least
3% of that stock's market-wide position limit.

- **Count**: how many clients are above the 3% threshold
- **Sum**: their combined position as % of MWPL
- **Average**: mean position size per reporting client

High Count with low Average is a crowded trade. Low Count with high Average is
concentration in one or two hands.

**Participant OI**  
NSE's "Participant wise Open Interest" file for the last ~21 trading days.
Long % and Short % are each participant's share of index futures open interest,
computed as long / (long + short). Dashed lines are the average across the window.

**Notes**  
Position date runs one trading day behind the publish date. Missing days are holidays.
Data refreshes every 5 minutes, or immediately via Refresh. Nothing is stored locally.
Source: NSE India. Internal research use only, not investment advice.
""")

with t1:
    date, df = load_mwpl()
    if df is None:
        st.error("No MWPL file found in the last 10 days.")
    else:
        st.caption(f"Position date: {date:%d %b %Y}")
        sym = df.columns[1]
        top = df[[sym, "Sum", "Count", "Average"]].dropna().nlargest(20, "Sum")
        st.subheader("Top 20 by total client position")
        st.altair_chart(
            alt.Chart(top).mark_bar().encode(
                y=alt.Y(f"{sym}:N", sort="-x", title=None),
                x=alt.X("Sum:Q", title="Total % of MWPL"),
                color=alt.Color("Count:Q", scale=alt.Scale(scheme="tealblues"), title="Clients"),
                tooltip=[sym, "Sum", "Count", "Average"],
            ).properties(height=480), use_container_width=True)

        st.dataframe(df, use_container_width=True, hide_index=True,
                     height=min(1200, 35 * len(df) + 40))
        st.download_button("CSV", xl(df, "MWPL"),
                           f"mwpl_cli_{date:%d%m%Y}.csv", "text/csv", key="dl_mwpl")

with t2:
    oi = load_oi()
    if oi is None:
        st.error("No participant OI files found.")
    else:
        st.caption(f"{len(oi)} trading days · latest {oi['Dates'].iloc[0]}")
        src = oi.iloc[::-1].reset_index(drop=True)
        order = list(src["Dates"])
        xax = alt.X("Dates:O", sort=order, title=None,
                    axis=alt.Axis(labelAngle=-45, labelOverlap=True))
        shade = alt.Scale(domain=["Long", "Short"], range=["#26a65b", "#e05260"])
        cols = st.columns(2)
        for i, p in enumerate(PARTS):
            m = src[["Dates", f"{p} Long %", f"{p} Short %"]].melt(
                "Dates", var_name="Side", value_name="Value")
            m["Side"] = m["Side"].str.extract(r"(Long|Short)")
            avg = m.groupby("Side")["Value"].mean().round(2)
            a = pd.DataFrame({"Side": avg.index, "Avg": avg.values})
            with cols[i % 2]:
                st.subheader(p)
                st.caption(f"Avg long {avg['Long']}% · avg short {avg['Short']}%")
                st.altair_chart(
                    (alt.Chart(m).mark_line(strokeWidth=2).encode(
                        x=xax,
                        y=alt.Y("Value:Q", scale=alt.Scale(zero=False), title="% of futures OI"),
                        color=alt.Color("Side:N", scale=shade)) +
                     alt.Chart(m).mark_point(size=55, filled=True).encode(
                         x=xax, y="Value:Q", color=alt.Color("Side:N", scale=shade),
                         tooltip=["Dates", "Side", "Value"]) +
                     alt.Chart(a).mark_rule(strokeDash=[5, 3], opacity=.6).encode(
                         y="Avg:Q", color=alt.Color("Side:N", scale=shade),
                         tooltip=["Side", "Avg"])
                     ).properties(height=320).interactive(), use_container_width=True)

        st.dataframe(oi, use_container_width=True, hide_index=True,
                     height=min(1200, 35 * len(oi) + 40))
        st.download_button("CSV", xl(oi, "Participant OI"),
                           f"participant_oi_{dt.date.today():%d%m%Y}.csv", "text/csv", key="dl_oi")
