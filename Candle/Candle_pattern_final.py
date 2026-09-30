import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Candles Pattern Backtest", layout="wide")

PATTERNS = [
    "BULLISH ENGULFING", "BULLISH HARAMI", "BULLISH MARUBOZU",
    "HAMMER", "INVERTED HAMMER", "MORNING STAR", "PIERCING",
    "RISING THREE METHODS", "THREE WHITE SOLDIERS", "TWEEZER BOTTOM"
]

def detect_patterns(df):
    n = len(df)
    r = {p: [""] * n for p in PATTERNS}
    for i in range(n):
        o,h,l,c = df.loc[i, ["Open","High","Low","Close"]]
        body = abs(c-o)

        if i >= 1:
            po,pc = df.loc[i-1, ["Open","Close"]]
            if po > pc and c > o and o <= pc and c >= po:
                r["BULLISH ENGULFING"][i] = "BULLISH ENGULFING"
            if po > pc and c > o and o > pc and c < po:
                r["BULLISH HARAMI"][i] = "BULLISH HARAMI"

        if body > 0 and c > o:
            if o-l <= .05*body and h-c <= .05*body:
                r["BULLISH MARUBOZU"][i] = "BULLISH MARUBOZU"

        if body > 0:
            lower = min(o,c)-l
            upper = h-max(o,c)
            if lower >= 2*body and upper <= .10*body:
                r["HAMMER"][i] = "HAMMER"
            if upper >= 2*body and lower <= .10*body:
                r["INVERTED HAMMER"][i] = "INVERTED HAMMER"

        if i >= 2:
            o1,c1 = df.loc[i-2,["Open","Close"]]
            o2,c2 = df.loc[i-1,["Open","Close"]]
            o3,c3 = df.loc[i,["Open","Close"]]
            if (c1 < o1 and abs(c2-o2) <= .5*abs(c1-o1)
                and c3 > o3 and c3 >= o1-.5*(o1-c1)):
                r["MORNING STAR"][i] = "MORNING STAR"
            if (c1 > o1 and c2 > o2 and c3 > o3 and c2 > c1 and c3 > c2
                and o2 >= o1 and o2 <= c1 and o3 >= o2 and o3 <= c2):
                r["THREE WHITE SOLDIERS"][i] = "THREE WHITE SOLDIERS"

        if i >= 1:
            o1,c1 = df.loc[i-1,["Open","Close"]]
            if c1 < o1 and c > o and c > o1-(o1-c1)/2 and c < o1:
                r["PIERCING"][i] = "PIERCING"

            o1,c1,l1 = df.loc[i-1,["Open","Close","Low"]]
            if c1 < o1 and c > o and abs(l-l1) <= .001*abs(l1):
                r["TWEEZER BOTTOM"][i] = "TWEEZER BOTTOM"

        if i >= 4:
            o1,c1 = df.loc[i-4,["Open","Close"]]
            small = all(abs(df.loc[j,"Close"]-df.loc[j,"Open"]) <= .5*abs(c1-o1)
                        for j in range(i-3,i))
            inside = all(df.loc[j,"High"] <= df.loc[i-4,"High"] and
                         df.loc[j,"Low"] >= df.loc[i-4,"Low"] for j in range(i-3,i))
            if c1 > o1 and small and inside and c > o and c > c1:
                r["RISING THREE METHODS"][i] = "RISING THREE METHODS"
    return r

def make_result(df, pattern, hold, flags):
    rows = []
    for i in range(len(df)):
        if flags[pattern][i] != pattern:
            continue
        entry = float(df.loc[i,"Close"])
        j = i + hold
        if j < len(df):
            exit_price = float(df.loc[j,"Close"])
            pnl = exit_price-entry
            pct = pnl/entry if entry else None
            exit_date = df.loc[j,"Date"]
            status = "EXIT"
        else:
            exit_price = pnl = pct = exit_date = None
            status = ""
        rows.append({
            "Entry Date":df.loc[i,"Date"], "Pattern":pattern,
            "Entry Price":entry, "Holding Candles":hold,
            "Exit Date":exit_date, "Exit":status,
            "Exit Price":exit_price, "PnL":pnl, "PnL %":pct
        })
    return pd.DataFrame(rows)

def excel_bytes(df, results, hold):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame({"Parameter":["Holding Candles"],"Value":[hold]}).to_excel(
            writer, sheet_name="PARAMETERS", index=False)
        summary=[]
        for p in PATTERNS:
            x=results[p]
            done=x.dropna(subset=["PnL"]) if not x.empty else x
            summary.append({
                "Pattern":p, "Entries":len(x),
                "Completed Exits":len(done),
                "Winning Trades":int((done["PnL"]>0).sum()) if len(done) else 0,
                "Losing Trades":int((done["PnL"]<0).sum()) if len(done) else 0,
                "Total PnL":float(done["PnL"].sum()) if len(done) else 0,
                "Average PnL %":float(done["PnL %"].mean()) if len(done) else 0
            })
            x.to_excel(writer, sheet_name=p[:31], index=False)
        pd.DataFrame(summary).to_excel(writer, sheet_name="SUMMARY", index=False)
    out.seek(0)
    return out

st.title("🕯️ Candles Pattern Backtest")
st.caption("OHLC → 10 Patterns → Entry → User Holding Candles → Exit → PnL")

hold = st.sidebar.number_input("Exit after how many candles?", 1, 10000, 10, 1)
uploaded = st.file_uploader("Upload Candle ohlc.xlsx", type=["xlsx","xls"])

if uploaded is None:
    st.info("Upload your OHLC Excel file to start.")
    st.stop()

if uploaded is None:
    st.info("Upload your OHLC Excel file to start.")
    st.stop()

try:
    raw = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"Could not read Excel file: {e}")
    st.stop()

lookup = {str(c).strip().lower(): c for c in raw.columns}

required = ["date", "open", "high", "low", "close"]

missing = [x for x in required if x not in lookup]

if missing:
    st.error("Missing columns: " + ", ".join(missing))
    st.stop()
df=raw[[lookup[x] for x in required]].copy()
df.columns=["Date","Open","High","Low","Close"]
for c in ["Open","High","Low","Close"]:
    df[c]=pd.to_numeric(df[c], errors="coerce")
df=df.dropna(subset=["Open","High","Low","Close"]).reset_index(drop=True)

flags=detect_patterns(df)
results={p:make_result(df,p,int(hold),flags) for p in PATTERNS}

summary=[]
for p in PATTERNS:
    x=results[p]; done=x.dropna(subset=["PnL"]) if not x.empty else x
    summary.append({
        "Pattern":p,"Entries":len(x),"Completed Exits":len(done),
        "Winning Trades":int((done["PnL"]>0).sum()) if len(done) else 0,
        "Losing Trades":int((done["PnL"]<0).sum()) if len(done) else 0,
        "Total PnL":float(done["PnL"].sum()) if len(done) else 0,
        "Average PnL %":float(done["PnL %"].mean()) if len(done) else 0
    })

sdf=pd.DataFrame(summary)
st.success(f"Loaded {len(df):,} candles. Holding period = {int(hold)} candles.")
a,b,c=st.columns(3)
a.metric("Pattern Entries", int(sdf["Entries"].sum()))
b.metric("Completed Exits", int(sdf["Completed Exits"].sum()))
c.metric("Total PnL", f'{sdf["Total PnL"].sum():,.4f}')

st.subheader("Summary")
st.dataframe(sdf, use_container_width=True)

tabs=st.tabs(PATTERNS)
for tab,p in zip(tabs,PATTERNS):
    with tab:
        if results[p].empty:
            st.warning("No occurrences of this pattern were found.")
        else:
            st.dataframe(results[p], use_container_width=True)

with st.expander("Original OHLC Data"):
    st.dataframe(df, use_container_width=True)

st.download_button(
    "⬇️ Download Excel Results",
    data=excel_bytes(df,results,int(hold)),
    file_name="Candle_ohlc.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
