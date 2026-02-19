import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="生産実績集計", layout="wide")
st.title("作業者別 生産実績集計")

# =========================
# Excel出力用 共通関数
# =========================
def to_excel(df, sheet_name="Sheet1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# =========================
# CSVアップロード
# =========================
uploaded_file = st.file_uploader("着完システムのCSVファイルをアップロードしてください", type=["csv"])
if uploaded_file is None:
    st.stop()

df = pd.read_csv(uploaded_file)

# =========================
# 列名整理
# =========================
df.columns = df.columns.str.strip()

required_cols = [
    "時刻", "受注番号", "受注品番",
    "ステーション", "操作", "製造数", "受注数", "作業者"
]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"CSVに次の列が不足しています: {missing}")
    st.stop()

# =========================
# 前処理
# =========================
df["時刻"] = pd.to_datetime(df["時刻"], errors="coerce")
df["日付"] = df["時刻"].dt.date

# 集計用のみ使用（UI・生データには使わない）
def normalize_station_for_calc(name):
    if pd.isna(name):
        return name
    if "仕上げ" in name:
        return "仕上げ"
    return name

df["集計用ステーション"] = df["ステーション"].apply(normalize_station_for_calc)

# =========================
# フィルタUI（※統合しない）
# =========================
st.subheader("検索・フィルタ")

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    order_no = st.text_input("受注番号")

with c2:
    items = ["すべて"] + sorted(df["受注品番"].dropna().unique())
    selected_item = st.selectbox("受注品番", items)

with c3:
    stations = ["すべて"] + sorted(df["ステーション"].dropna().unique())
    # ★変更点: selectbox を multiselect に変更
    selected_stations = st.multiselect("ステーション", stations, default=["すべて"])

with c4:
    workers = ["すべて"] + sorted(df["作業者"].dropna().unique())
    selected_workers = st.multiselect("作業者", workers, default=["すべて"])

with c5:
    min_d, max_d = df["時刻"].min(), df["時刻"].max()
    date_range = st.date_input(
        "日付範囲",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d
    )

# =========================
# フィルタ処理（実データのみ）
# =========================
filtered_df = df.copy()

if order_no:
    filtered_df = filtered_df[
        filtered_df["受注番号"].astype(str).str.contains(order_no, na=False)
    ]

if selected_item != "すべて":
    filtered_df = filtered_df[filtered_df["受注品番"] == selected_item]

# ★変更点: 複数選択に対応するため isin を使用
if "
