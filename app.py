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
    selected_station = st.selectbox("ステーション", stations)

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

if selected_station != "すべて":
    filtered_df = filtered_df[filtered_df["ステーション"] == selected_station]

if "すべて" not in selected_workers:
    filtered_df = filtered_df[
        filtered_df["作業者"].isin(selected_workers)
    ]

if len(date_range) == 2:
    s, e = date_range
    filtered_df = filtered_df[
        (filtered_df["時刻"] >= pd.to_datetime(s)) &
        (filtered_df["時刻"] <= pd.to_datetime(e)+ pd.Timedelta(days=1))
    ]

# =========================
# 製造数集計用データ
# （仕上げのみ統合）
# =========================
finish_raw = filtered_df[
    (filtered_df["集計用ステーション"] == "仕上げ") &
    (filtered_df["操作"].isin(["完了", "中断"]))
]

finish_dedup = (
    finish_raw
    .groupby(
        ["受注番号", "受注品番", "集計用ステーション", "操作", "時刻"],
        as_index=False
    )
    .agg(
        製造数=("製造数", "max"),
        受注数=("受注数", "max")
    )
)

# =========================
# 受注単位 仕上げ完了率
# =========================
order_summary = (
    finish_dedup
    .groupby(["受注番号", "受注品番"], as_index=False)
    .agg(
        受注数=("受注数", "max"),
        仕上げ製造数=("製造数", "sum")
    )
)

order_summary["仕上げ完了率(%)"] = (
    order_summary["仕上げ製造数"] / order_summary["受注数"] * 100
).round(1)


# =====================================================
# ★ 作業者別 日別 工程別 製造実績 ＋ 作業時間（差分方式）
# =====================================================

import pandas as pd

worker_base = filtered_df.copy()

start_ops = ["開始", "再開"]
end_ops = ["中断", "完了"]

# 時刻を datetime に（未変換なら）
worker_base["時刻"] = pd.to_datetime(worker_base["時刻"])

# -----------------------------------------------------
# ① 同一時刻・同一条件の重複排除
# -----------------------------------------------------
worker_dedup = (
    worker_base
    .groupby(
        ["日付", "受注番号", "受注品番","ステーション", "作業者", "操作", "時刻"],
        as_index=False
    )
    .agg(製造数=("製造数", "max"))
)

# -----------------------------------------------------
# ② 開始系 / 終了系 に分離
# -----------------------------------------------------
start_df = worker_dedup[worker_dedup["操作"].isin(start_ops)]
end_df   = worker_dedup[worker_dedup["操作"].isin(end_ops)]

# -----------------------------------------------------
# ③ 区間単位の開始・終了情報
# -----------------------------------------------------
start_agg = (
    start_df
    .groupby(["日付", "受注番号","受注品番", "ステーション", "作業者"], as_index=False)
    .agg(
        開始時製造数=("製造数", "min"),
        開始時刻=("時刻", "min")
    )
)

end_agg = (
    end_df
    .groupby(["日付", "受注番号","受注品番", "ステーション", "作業者"], as_index=False)
    .agg(
        終了時製造数=("製造数", "max"),
        終了時刻=("時刻", "max")
    )
)

# -----------------------------------------------------
# ④ マージ → 実績算出
# -----------------------------------------------------
worker_diff = start_agg.merge(
    end_agg,
    on=["日付", "受注番号","受注品番", "ステーション", "作業者"],
    how="inner"
)

# 製造実績（差分）
worker_diff["実績製造数"] = (
    worker_diff["終了時製造数"] - worker_diff["開始時製造数"]
)

# 作業時間（分）
worker_diff["作業時間_分"] = (
    (worker_diff["終了時刻"] - worker_diff["開始時刻"])
    .dt.total_seconds() / 60
)

# 異常系除外
worker_diff = worker_diff[
    (worker_diff["実績製造数"] > 0) &
    (worker_diff["作業時間_分"] > 0)
]

# -----------------------------------------------------
# ⑤ 日別 × 作業者 × 工程
# -----------------------------------------------------
worker_daily_station = (
    worker_diff
    .groupby(["日付", "作業者", "ステーション"], as_index=False)
    .agg(
        日別製造数=("実績製造数", "sum"),
        作業時間_分=("作業時間_分", "sum")
    )
)

st.subheader("作業者別 日別・工程別 製造実績 ＋ 作業時間")
st.dataframe(
    worker_daily_station.sort_values(["日付", "作業者", "ステーション"]),
    use_container_width=True
)
st.download_button(
    "📥 作業者別 × 工程別 実績をExcel出力",
    data=to_excel(worker_daily_station, "作業者_工程別合計"),
    file_name="作業者別_工程別_製造実績.xlsx"
)

# -----------------------------------------------------
# ⑥ 全工程合算（日別 × 作業者）
# -----------------------------------------------------
worker_daily_total = (
    worker_daily_station
    .groupby(["日付", "作業者"], as_index=False)
    .agg(
        日別製造数=("日別製造数", "sum"),
        作業時間_分=("作業時間_分", "sum")
    )
)

st.subheader("作業者別 日別 製造実績 合計（全工程合算）")
st.dataframe(
    worker_daily_total.sort_values(["日付", "作業者"]),
    use_container_width=True
)
st.download_button(
    "📥 作業者別 × 日別 実績をExcel出力",
    data=to_excel(worker_daily_total, "作業者_日別合計"),
    file_name="作業者別_日別_製造実績.xlsx"
)
# =====================================================
# ★ 作業者別 × 受注品番別 製造数・作業時間
# =====================================================

worker_partno = (
    worker_diff
    .groupby(
        ["日付", "作業者", "受注品番"],
        as_index=False
    )
    .agg(
        製造数=("実績製造数", "sum"),
        作業時間_分=("作業時間_分", "sum")
    )
)

st.subheader("作業者別 × 受注品番別 製造実績・作業時間")
st.dataframe(
    worker_partno.sort_values(["日付", "作業者", "受注品番"]),
    use_container_width=True


)
st.download_button(
    "📥 作業者別 × 受注品番別 実績をExcel出力",
    data=to_excel(worker_partno, "作業者_品番別"),
    file_name="作業者別_受注品番別_製造実績.xlsx"
)
