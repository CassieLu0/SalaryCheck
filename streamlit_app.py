import re
from datetime import datetime, date
from io import BytesIO

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="DSP Salary Summary Tool", layout="wide")
st.title("DSP Salary Summary Tool")

# =========================
# Helpers
# =========================
def safe_filename(s: str) -> str:
    s = str(s).strip()
    # remove characters invalid for filenames on common OS
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", "", s)
    return s or "工资单"

def parse_fleet_and_period_from_filename(name: str):
    """
    Try parse:
      <fleet>-派费明细-02_02_2026-08_02_2026-...xlsx
    Return (fleet, start_date, end_date) with date objects or None.
    """
    fleet = None
    d1 = None
    d2 = None

    if not name:
        return fleet, d1, d2

    # fleet: anything before "-派费明细-"
    m_fleet = re.search(r"^(.*?)-派费明细-", name)
    if m_fleet:
        fleet = m_fleet.group(1).strip()

    # period: try find "MM_DD_YYYY-MM_DD_YYYY"
    m_period = re.search(r"(\d{2}_\d{2}_\d{4})-(\d{2}_\d{2}_\d{4})", name)
    if m_period:
        s1, s2 = m_period.group(1), m_period.group(2)
        try:
            d1 = datetime.strptime(s1, "%m_%d_%Y").date()
            d2 = datetime.strptime(s2, "%m_%d_%Y").date()
        except Exception:
            d1, d2 = None, None

    return fleet, d1, d2

def fmt_period(d1: date | None, d2: date | None) -> str:
    if not d1 or not d2:
        return "未知帐期"
    return f"{d1.strftime('%m_%d_%Y')}-{d2.strftime('%m_%d_%Y')}"

def build_workbook(
    df_delivery: pd.DataFrame,
    df_route_price: pd.DataFrame,
    df_offset: pd.DataFrame,
    df_claim_raw: pd.DataFrame | None,
    sheet_names: list[str],
):
    # =========================
    # Sheet names
    # =========================
    SHEET_ROUTE = "参数_线路明细"
    SHEET_OFFSET = "冲抵明细"
    SHEET_CLAIM = "理赔明细"

    # =========================
    # Columns mapping
    # =========================
    COL_DRIVER_ID = "快递员ID"
    COL_DRIVER = "快递员"
    COL_ROUTE = "区域/线路"
    COL_WEIGHT = "结算重量lb"
    COL_TASK = "任务号"
    COL_STOP = "STOP序号"

    # =========================
    # Weight buckets (3 tiers)
    # =========================
    BUCKETS = [
        ("0-5lb", 0, 5),         # [0,5)
        ("5-20lb", 5, 20),       # [5,20)
        ("20lb以上", 20, None),   # [20, +inf)
    ]

    def weight_bucket(w):
        try:
            w = float(w)
        except Exception:
            w = 0.0

        for name, lb_min, lb_max in BUCKETS:
            if lb_max is None:
                if w >= lb_min:
                    return name
            else:
                if lb_min <= w < lb_max:
                    return name
        return BUCKETS[0][0]

    # =========================
    # Build offset summary: by driver_id
    # =========================
    offset_cnt = {}
    offset_amt = {}
    if df_offset is not None and not df_offset.empty:
        if "快递员ID" in df_offset.columns and "费用合计_未税" in df_offset.columns:
            dfo = df_offset.copy()
            dfo["快递员ID"] = dfo["快递员ID"].astype(str).str.strip()
            dfo["费用合计_未税"] = pd.to_numeric(dfo["费用合计_未税"], errors="coerce").fillna(0.0)

            g_off = dfo.groupby("快递员ID", dropna=False).agg(
                cnt=("费用合计_未税", "size"),
                amt=("费用合计_未税", "sum"),
            )
            offset_cnt = g_off["cnt"].to_dict()
            offset_amt = g_off["amt"].to_dict()

    # =========================
    # Build claim summary: claim has no ID -> map name -> most common ID from delivery
    # =========================
    claim_cnt = {}  # (did, dtype) -> cnt
    claim_amt = {}  # (did, dtype) -> amt

    if df_claim_raw is not None and not df_claim_raw.empty:
        df_claim = df_claim_raw.copy()
        # need columns
        need_cl = ["快递员", "理赔类型", "费用合计_未税"]
        if all(c in df_claim.columns for c in need_cl):
            # name -> most common id
            name_to_id = (
                df_delivery[[COL_DRIVER, COL_DRIVER_ID]]
                .dropna()
                .assign(**{
                    COL_DRIVER: lambda x: x[COL_DRIVER].astype(str).str.strip(),
                    COL_DRIVER_ID: lambda x: x[COL_DRIVER_ID].astype(str).str.strip(),
                })
                .groupby(COL_DRIVER)[COL_DRIVER_ID]
                .agg(lambda s: s.value_counts().index[0])
                .to_dict()
            )

            def map_claim_type(x: str):
                x = str(x).strip()
                if x == "轨迹断更":
                    return "断更"
                if x == "虚假签收":
                    return "虚假签收"
                return None

            df_claim["_name"] = df_claim["快递员"].astype(str).str.strip()
            df_claim["_did"] = df_claim["_name"].map(name_to_id)
            df_claim["_dtype"] = df_claim["理赔类型"].apply(map_claim_type)
            df_claim = df_claim[df_claim["_dtype"].notna()].copy()
            df_claim = df_claim[df_claim["_did"].notna()].copy()

            df_claim["费用合计_未税"] = pd.to_numeric(df_claim["费用合计_未税"], errors="coerce").fillna(0.0)

            g_cl = df_claim.groupby(["_did", "_dtype"], dropna=False).agg(
                cnt=("费用合计_未税", "size"),
                amt=("费用合计_未税", "sum"),
            )

            for (did, dtype), row in g_cl.iterrows():
                did_s = str(did).strip()
                dtype_s = str(dtype).strip()
                claim_cnt[(did_s, dtype_s)] = int(row["cnt"])
                claim_amt[(did_s, dtype_s)] = float(row["amt"])

    # =========================
    # Prepare delivery dataframe for counting
    # =========================
    df = df_delivery.copy()
    df[COL_DRIVER_ID] = df[COL_DRIVER_ID].astype(str).str.strip()
    df[COL_DRIVER] = df[COL_DRIVER].astype(str).str.strip()
    df["_route"] = df[COL_ROUTE].astype(str).fillna("").str.strip()
    df["_weight"] = pd.to_numeric(df[COL_WEIGHT], errors="coerce").fillna(0.0)
    df["_bucket"] = df["_weight"].apply(weight_bucket)

    # 首票/联单：同 driver + 任务号 + STOP序号，第一条首票，其余联单
    gkey = [COL_DRIVER_ID, COL_TASK, COL_STOP]
    df["_rank_in_stop"] = df.groupby(gkey).cumcount() + 1
    df["_ticket"] = df["_rank_in_stop"].eq(1).map({True: "首票", False: "联单"})

    # 件数统计：driver + route + bucket + ticket
    cnt = (
        df.groupby([COL_DRIVER_ID, COL_DRIVER, "_route", "_bucket", "_ticket"], dropna=False)
          .size()
          .reset_index(name="件数")
    )

    def get_cnt(did, route, bucket, ticket):
        m = cnt[
            (cnt[COL_DRIVER_ID] == did) &
            (cnt["_route"] == route) &
            (cnt["_bucket"] == bucket) &
            (cnt["_ticket"] == ticket)
        ]
        return 0 if m.empty else int(m.iloc[0]["件数"])

    # 一行一个 driver + route
    drivers = (
        df[[COL_DRIVER_ID, COL_DRIVER, "_route"]]
          .drop_duplicates()
          .sort_values(["_route", COL_DRIVER])
    )

    # headers：每个档位各 2 组（首票/联单），每组 2 列（件数/金额）
    headers = []
    for b in [x[0] for x in BUCKETS]:
        headers.append((b, "首票"))
        headers.append((b, "联单"))

    # =========================
    # Build workbook
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "司机汇总"

    # 参数_线路明细（单价来源）
    df_route_price = df_route_price.copy()
    df_route_price["匹配键"] = df_route_price["线路"].astype(str).str.strip() + "|" + df_route_price["档位"].astype(str).str.strip()

    ws_route = wb.create_sheet(SHEET_ROUTE)
    ws_route.append(["线路", "档位", "首票单价", "联单单价", "匹配键"])
    for cell in ws_route[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for _, row in df_route_price.iterrows():
        ws_route.append([
            row["线路"],
            row["档位"],
            float(row["首票单价"]),
            float(row["联单单价"]),
            row["匹配键"],
        ])

    # （可选）把明细写进 output 方便对账
    def write_df_to_sheet(book: Workbook, name: str, dfx: pd.DataFrame):
        wsx = book.create_sheet(name)
        wsx.append(list(dfx.columns))
        for cell in wsx[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in dfx.itertuples(index=False):
            wsx.append(list(row))

    if df_offset is not None and not df_offset.empty:
        write_df_to_sheet(wb, SHEET_OFFSET, df_offset)

    if df_claim_raw is not None and not df_claim_raw.empty:
        write_df_to_sheet(wb, SHEET_CLAIM, df_claim_raw)

    # =========================
    # Styles
    # =========================
    thin = Side(style="thin", color="666666")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bold = Font(bold=True)
    title_font = Font(bold=True, size=16)

    fill_head = PatternFill("solid", fgColor="F2D6C9")
    fill_sub = PatternFill("solid", fgColor="F7E6D8")
    fill_blue = PatternFill("solid", fgColor="D9E1F2")

    # 扣款块：断更、虚假签收、冲抵、合计
    ded_blocks = ["断更", "虚假签收", "冲抵", "合计"]

    last_col_guess = 5 + len(headers) * 2 + 2 + len(ded_blocks) * 2

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col_guess)
    ws.cell(1, 1).value = "司机工资单汇总（按线路 + 重量档位 + 首票/联单）"
    ws.cell(1, 1).font = title_font
    ws.cell(1, 1).alignment = center

    # A顺号 B快递员 C金额 D件数 E线路
    ws.merge_cells("A3:A5"); ws["A3"].value = "顺号"
    ws.merge_cells("B3:B5"); ws["B3"].value = "快递员"
    ws.merge_cells("C3:D3"); ws["C3"].value = "应付工资"
    ws.merge_cells("C4:C5"); ws["C4"].value = "金额"
    ws.merge_cells("D4:D5"); ws["D4"].value = "件数"
    ws.merge_cells("E3:E5"); ws["E3"].value = "线路"

    start_col = 6  # F
    end_delivery_col = start_col + len(headers) * 2 - 1
    ws.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_delivery_col)
    ws.cell(3, start_col).value = "送货工资"

    col = start_col
    for (bucket, ticket) in headers:
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        ws.cell(4, col).value = f"{bucket}-{ticket}"
        ws.cell(5, col).value = "件数"
        ws.cell(5, col + 1).value = "金额"
        col += 2

    # 送货合计
    ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
    ws.cell(4, col).value = "送货合计"
    ws.cell(5, col).value = "件数"
    ws.cell(5, col + 1).value = "金额"
    delivery_total_cnt_col = col
    delivery_total_amt_col = col + 1
    col += 2

    # 扣款块
    ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + len(ded_blocks) * 2 - 1)
    ws.cell(3, col).value = "工资扣款"

    for i, name in enumerate(ded_blocks):
        ws.merge_cells(start_row=4, start_column=col + i * 2, end_row=4, end_column=col + i * 2 + 1)
        ws.cell(4, col + i * 2).value = name
        ws.cell(5, col + i * 2).value = "件数"
        ws.cell(5, col + i * 2 + 1).value = "金额"

    ded_start_col = col
    ded_total_cnt_col = col + (len(ded_blocks) - 1) * 2
    ded_total_amt_col = ded_total_cnt_col + 1

    # Header styles
    for rr in range(3, 6):
        ws.row_dimensions[rr].height = 22
        for cc in range(1, ded_total_amt_col + 1):
            cell = ws.cell(rr, cc)
            cell.alignment = center
            cell.font = bold
            cell.border = border
            cell.fill = fill_head if rr == 3 else fill_sub

    # Column widths
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 14

    # =========================
    # Write data rows
    # =========================
    row_start = 6

    for idx, (did, dname, route) in enumerate(drivers.itertuples(index=False), start=1):
        r = row_start + idx - 1

        ws.cell(r, 1).value = idx
        ws.cell(r, 2).value = f"{dname} ({did})"
        ws.cell(r, 5).value = route

        c = start_col
        delivery_cnt_cells = []
        delivery_amt_cells = []

        for (bucket, ticket) in headers:
            cnt_val = get_cnt(did, route, bucket, ticket)
            ws.cell(r, c).value = cnt_val

            is_first = (ticket == "首票")
            col_idx = 3 if is_first else 4  # 参数_线路明细: C=3 D=4

            formula = (
                f'=IFERROR('
                f'{get_column_letter(c)}{r}*'
                f'INDEX({SHEET_ROUTE}!$A:$E,'
                f'MATCH($E{r}&"|"&"{bucket}",{SHEET_ROUTE}!$E:$E,0),'
                f'{col_idx}'
                f'),0)'
            )
            ws.cell(r, c + 1).value = formula

            delivery_cnt_cells.append(f"{get_column_letter(c)}{r}")
            delivery_amt_cells.append(f"{get_column_letter(c + 1)}{r}")
            c += 2

        # 送货合计
        ws.cell(r, delivery_total_cnt_col).value = f"=SUM({','.join(delivery_cnt_cells)})"
        ws.cell(r, delivery_total_amt_col).value = f"=SUM({','.join(delivery_amt_cells)})"

        # 扣款：断更/虚假签收/冲抵 + 合计
        did_str = str(did).strip()

        ws.cell(r, ded_start_col + 0).value = claim_cnt.get((did_str, "断更"), 0)
        ws.cell(r, ded_start_col + 1).value = float(claim_amt.get((did_str, "断更"), 0.0))

        ws.cell(r, ded_start_col + 2).value = claim_cnt.get((did_str, "虚假签收"), 0)
        ws.cell(r, ded_start_col + 3).value = float(claim_amt.get((did_str, "虚假签收"), 0.0))

        ws.cell(r, ded_start_col + 4).value = int(offset_cnt.get(did_str, 0))
        ws.cell(r, ded_start_col + 5).value = float(offset_amt.get(did_str, 0.0))

        ws.cell(r, ded_start_col + 6).value = (
            f"=SUM({get_column_letter(ded_start_col)}{r},"
            f"{get_column_letter(ded_start_col+2)}{r},"
            f"{get_column_letter(ded_start_col+4)}{r})"
        )
        ws.cell(r, ded_start_col + 7).value = (
            f"=SUM({get_column_letter(ded_start_col+1)}{r},"
            f"{get_column_letter(ded_start_col+3)}{r},"
            f"{get_column_letter(ded_start_col+5)}{r})"
        )

        # 应付工资：件数=送货合计件数；金额=送货合计金额-扣款合计金额
        ws.cell(r, 4).value = f"={get_column_letter(delivery_total_cnt_col)}{r}"
        ws.cell(r, 3).value = f"={get_column_letter(delivery_total_amt_col)}{r}-{get_column_letter(ded_total_amt_col)}{r}"

        # Row styles
        for cc in range(1, ded_total_amt_col + 1):
            cell = ws.cell(r, cc)
            cell.alignment = center
            cell.border = border
            if cc == 3:
                cell.fill = fill_blue

    ws.freeze_panes = "A6"
    return wb

# =========================
# Session state
# =========================
if "generated" not in st.session_state:
    st.session_state.generated = False
if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "output_name" not in st.session_state:
    st.session_state.output_name = None

# =========================
# Read workbook sheets
# =========================
try:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
except Exception as e:
    st.error(f"无法读取Excel：{e}")
    st.stop()

if SHEET_DELIVERY not in sheet_names:
    st.error(f"缺少工作表：{SHEET_DELIVERY}；当前sheet：{sheet_names}")
    st.stop()

df_delivery = pd.read_excel(uploaded_file, sheet_name=SHEET_DELIVERY)
df_delivery.columns = df_delivery.columns.astype(str).str.strip()

st.subheader("Raw Data Preview (派费明细)")
st.dataframe(df_delivery.head(50), use_container_width=True)

# =========================
# Columns mapping
# =========================
COL_DRIVER_ID = "快递员ID"
COL_DRIVER = "快递员"
COL_ROUTE = "区域/线路"
COL_WEIGHT = "结算重量lb"
COL_TASK = "任务号"
COL_STOP = "STOP序号"

need = [COL_DRIVER_ID, COL_DRIVER, COL_ROUTE, COL_WEIGHT, COL_TASK, COL_STOP]
miss = [c for c in need if c not in df_delivery.columns]
if miss:
    st.error(f"派费明细缺少列: {miss}")
    st.stop()

# =========================
# Auto routes from delivery
# =========================
routes = (
    df_delivery[COL_ROUTE]
    .astype(str)
    .fillna("")
    .str.strip()
)
routes = sorted([r for r in routes.unique().tolist() if r])
if not routes:
    st.error("未从 派费明细 的“区域/线路”提取到任何线路，请检查源表。")
    st.stop()

# =========================
# Fleet & period inputs (default from filename)
# =========================
default_fleet, default_start, default_end = parse_fleet_and_period_from_filename(uploaded_file.name)

with st.expander("账单信息（车队名称 & 帐期范围）", expanded=True):
    col1, col2, col3 = st.columns([2, 1, 1])
    fleet_name = col1.text_input("车队名称", value=default_fleet or "")
    start_date = col2.date_input("帐期开始", value=default_start or date.today())
    end_date = col3.date_input("帐期结束", value=default_end or date.today())

# =========================
# Weight tiers (3 tiers, no '超重件' word)
# =========================
tiers = ["0-5lb", "5-20lb", "20lb以上"]

# =========================
# Price input table
# =========================
st.subheader("线路单价 (输入价格)")

base_rows = []
for r in routes:
    for t in tiers:
        base_rows.append({"线路": r, "档位": t, "首票单价": 0.0, "联单单价": 0.0})

df_price = pd.DataFrame(base_rows)

edited_df = st.data_editor(
    df_price,
    use_container_width=True,
    hide_index=True,
    column_config={
        "线路": st.column_config.TextColumn("线路", disabled=True),
        "档位": st.column_config.TextColumn("档位", disabled=True),
        "首票单价": st.column_config.NumberColumn("首票单价", min_value=0.0, step=0.01, format="%.2f"),
        "联单单价": st.column_config.NumberColumn("联单单价", min_value=0.0, step=0.01, format="%.2f"),
    },
)

# Validate price table
df_route_price = edited_df.copy()
df_route_price.columns = df_route_price.columns.astype(str).str.strip()
need2 = ["线路", "档位", "首票单价", "联单单价"]
miss2 = [c for c in need2 if c not in df_route_price.columns]
if miss2:
    st.error(f"网页单价表缺少列: {miss2}")
    st.stop()

df_route_price["线路"] = df_route_price["线路"].astype(str).str.strip()
df_route_price["档位"] = df_route_price["档位"].astype(str).str.strip()
df_route_price["首票单价"] = pd.to_numeric(df_route_price["首票单价"], errors="coerce")
df_route_price["联单单价"] = pd.to_numeric(df_route_price["联单单价"], errors="coerce")

missing_price = df_route_price["首票单价"].isna() | df_route_price["联单单价"].isna()
if missing_price.any():
    st.warning("单价未填写完整或存在非数字：请补全 首票单价 / 联单单价（否则无法生成账单）")

# =========================
# Read offset & claim sheets
# =========================
df_offset = pd.DataFrame()
if SHEET_OFFSET in sheet_names:
    df_offset = pd.read_excel(uploaded_file, sheet_name=SHEET_OFFSET)
    df_offset.columns = df_offset.columns.astype(str).str.strip()

df_claim_raw = None
if SHEET_CLAIM in sheet_names:
    df_claim_raw = pd.read_excel(uploaded_file, sheet_name=SHEET_CLAIM)
    df_claim_raw.columns = df_claim_raw.columns.astype(str).str.strip()

# =========================
# Confirm generate button
# =========================
btn_col1, btn_col2 = st.columns([1, 2])

with btn_col1:
    generate_clicked = st.button("✅ 确认生成账单", type="primary", use_container_width=True)

if generate_clicked:
    # hard validation before generate
    if missing_price.any():
        st.error("单价未填写完整，无法生成账单。请先补全单价。")
        st.stop()

    if not fleet_name.strip():
        st.error("车队名称不能为空（用于输出文件名）。")
        st.stop()

    if start_date > end_date:
        st.error("帐期开始不能晚于帐期结束。")
        st.stop()

    # Build workbook
    wb = build_workbook(
        df_delivery=df_delivery,
        df_route_price=df_route_price,
        df_offset=df_offset,
        df_claim_raw=df_claim_raw,
        sheet_names=sheet_names,
    )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    period_str = fmt_period(start_date, end_date)
    out_name = f"{safe_filename(fleet_name)}-{period_str}-工资单.xlsx"

    st.session_state.generated = True
    st.session_state.output_bytes = buffer.getvalue()
    st.session_state.output_name = out_name

    st.success("账单已生成，可以下载了。")

# =========================
# Download button (turn green after generated)
# =========================
if st.session_state.generated and st.session_state.output_bytes:
    st.markdown(
        """
        <style>
        /* Make download button green */
        div[data-testid="stDownloadButton"] button {
            background-color: #16a34a !important; /* green */
            color: white !important;
            border: 1px solid #15803d !important;
        }
        div[data-testid="stDownloadButton"] button:hover {
            background-color: #15803d !important;
            border-color: #166534 !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.download_button(
        label="⬇️ 下载工资单",
        data=st.session_state.output_bytes,
        file_name=st.session_state.output_name or "工资单.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("请先填写单价，然后点击「确认生成账单」。生成后会出现绿色下载按钮。")
