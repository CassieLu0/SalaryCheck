import re
from datetime import datetime
from io import BytesIO

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# =========================
# Page
# =========================
st.set_page_config(page_title="DSP Salary Summary Tool", layout="wide")
st.title("DSP Salary Summary Tool")

# =========================
# Constants
# =========================
SHEET_DELIVERY = "派费明细"
SHEET_ROUTE = "参数_线路明细"
SHEET_OFFSET = "冲抵明细"
SHEET_CLAIM = "理赔明细"

COL_DRIVER_ID = "快递员ID"
COL_DRIVER = "快递员"
COL_ROUTE = "区域/线路"
COL_WEIGHT = "结算重量lb"
COL_TASK = "任务号"
COL_STOP = "STOP序号"

# Weight tiers (no "超重件" word)
TIERS = ["0-5lb", "5-20lb", "20lb以上"]
BUCKETS = [
    ("0-5lb", 0, 5),         # [0,5)
    ("5-20lb", 5, 20),       # [5,20)
    ("20lb以上", 20, None),   # [20, +inf)
]

# =========================
# Helpers
# =========================
def safe_filename(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", "", s)
    return s or "工资单"

def parse_fleet_and_period_from_filename(name: str):
    """
    filename example:
      SEA-ZEP-派费明细-09_02_2026-15_02_2026-GOFO.xlsx

    date format supports:
      - MM_DD_YYYY
      - DD_MM_YYYY  (your sample)
    """
    if not name:
        return None, None

    m_fleet = re.search(r"^(.*?)-派费明细-", name)
    m_period = re.search(r"(\d{2}_\d{2}_\d{4})-(\d{2}_\d{2}_\d{4})", name)

    fleet = m_fleet.group(1).strip() if m_fleet else None
    if not m_period:
        return fleet, None

    p1, p2 = m_period.group(1), m_period.group(2)

    ok = False
    for fmt in ("%m_%d_%Y", "%d_%m_%Y"):
        try:
            datetime.strptime(p1, fmt)
            datetime.strptime(p2, fmt)
            ok = True
            break
        except Exception:
            pass

    period = f"{p1}-{p2}" if ok else None
    return fleet, period

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
# Workbook builder
# =========================
def build_workbook(
    df_delivery: pd.DataFrame,
    df_route_price: pd.DataFrame,
    df_offset: pd.DataFrame,
    df_claim_raw: pd.DataFrame | None,
    max_claim_types: int = 20,   # 防止理赔类型太多撑爆表头
):
    # -------------------------
    # 1) 基础校验 & 预处理
    # -------------------------
    must_cols = [COL_DRIVER_ID, COL_DRIVER, COL_ROUTE, COL_WEIGHT, COL_TASK, COL_STOP]
    miss = [c for c in must_cols if c not in df_delivery.columns]
    if miss:
        raise ValueError(f"派费明细缺少列: {miss}")

    df = df_delivery.copy()
    df[COL_DRIVER_ID] = df[COL_DRIVER_ID].astype(str).str.strip()
    df[COL_DRIVER] = df[COL_DRIVER].astype(str).str.strip()
    df["_route"] = df[COL_ROUTE].astype(str).fillna("").str.strip()
    df["_weight"] = pd.to_numeric(df[COL_WEIGHT], errors="coerce").fillna(0.0)
    df["_bucket"] = df["_weight"].apply(weight_bucket)

    gkey = [COL_DRIVER_ID, COL_TASK, COL_STOP]
    df["_rank_in_stop"] = df.groupby(gkey).cumcount() + 1
    df["_ticket"] = df["_rank_in_stop"].eq(1).map({True: "首票", False: "联单"})

    cnt = (
        df.groupby([COL_DRIVER_ID, COL_DRIVER, "_route", "_bucket", "_ticket"], dropna=False)
          .size().reset_index(name="件数")
    )

    def get_cnt(did, route, bucket, ticket):
        m = cnt[
            (cnt[COL_DRIVER_ID] == did) &
            (cnt["_route"] == route) &
            (cnt["_bucket"] == bucket) &
            (cnt["_ticket"] == ticket)
        ]
        return 0 if m.empty else int(m.iloc[0]["件数"])

    # -------------------------
    # 2) 冲抵汇总（按司机ID）
    # -------------------------
    offset_cnt, offset_amt = {}, {}
    offset_dids = set()

    if df_offset is not None and not df_offset.empty:
        need_off = ["快递员ID", "费用合计_未税"]
        miss_off = [c for c in need_off if c not in df_offset.columns]
        if miss_off:
            raise ValueError(f"冲抵明细缺少列: {miss_off}")

        dfo = df_offset.copy()
        dfo["快递员ID"] = dfo["快递员ID"].astype(str).str.strip()
        dfo["费用合计_未税"] = pd.to_numeric(dfo["费用合计_未税"], errors="coerce").fillna(0.0)

        g_off = dfo.groupby("快递员ID", dropna=False).agg(
            cnt=("费用合计_未税", "size"),
            amt=("费用合计_未税", "sum"),
        )
        offset_cnt = g_off["cnt"].to_dict()
        offset_amt = g_off["amt"].to_dict()
        offset_dids = set(g_off.index.astype(str).str.strip().tolist())

    # -------------------------
    # 3) 理赔汇总（按司机ID；理赔类型动态列）
    #    - 不再丢弃未知类型
    #    - 匹配不到ID的输出到 sheet：理赔_未匹配
    # -------------------------
    claim_cnt, claim_amt = {}, {}
    claim_dids = set()
    unmatched_claim = pd.DataFrame()

    # name -> most common ID (来自派费明细)
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

    # 你原来只映射两个类型，这里保留映射，但未知类型不丢
    def normalize_claim_type(x: str) -> str:
        x = str(x).strip()
        if x == "轨迹断更":
            return "断更"
        if x == "虚假签收":
            return "虚假签收"
        # 其他类型：保留原字样（或你也可以 return "其他"）
        return x or "其他"

    if df_claim_raw is not None and not df_claim_raw.empty:
        need_cl = ["快递员", "理赔类型", "费用合计_未税"]
        miss_cl = [c for c in need_cl if c not in df_claim_raw.columns]
        if miss_cl:
            raise ValueError(f"理赔明细缺少列: {miss_cl}")

        dfc = df_claim_raw.copy()
        dfc["_name"] = dfc["快递员"].astype(str).str.strip()
        dfc["_dtype"] = dfc["理赔类型"].apply(normalize_claim_type)
        dfc["_did"] = dfc["_name"].map(name_to_id)
        dfc["费用合计_未税"] = pd.to_numeric(dfc["费用合计_未税"], errors="coerce").fillna(0.0)

        # 未匹配到ID的，单独留表，不 stop
        unmatched_claim = dfc[dfc["_did"].isna()].copy()

        dfc_ok = dfc[dfc["_did"].notna()].copy()
        dfc_ok["_did"] = dfc_ok["_did"].astype(str).str.strip()

        g_cl = dfc_ok.groupby(["_did", "_dtype"], dropna=False).agg(
            cnt=("费用合计_未税", "size"),
            amt=("费用合计_未税", "sum"),
        )

        for (did, dtype), row in g_cl.iterrows():
            did_s = str(did).strip()
            dtype_s = str(dtype).strip()
            claim_cnt[(did_s, dtype_s)] = int(row["cnt"])
            claim_amt[(did_s, dtype_s)] = float(row["amt"])

        claim_dids = set(dfc_ok["_did"].unique().tolist())

    # 动态理赔类型列表（限制数量）
    claim_types = sorted({dtype for (_, dtype) in claim_cnt.keys()})
    if len(claim_types) > max_claim_types:
        # 超过上限：保留金额最高的前N类，其余归为“其他”
        # 做法：按类型汇总金额排序
        type_sum = {}
        for (did, dtype), amt in claim_amt.items():
            type_sum[dtype] = type_sum.get(dtype, 0.0) + float(amt)
        claim_types = [t for t, _ in sorted(type_sum.items(), key=lambda x: x[1], reverse=True)[:max_claim_types]]
        # 重新把不在前N的合并到“其他”
        for (did, dtype), c in list(claim_cnt.items()):
            if dtype not in claim_types:
                claim_cnt[(did, "其他")] = claim_cnt.get((did, "其他"), 0) + c
                claim_amt[(did, "其他")] = claim_amt.get((did, "其他"), 0.0) + claim_amt.get((did, dtype), 0.0)
                del claim_cnt[(did, dtype)]
                del claim_amt[(did, dtype)]
        if "其他" not in claim_types:
            claim_types.append("其他")

    # -------------------------
    # 4) drivers 集合：派费司机 ∪ 冲抵司机 ∪ 理赔司机
    #    并且：扣款只写到该司机第一条线路行，避免重复
    # -------------------------
    delivery_drivers = (
        df[[COL_DRIVER_ID, COL_DRIVER, "_route"]]
        .drop_duplicates()
        .sort_values(["_route", COL_DRIVER])
    )

    delivery_dids = set(delivery_drivers[COL_DRIVER_ID].astype(str).str.strip().tolist())
    all_dids = sorted(delivery_dids | offset_dids | claim_dids)

    # build rows for output:
    # - 先放派费里的（保持原有按线路的多行）
    drivers_rows = delivery_drivers.copy()

    # - 补充：只存在扣款但派费没有的司机，给一行 route=""
    #   name 用派费映射，如果完全没有则显示 did
    extra_dids = [d for d in all_dids if d not in delivery_dids]
    if extra_dids:
        # did -> most common name from delivery if exists
        did_to_name = (
            df_delivery[[COL_DRIVER_ID, COL_DRIVER]]
            .dropna()
            .assign(**{
                COL_DRIVER_ID: lambda x: x[COL_DRIVER_ID].astype(str).str.strip(),
                COL_DRIVER: lambda x: x[COL_DRIVER].astype(str).str.strip(),
            })
            .groupby(COL_DRIVER_ID)[COL_DRIVER]
            .agg(lambda s: s.value_counts().index[0])
            .to_dict()
        )
        extra = pd.DataFrame({
            COL_DRIVER_ID: extra_dids,
            COL_DRIVER: [did_to_name.get(d, f"UNKNOWN") for d in extra_dids],
            "_route": ["" for _ in extra_dids],
        })
        drivers_rows = pd.concat([drivers_rows, extra], ignore_index=True)

    # 标记每个司机的“第一行”用于写扣款（避免重复）
    drivers_rows["_is_first_row"] = drivers_rows.groupby(COL_DRIVER_ID).cumcount().eq(0)

    # -------------------------
    # 5) Workbook + 参数_线路明细
    # -------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "司机汇总"

    pr = df_route_price.copy()
    pr["线路"] = pr["线路"].astype(str).str.strip()
    pr["档位"] = pr["档位"].astype(str).str.strip()
    pr["首票单价"] = pd.to_numeric(pr["首票单价"], errors="coerce")
    pr["联单单价"] = pd.to_numeric(pr["联单单价"], errors="coerce")
    if pr[["首票单价", "联单单价"]].isna().any().any():
        raise ValueError("单价表存在空值/非数字，无法生成。")
    pr["匹配键"] = pr["线路"] + "|" + pr["档位"]

    ws_route = wb.create_sheet(SHEET_ROUTE)
    ws_route.append(["线路", "档位", "首票单价", "联单单价", "匹配键"])
    for cell in ws_route[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for _, row in pr.iterrows():
        ws_route.append([row["线路"], row["档位"], float(row["首票单价"]), float(row["联单单价"]), row["匹配键"]])

    # 输出原始明细（方便对账）
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
    if unmatched_claim is not None and not unmatched_claim.empty:
        write_df_to_sheet(wb, "理赔_未匹配", unmatched_claim)

    # -------------------------
    # 6) 样式 & 表头（扣款块动态：理赔类型 + 冲抵 + 合计）
    # -------------------------
    thin = Side(style="thin", color="666666")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bold = Font(bold=True)
    title_font = Font(bold=True, size=16)

    fill_head = PatternFill("solid", fgColor="F2D6C9")
    fill_sub = PatternFill("solid", fgColor="F7E6D8")
    fill_blue = PatternFill("solid", fgColor="D9E1F2")

    # 送货 headers
    headers = []
    for b in [x[0] for x in BUCKETS]:
        headers.append((b, "首票"))
        headers.append((b, "联单"))

    # 扣款块：动态理赔类型 + 冲抵 + 合计
    ded_blocks = claim_types + ["冲抵", "合计"]

    last_col_guess = 5 + len(headers) * 2 + 2 + len(ded_blocks) * 2

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col_guess)
    ws.cell(1, 1).value = "司机工资单汇总（按线路 + 重量档位 + 首票/联单）"
    ws.cell(1, 1).font = title_font
    ws.cell(1, 1).alignment = center

    ws.merge_cells("A3:A5"); ws["A3"].value = "顺号"
    ws.merge_cells("B3:B5"); ws["B3"].value = "快递员"
    ws.merge_cells("C3:D3"); ws["C3"].value = "应付工资"
    ws.merge_cells("C4:C5"); ws["C4"].value = "金额"
    ws.merge_cells("D4:D5"); ws["D4"].value = "件数"
    ws.merge_cells("E3:E5"); ws["E3"].value = "线路"

    start_col = 6
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

    ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
    ws.cell(4, col).value = "送货合计"
    ws.cell(5, col).value = "件数"
    ws.cell(5, col + 1).value = "金额"
    delivery_total_cnt_col = col
    delivery_total_amt_col = col + 1
    col += 2

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

    for rr in range(3, 6):
        ws.row_dimensions[rr].height = 22
        for cc in range(1, ded_total_amt_col + 1):
            cell = ws.cell(rr, cc)
            cell.alignment = center
            cell.font = bold
            cell.border = border
            cell.fill = fill_head if rr == 3 else fill_sub

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 14

    # -------------------------
    # 7) 写行：扣款只写到该司机第一行
    # -------------------------
    row_start = 6

    for idx, row in enumerate(drivers_rows.itertuples(index=False), start=1):
        did = str(getattr(row, COL_DRIVER_ID)).strip()
        dname = str(getattr(row, COL_DRIVER)).strip()
        route = str(getattr(row, "_route")).strip()
        is_first_row = bool(getattr(row, "_is_first_row"))

        r = row_start + idx - 1

        ws.cell(r, 1).value = idx
        ws.cell(r, 2).value = f"{dname} ({did})"
        ws.cell(r, 5).value = route

        # 送货块
        c = start_col
        delivery_cnt_cells = []
        delivery_amt_cells = []

        for (bucket, ticket) in headers:
            cnt_val = get_cnt(did, route, bucket, ticket) if route else 0
            ws.cell(r, c).value = cnt_val

            col_idx = 3 if ticket == "首票" else 4
            formula = (
                f'=IFERROR('
                f'{get_column_letter(c)}{r}*'
                f'INDEX({SHEET_ROUTE}!$A:$E,'
                f'MATCH($E{r}&"|"&"{bucket}",{SHEET_ROUTE}!$E:$E,0),'
                f'{col_idx}'
                f'),0)'
            )
            ws.cell(r, c + 1).value = formula if route else 0

            delivery_cnt_cells.append(f"{get_column_letter(c)}{r}")
            delivery_amt_cells.append(f"{get_column_letter(c + 1)}{r}")
            c += 2

        ws.cell(r, delivery_total_cnt_col).value = f"=SUM({','.join(delivery_cnt_cells)})"
        ws.cell(r, delivery_total_amt_col).value = f"=SUM({','.join(delivery_amt_cells)})"

        # 扣款块：只在第一行写入，其他线路行置 0（避免重复）
        # ded_blocks = claim_types + ["冲抵", "合计"]
        # 对每个类型写 cnt/amt
        write_col = ded_start_col
        if is_first_row:
            # 理赔类型
            for t in claim_types:
                ws.cell(r, write_col).value = claim_cnt.get((did, t), 0)
                ws.cell(r, write_col + 1).value = float(claim_amt.get((did, t), 0.0))
                write_col += 2

            # 冲抵
            ws.cell(r, write_col).value = int(offset_cnt.get(did, 0))
            ws.cell(r, write_col + 1).value = float(offset_amt.get(did, 0.0))
            write_col += 2

            # 合计（所有扣款金额合并）
            # cnt 合计：理赔件数+冲抵条数
            cnt_terms = []
            amt_terms = []
            for j in range(ded_start_col, ded_start_col + (len(ded_blocks) - 1) * 2, 2):
                cnt_terms.append(f"{get_column_letter(j)}{r}")
                amt_terms.append(f"{get_column_letter(j+1)}{r}")

            ws.cell(r, ded_total_cnt_col).value = f"=SUM({','.join(cnt_terms)})"
            ws.cell(r, ded_total_amt_col).value = f"=SUM({','.join(amt_terms)})"
        else:
            # 非第一行：全部扣款置 0
            for k in range(ded_start_col, ded_total_amt_col + 1):
                ws.cell(r, k).value = 0

        # 应付工资：件数=送货合计件数；金额=送货合计金额-扣款合计金额
        ws.cell(r, 4).value = f"={get_column_letter(delivery_total_cnt_col)}{r}"
        ws.cell(r, 3).value = f"={get_column_letter(delivery_total_amt_col)}{r}-{get_column_letter(ded_total_amt_col)}{r}"

        # 样式
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
st.session_state.setdefault("generated", False)
st.session_state.setdefault("output_bytes", None)
st.session_state.setdefault("output_name", None)
st.session_state.setdefault("last_uploaded_name", None)

# =========================
# Upload
# =========================
uploaded_file = st.file_uploader("Upload salary Excel (xlsx)", type=["xlsx"])
if uploaded_file is None:
    st.info("请先上传 xlsx 文件。")
    st.stop()

# New file => reset
if st.session_state["last_uploaded_name"] != uploaded_file.name:
    st.session_state["last_uploaded_name"] = uploaded_file.name
    st.session_state["generated"] = False
    st.session_state["output_bytes"] = None
    st.session_state["output_name"] = None

# Parse fleet & period (read-only)
fleet_name, period_str = parse_fleet_and_period_from_filename(uploaded_file.name)
if not fleet_name or not period_str:
    st.error(
        "文件名无法解析【车队名称/帐期范围】。\n"
        "请确保文件名格式：<车队>-派费明细-MM_DD_YYYY-MM_DD_YYYY-...xlsx（也支持DD_MM_YYYY）\n"
        f"当前文件名：{uploaded_file.name}"
    )
    st.stop()

st.caption(f"车队：{fleet_name} ｜ 帐期：{period_str}")

# =========================
# Only extract routes (light read)
# =========================
try:
    df_routes_only = pd.read_excel(uploaded_file, sheet_name=SHEET_DELIVERY, usecols=[COL_ROUTE])
    df_routes_only.columns = df_routes_only.columns.astype(str).str.strip()
except Exception as e:
    st.error(f"无法读取Excel或缺少列 {COL_ROUTE}：{e}")
    st.stop()

routes = sorted([r for r in df_routes_only[COL_ROUTE].astype(str).fillna("").str.strip().unique().tolist() if r])
if not routes:
    st.error("未从 派费明细 的“区域/线路”提取到任何线路，请检查源表。")
    st.stop()

st.caption(f"已识别线路数：{len(routes)}")

# =========================
# Price form (submit then run heavy)
# =========================
base_rows = [{"线路": r, "档位": t, "首票单价": 0.0, "联单单价": 0.0} for r in routes for t in TIERS]
df_price = pd.DataFrame(base_rows)

with st.form("price_form", clear_on_submit=False):
    st.subheader("线路单价（填完后点确认才生成账单）")
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
    submit = st.form_submit_button("✅ 确认生成账单", type="primary", use_container_width=True)

# =========================
# Heavy work only after submit
# =========================
if submit:
    df_route_price = edited_df.copy()
    df_route_price["首票单价"] = pd.to_numeric(df_route_price["首票单价"], errors="coerce")
    df_route_price["联单单价"] = pd.to_numeric(df_route_price["联单单价"], errors="coerce")

    missing_price = df_route_price["首票单价"].isna() | df_route_price["联单单价"].isna()
    if missing_price.any():
        st.error("单价未填写完整或存在非数字，无法生成账单。")
        st.stop()

    # Read full workbook
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

    df_offset = pd.DataFrame()
    if SHEET_OFFSET in sheet_names:
        df_offset = pd.read_excel(uploaded_file, sheet_name=SHEET_OFFSET)
        df_offset.columns = df_offset.columns.astype(str).str.strip()

    df_claim_raw = None
    if SHEET_CLAIM in sheet_names:
        df_claim_raw = pd.read_excel(uploaded_file, sheet_name=SHEET_CLAIM)
        df_claim_raw.columns = df_claim_raw.columns.astype(str).str.strip()

    try:
        wb = build_workbook(
            df_delivery=df_delivery,
            df_route_price=df_route_price,
            df_offset=df_offset,
            df_claim_raw=df_claim_raw,
        )
    except Exception as e:
        st.error(f"生成失败：{e}")
        st.stop()

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    out_name = f"{safe_filename(fleet_name)}-{period_str}-工资单.xlsx"
    st.session_state["generated"] = True
    st.session_state["output_bytes"] = buf.getvalue()
    st.session_state["output_name"] = out_name

    st.success("账单已生成，可以下载了。")

# =========================
# Download (green after generated)
# =========================
if st.session_state.get("generated") and st.session_state.get("output_bytes"):
    st.markdown(
        """
        <style>
        div[data-testid="stDownloadButton"] button {
            background-color: #16a34a !important;
            color: white !important;
            border: 1px solid #15803d !important;
        }
        div[data-testid="stDownloadButton"] button:hover {
            background-color: #15803d !important;
            border-color: #166534 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.download_button(
        label="⬇️ 下载工资单",
        data=st.session_state["output_bytes"],
        file_name=st.session_state["output_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("请先填写单价，然后点击「确认生成账单」。生成后会出现绿色下载按钮。")
