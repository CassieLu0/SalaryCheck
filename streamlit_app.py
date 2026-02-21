import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

st.title("DSP Salary Summary Tool")

# upload file
uploaded_file = st.file_uploader("Upload salary Excel", type=["xlsx", 'csv'])

if uploaded_file is None:
    st.info("Please upload an Excel or csv file to continue.")
    st.stop()

SHEET_DELIVERY = "派费明细"
SHEET_ROUTE    = "参数_线路明细"

st.subheader("线路单价 (输入价格)")


# 1) define the fixed structure (线路 x 档位)
routes = ["SEA01-004", "SEA01-005"]
tiers = ["6磅以下", "6-10磅", "10-20磅", "20磅以上", "超重件"]

# 2) build the table like Excel
base_rows = []
for r in routes:
    for t in tiers:
        base_rows.append({"线路": r, "档位": t, "首票单价": None, "联单单价": None})

# 3) editable grid, only allow editing the 2 price columns
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
df_route_price = edited_df.copy()
df_route_price.columns = df_route_price.columns.astype(str).str.strip()

st.dataframe(df_route_price[["线路","档位","首票单价","联单单价"]])

if uploaded_file.name.endswith(".csv"):
    st.error("你的逻辑依赖 Excel 多 sheet，csv 不支持，请上传 xlsx")
    st.stop()

df = pd.read_excel(uploaded_file, sheet_name=SHEET_DELIVERY)
df.columns = df.columns.astype(str).str.strip()

st.subheader("Raw Data Preview (派费明细)")
st.dataframe(df.head(50))


# full logic here

df = pd.read_excel(uploaded_file, sheet_name="派费明细")
# =========================
# 1) Sheet 名
# =========================
SHEET_DELIVERY = "派费明细"
SHEET_ROUTE    = "参数_线路明细"   # 你已经加了

df_delivery = pd.read_excel(uploaded_file, sheet_name=SHEET_DELIVERY)
df_route_price = edited_df.copy()

# =========================
# 2) 派费明细列名（不一致就在这里改）
# =========================
COL_DRIVER_ID = "快递员ID"
COL_DRIVER    = "快递员"
COL_ROUTE     = "区域/线路"        # 你说 _route 就是这个
COL_WEIGHT    = "结算重量lb"
COL_TASK      = "任务号"
COL_STOP      = "STOP序号"


# =========================
# 3) 重量档位（只负责分桶，不写单价）
# =========================
BUCKETS = [
    ("6磅以下", 0, 6),
    ("6-10磅", 6, 10),
    ("10-20磅", 10, 20),
    ("20磅以上", 20, None),
]
OVERWEIGHT_LB = 30.0
OVERWEIGHT_BUCKET = "超重件"


def weight_bucket(w):
    try:
        w = float(w)
    except Exception:
        w = 0.0

    if w > OVERWEIGHT_LB:
        return OVERWEIGHT_BUCKET

    for name, lb_min, lb_max in BUCKETS:
        if lb_max is None:
            if w >= lb_min:
                return name
        else:
            if (w >= lb_min) and (w < lb_max):
                return name
    return BUCKETS[-1][0]


# =========================
# 4) 读数据
# =========================
df = pd.read_excel(uploaded_file, sheet_name=SHEET_DELIVERY)
df.columns = df.columns.astype(str).str.strip()

need = [COL_DRIVER_ID, COL_DRIVER, COL_ROUTE, COL_WEIGHT, COL_TASK, COL_STOP]
miss = [c for c in need if c not in df.columns]
if miss:
    raise ValueError(f"派费明细缺少列: {miss}")

need2 = ["线路", "档位", "首票单价", "联单单价"]
miss2 = [c for c in need2 if c not in df_route_price.columns]
if miss2:
    raise ValueError(f"网页单价表缺少列: {miss2}")

df[COL_DRIVER_ID] = df[COL_DRIVER_ID].astype(str).str.strip()
df[COL_DRIVER]    = df[COL_DRIVER].astype(str).str.strip()
df["_route"]      = df[COL_ROUTE].astype(str).fillna("DEFAULT").str.strip()
df["_weight"]     = pd.to_numeric(df[COL_WEIGHT], errors="coerce").fillna(0.0)
df["_bucket"]     = df["_weight"].apply(weight_bucket)

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
for b in [x[0] for x in BUCKETS] + [OVERWEIGHT_BUCKET]:
    headers.append((b, "首票"))
    headers.append((b, "联单"))

# =========================
# 5) 建 Workbook（这段一定要在最前）
# =========================
wb = Workbook()
ws = wb.active
ws.title = "司机汇总"

ws_route = wb.create_sheet("参数_线路明细")
ws_route.append(["线路", "档位", "首票单价", "联单单价", "匹配键"])
for cell in ws_route[1]:
    cell.font =  Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for _, row in df_route_price.iterrows():
    ws_route.append([
        row["线路"],
        row["档位"],
        float(row["首票单价"]),
        float(row["联单单价"]),
        row["匹配键"],
    ])


# =========================
# 6) 样式
# =========================
thin = Side(style="thin", color="666666")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
bold = Font(bold=True)
title_font = Font(bold=True, size=16)

fill_head = PatternFill("solid", fgColor="F2D6C9")
fill_sub  = PatternFill("solid", fgColor="F7E6D8")
fill_blue = PatternFill("solid", fgColor="D9E1F2")



# =========================
# 8) 写 参数_线路明细（真正单价来源）
# =========================
df_route_price = edited_df.copy()

df_route_price["线路"] = df_route_price["线路"].astype(str).str.strip()
df_route_price["档位"] = df_route_price["档位"].astype(str).str.strip()
df_route_price["首票单价"] = pd.to_numeric(df_route_price["首票单价"], errors="coerce")
df_route_price["联单单价"] = pd.to_numeric(df_route_price["联单单价"], errors="coerce")

missing_price = df_route_price["首票单价"].isna() | df_route_price["联单单价"].isna()
if missing_price.any():
    st.error("单价未填写完整，请补全 首票单价 和 联单单价")
    st.stop()

df_route_price["匹配键"] = df_route_price["线路"] + "|" + df_route_price["档位"]

# =========================
# 9) 表头（司机汇总）
# =========================
last_col_guess = 5 + len(headers) * 2 + 2 + 6
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

start_col = 6  # F 开始
end_delivery_col = start_col + len(headers) * 2 - 1
ws.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_delivery_col)
ws.cell(3, start_col).value = "送货工资"

col = start_col
for (bucket, ticket) in headers:
    ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+1)
    ws.cell(4, col).value = f"{bucket}-{ticket}"
    ws.cell(5, col).value = "件数"
    ws.cell(5, col+1).value = "金额"
    col += 2

# 送货合计
ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+1)
ws.cell(4, col).value = "送货合计"
ws.cell(5, col).value = "件数"
ws.cell(5, col+1).value = "金额"
delivery_total_cnt_col = col
delivery_total_amt_col = col + 1
col += 2

# 扣款块（先占位，后续你要接冲抵/理赔再做）
ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col+5)
ws.cell(3, col).value = "工资扣款"

ded_blocks = ["断更", "虚假签收", "合计"]
for i, name in enumerate(ded_blocks):
    ws.merge_cells(start_row=4, start_column=col + i*2, end_row=4, end_column=col + i*2 + 1)
    ws.cell(4, col + i*2).value = name
    ws.cell(5, col + i*2).value = "件数"
    ws.cell(5, col + i*2 + 1).value = "金额"

ded_start_col = col
ded_total_cnt_col = col + 4
ded_total_amt_col = col + 5

# 表头样式
for r in range(3, 6):
    ws.row_dimensions[r].height = 22
    for c in range(1, ded_total_amt_col + 1):
        cell = ws.cell(r, c)
        cell.alignment = center
        cell.font = bold
        cell.border = border
        cell.fill = fill_head if r == 3 else fill_sub

# 列宽
ws.column_dimensions["A"].width = 6
ws.column_dimensions["B"].width = 28
ws.column_dimensions["C"].width = 14
ws.column_dimensions["D"].width = 10
ws.column_dimensions["E"].width = 14

# =========================
# 10) 写数据行（核心：金额用安全公式 INDEX+MATCH）
# =========================
row_start = 6

for idx, (did, dname, route) in enumerate(drivers.itertuples(index=False), start=1):
    r = row_start + idx - 1

    ws.cell(r, 1).value = idx
    ws.cell(r, 2).value = f"{dname} ({did})"
    ws.cell(r, 5).value = route  # 线路写入 E 列

    c = start_col
    delivery_cnt_cells = []
    delivery_amt_cells = []

    for (bucket, ticket) in headers:
        cnt_val = get_cnt(did, route, bucket, ticket)
        ws.cell(r, c).value = cnt_val

        is_first = (ticket == "首票")
        col_idx = 3 if is_first else 4  # 参数_线路明细: C=3 D=4

        # =IFERROR( 件数 * INDEX(A:E, MATCH(线路|档位, E:E,0), col_idx ), 0 )
        formula = (
            f'=IFERROR('
            f'{get_column_letter(c)}{r}*'
            f'INDEX(参数_线路明细!$A:$E,'
            f'MATCH($E{r}&"|"&"{bucket}",参数_线路明细!$E:$E,0),'
            f'{col_idx}'
            f'),0)'
        )
        ws.cell(r, c+1).value = formula

        delivery_cnt_cells.append(f"{get_column_letter(c)}{r}")
        delivery_amt_cells.append(f"{get_column_letter(c+1)}{r}")

        c += 2

    # 送货合计
    ws.cell(r, delivery_total_cnt_col).value = f"=SUM({','.join(delivery_cnt_cells)})"
    ws.cell(r, delivery_total_amt_col).value = f"=SUM({','.join(delivery_amt_cells)})"

    # 扣款先置 0
    for k in range(ded_start_col, ded_total_amt_col + 1):
        ws.cell(r, k).value = 0

    # 应付工资：件数=送货合计件数，金额=送货合计金额-扣款合计金额
    ws.cell(r, 4).value = f"={get_column_letter(delivery_total_cnt_col)}{r}"
    ws.cell(r, 3).value = f"={get_column_letter(delivery_total_amt_col)}{r}-{get_column_letter(ded_total_amt_col)}{r}"

    # 行样式
    for cc in range(1, ded_total_amt_col + 1):
        cell = ws.cell(r, cc)
        cell.alignment = center
        cell.border = border
        if cc == 3:
            cell.fill = fill_blue


# =========================
# 11) 冻结与保存
# =========================
ws.freeze_panes = "A6"
# save into memory instead of local path
buffer = BytesIO()
wb.save(buffer)
buffer.seek(0)

# download button replaces print
st.download_button(
    label="Download 工资单",
    data=buffer,
    file_name="工资单.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
