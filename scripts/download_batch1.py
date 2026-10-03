from pathlib import Path

import pandas as pd

# 统一输出目录：项目根目录下的 data/output/
OUT_DIR = Path(__file__).resolve().parents[1] / "data" / "output"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# =========================================
# 1. FRED 数据序列
# =========================================

series = {
    "PPI": "CHNPIEATI01GYM",          # 中国PPI同比
    "CPI": "CHNCPIALLMINMEI",         # 中国CPI
    "汇率": "EXCHUS",                 # 人民币兑美元
    "出口": "CHNXTEXVA01NCMLM",       # 中国出口，人民币，月度
    "进口": "CHNXTIMVA01NCMLM",       # 中国进口，人民币，月度
    "铜价": "PCOPPUSDM",               # 全球铜价
    "铁矿石": "PIORECRUSDM",           # 全球铁矿石价格
}

# =========================================
# 2. 下载月度/季度等数据
# =========================================

data_all = None

for name, code in series.items():

    print(f"正在下载：{name} ({code})")

    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={code}"

    data = pd.read_csv(url)

    # 第一列日期
    data = data.rename(columns={
        data.columns[0]: "date",
        data.columns[1]: name
    })

    data["date"] = pd.to_datetime(data["date"])

    # 转成数值
    data[name] = pd.to_numeric(data[name], errors="coerce")

    # 如果不是月度数据，统一到月度
    data = (
        data.set_index("date")
        .resample("MS")
        .mean()
        .reset_index()
    )

    # 合并
    if data_all is None:
        data_all = data
    else:
        data_all = pd.merge(
            data_all,
            data,
            on="date",
            how="outer"
        )

# =========================================
# 3. 下载 Brent 原油（日度）
# =========================================

print("正在下载：Brent原油")

url = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DCOILBRENTEU"

oil = pd.read_csv(url)

oil = oil.rename(columns={
    oil.columns[0]: "date",
    oil.columns[1]: "原油"
})

oil["date"] = pd.to_datetime(oil["date"])
oil["原油"] = pd.to_numeric(oil["原油"], errors="coerce")

# 日度 → 月度平均
oil = (
    oil.set_index("date")
    .resample("MS")
    .mean()
    .reset_index()
)

data_all = pd.merge(
    data_all,
    oil,
    on="date",
    how="outer"
)

# =========================================
# 4. 排序
# =========================================

data_all = data_all.sort_values("date")

# =========================================
# 5. 只保留 2006 年以后
# =========================================

data_test = data_all[
    data_all["date"] >= "2006-01-01"
].copy()

# =========================================
# 6. 数据完整性检查
# =========================================

result = []

for col in data_test.columns:

    if col == "date":
        continue

    s = data_test[col]

    result.append({
        "变量": col,
        "开始日期": s.first_valid_index(),
        "结束日期": s.last_valid_index(),
        "有效观测数": s.notna().sum(),
        "缺失数量": s.isna().sum(),
        "缺失比例": round(s.isna().mean(), 4)
    })

check = pd.DataFrame(result)

# 日期索引换成实际日期
for col in ["开始日期", "结束日期"]:
    check[col] = check[col].apply(
        lambda x: data_test.loc[x, "date"]
        if pd.notna(x)
        else None
    )

# =========================================
# 7. 输出结果
# =========================================

print("\n================ 数据维度 ================")
print(data_test.shape)

print("\n================ 数据检查 ================")
print(check)

print("\n================ 前5行 ================")
print(data_test.head())

# =========================================
# 8. 保存 Excel
# =========================================

with pd.ExcelWriter(OUT_DIR / "PPI_第一批数据测试.xlsx") as writer:

    data_test.to_excel(
        writer,
        sheet_name="数据",
        index=False
    )

    check.to_excel(
        writer,
        sheet_name="数据检查",
        index=False
    )

print("\n完成！")
print("文件：", OUT_DIR / "PPI_第一批数据测试.xlsx")
