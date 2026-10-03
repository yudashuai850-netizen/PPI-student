from pathlib import Path

import pandas as pd

# 统一输出目录：项目根目录下的 data/output/
OUT_DIR = Path(__file__).resolve().parents[1] / "data" / "output"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# =========================
# 1. 设置需要下载的数据
# =========================

series = {
    "PPI": "CHNPIEATI01GYM",
}

# =========================
# 2. 下载数据
# =========================

data_all = pd.DataFrame()

for name, code in series.items():

    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={code}"

    print(f"正在下载：{name}")

    data = pd.read_csv(url)

    data = data.rename(columns={
        data.columns[0]: "date",
        data.columns[1]: name
    })

    data["date"] = pd.to_datetime(data["date"])

    if data_all.empty:
        data_all = data
    else:
        data_all = pd.merge(
            data_all,
            data,
            on="date",
            how="outer"
        )

# =========================
# 3. 排序
# =========================

data_all = data_all.sort_values("date")

# =========================
# 4. 输出基本信息
# =========================

print("\n数据基本情况：")
print(data_all.head())

print("\n数据维度：")
print(data_all.shape)

print("\n日期范围：")
print(data_all["date"].min())
print(data_all["date"].max())

print("\n缺失值：")
print(data_all.isna().sum())

# =========================
# 5. 保存
# =========================

data_all.to_excel(
    OUT_DIR / "PPI_数据测试.xlsx",
    index=False
)

print("\n已经生成：", OUT_DIR / "PPI_数据测试.xlsx")
