# PythonProject

个人 Python 工作目录：宏观数据抓取（PPI 论文用）+ 一些零散练习脚本。

**整理日期：2026-10-03**

---

## 一、目录结构

```
PythonProject/
├── .venv/                     ← 虚拟环境（未做任何改动）
├── .idea/                     ← PyCharm 工程配置（未做任何改动）
├── .gitignore
├── requirements.txt           ← 虚拟环境依赖快照
├── README.md                  ← 本文件
│
├── data/                      ← 所有数据文件集中放这里
│   ├── LoanStats_with_group.xlsx
│   ├── PPI_第一批数据测试.xlsx
│   └── output/                ← 脚本跑出来的结果统一写到这里
│
├── ppi_paper/                 ← 论文主体（原 demo1 文件夹）
│   ├── 论文代码.py             ← PPI 高维宏观试提数据：多来源自动抓取 V2.0
│   └── PPI_试提_V2_输出/       ← 论文代码的输出目录（01_原始 / 02_处理后 / 03_日志）
│
└── scripts/                   ← 零散脚本、练习、测试代码
    ├── test.py
    ├── download_test.py
    ├── download_batch1.py
    ├── nbs_test.py
    ├── 宏观经济分析复现.py
    ├── 免费电影爬虫.py
    └── 计量经济学.py
```

## 二、每个文件做什么

| 文件 | 作用 | 输出 |
|---|---|---|
| `ppi_paper/论文代码.py` | 论文主程序。按《PPI试提清单与提取记录模板_V1.0》抓 1 个目标 + 33 个候选变量（AKShare / FRED / BIS / 国家统计局） | `ppi_paper/PPI_试提_V2_输出/` |
| `scripts/test.py` | FRED PPI 序列（CHNPIEATI01GYM）快速下载测试 | `data/output/china_ppi.csv` |
| `scripts/download_test.py` | FRED 下载流程测试（单序列） | `data/output/PPI_数据测试.xlsx` |
| `scripts/download_batch1.py` | 批量下载 PPI/CPI/汇率/进出口/铜价/铁矿石 + Brent 原油 | `data/output/PPI_第一批数据测试.xlsx` |
| `scripts/nbs_test.py` | 国家统计局 `easyquery.htm` 接口连通性测试（只打印，不存文件） | 无 |
| `scripts/宏观经济分析复现.py` | 把指数日度数据转成月度数据 | 只打印，需要 `data/指数数据.xlsx` |
| `scripts/免费电影爬虫.py` | IMDb 搜索页爬虫练习 | `data/output/movies.csv` |
| `scripts/计量经济学.py` | statsmodels 经典数据集（Advertising）练习 | 无 |

## 三、环境说明

- `.venv` 是 Python **3.12.10** 虚拟环境，基于 `C:\Users\30848\AppData\Local\Programs\Python\Python312` 创建。**本次整理完全没有动它。**
- `requirements.txt` 是当前环境的依赖快照，主要包含：`akshare 1.19.1`、`pandas 3.0.6`、`numpy 2.5.3`、`requests 2.34.2`、`beautifulsoup4 4.15.0`、`lxml 6.1.3`、`openpyxl 3.1.5`、`xlrd 2.0.2` 等。
- 如果需要在新电脑上重建环境：

  ```
  python -m venv .venv
  .venv\Scripts\activate
  pip install -r requirements.txt
  ```

## 四、怎么运行

在 PyCharm 里直接右键脚本 → Run 即可，工作目录用默认值就行。

之所以不用管工作目录，是因为整理时给几个会写文件的脚本加了 3 行代码，让它们**永远把结果写到项目根目录的 `data/output/`**，不管从哪儿启动都不会再把文件丢在项目根目录。改动的只有输出路径，业务逻辑一行没动。

## 五、本次整理做了什么

1. 新建 `data/`、`data/output/`、`scripts/`、`ppi_paper/` 四个文件夹。
2. `demo1/` → 改名为 `ppi_paper/`（`论文代码.py` 和它的输出目录一起搬，脚本内部用的是相对自身的位置，所以搬完照样能跑）。
3. 根目录散落的 2 个 Excel 移进 `data/`。
4. 根目录散落的 6 个脚本 + `demo1/计量经济学.py` 移进 `scripts/`。
5. 给 4 个会写文件的脚本补上统一的输出目录（见上面"怎么运行"）。
6. 新建本 README，把结构、用途、运行方式写清楚。

**没有删除任何文件，没有改动 `.venv`，没有改动 `.idea`。**

整理前的完整备份在：

```
C:\Users\30848\PycharmProjects\_backup_PythonProject_2026-10-03\
```

## 六、还没解决 / 建议

- `scripts/宏观经济分析复现.py` 需要的 `data/指数数据.xlsx` 目前在电脑上找不到，脚本会报"文件不存在"。找到后放进 `data/` 就能跑。
- `.idea/workspace.xml` 里有几条失效的运行配置（指向已经不存在的 `main.py`、`demo1/满屏飘.py`）。因为 PyCharm 当时是开着的，改它会被 IDE 覆盖，所以没动。**关掉 PyCharm 后可以让我清理**，或者自己在 Run/Debug Configurations 里右键删掉。
- `data/LoanStats_with_group.xlsx`（10MB，LendingClub 贷款数据）和 PPI 论文没关系，如果确定不用了可以移到别的地方，我只是先归到 `data/` 里。
