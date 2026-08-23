# huashu-excel · 花叔数据分析大师

给 AI agent 的数据分析与 Excel 全流程 skill。
装进 Claude Code、豆包、CodeBuddy 或任何能跑 Python 的 agent，都能用。

**它要解决的问题只有一个：让 AI 算出来的数字经得起追问。**

---

## 为什么需要它

三届微软 Excel 世界冠军 Andrew Ngai 这样评价 AI 做数据分析：

> 如果你用错误的数据训练 AI，它会给你错误的结果——但它还会装出对这些错误结果非常自信的样子。

这不是假想。下面是一个实测。

一份普通的中国公司销售表：标题占了第一行、表头两级、地区列用合并单元格、
金额带千分位逗号、中间夹着「华东小计」、尾巴三行是「合计/占比/同比」。

用主流做法处理它——`pd.read_excel()`，发现表头位置不对就加 `header=`，
发现数字读不出来就清掉千分位——然后算 1 月总销售额：

```
清理后求和：34,893,234        真值：13,367,767
总额误差：+161.0%
均值误差：+117.5%
零报错  零 NaN  零警告
```

错的原因：`合计` 行被当成了一家门店，`华东小计` 又被当成一家，重复行算了两次。

**一个 +161% 的错误，交付时长得和正确答案一模一样。**

Panko (1998) 的研究说，86% 的电子表格含有错误。
欧洲电子表格风险兴趣组（EuSpRIG）从 1995 年起持续收录见诸媒体的事故，
其中最著名的是 Reinhart-Rogoff 论文——Excel 选区少选 5 行，
把 +2.2% 的 GDP 增速算成 −0.1%，而那篇论文当时是全球紧缩政策的主要依据。

---

## 六步标准作业流程

```
1  体检   这张表长什么样 —— 结构、类型、脏点。先看再算
2  清洗   变成规范分析表，每一步可追溯可回放
3  口径   指标定义、时间窗口、分母、单位、去不去重 —— 写下来才算数
4  分析   扫陷阱，再按问题类型走配方
5  对账   行数守恒、总和守恒、与表内合计交叉验证
6  交付   Excel / 图表 / 报告，口径随数字一起交付
```

| 脚本 | 干什么 | 什么时候 |
|---|---|---|
| `profile_table.py` | 表结构体检 | 算任何数字之前 |
| `clean_table.py` | 清洗 + 生成可审计的 pandas 脚本 | 体检之后 |
| `scan_traps.py` | 分析陷阱扫描 | 下结论之前 |
| `verify_numbers.py` | 数字对账 | 交付之前 |
| `make_chart.py` | 图表推荐与生成 | 交付时 |

---

## 四个和别的工具不一样的地方

**一、先看原始单元格，再动 pandas。**
`pd.read_excel()` 读进来的那一刻，合并单元格、单元格格式、原始类型就全丢了。
现有工具都是在信息已经损毁之后才开始判断。体检脚本先用 openpyxl 读原始格子。

**二、把表里的「合计」行当成免费的校验和。**
所有工具都把汇总行当噪音过滤掉。但那是原表作者用公式算出来的真值。
拿清洗后的明细自己求和去对它——对不上就说明有一方错了，两种情况都必须报出来。
这对应 ICAEW《Financial Modelling Code》的 `Include a master check`。

**三、默认给五数概括，不给均值。**
业务数据几乎总是右偏的（少数大客户、少数爆款）。
`df.describe()` 把 mean/std 放在最前面，而均值在偏态分布下描述的不是任何真实对象。
默认输出 min/Q1/中位数/Q3/max + IQR，这是 Tukey 的抗差统计。

**四、机器判事实，人判品味。**
「行数对不对、总和守不守恒、这个数能不能追回源单元格」——机器验死。
「这个分析有没有意义」——明确交还给人，绝不用一个分数冒充客观。

---

## 图表

图表部分继承 Gene Zelazny《用图表说话》的核心律条——**先确定要传达的信息，
再选图表形式**——并补上那本书里没有的两件事：

**用实证决定视觉编码。** Cleveland & McGill (1984) 的图形感知实验测出了
人读取不同视觉编码的精度阶梯：位置(共同基线) > 位置(非对齐) > 长度/方向/角度 >
面积 > 体积/曲率 > 明暗/色饱和。位置判断比长度准 1.4–2.5 倍，比角度准 1.96 倍。

饼图靠角度和面积编码，正好落在下游。所以 `make_chart.py`
**在类别超过 3 个时会拒绝生成饼图**，并给出理由和替代方案。
这不是审美偏好，是有测量结果的。

**检查这张图有没有在误导人。** 生成时自动执行：柱形图数值轴强制从 0
（它编码长度）、折线图不强制从 0（它编码位置和斜率）、类别按值排序、
单系列去图例、套用色盲友好色板。完整清单见 `references/charts.md`。

---

## 安装

skill 目录放到 agent 的 skills 路径下即可：

```bash
# Claude Code / 通用
git clone https://github.com/<your-account>/huashu-excel ~/.claude/skills/huashu-excel

# 或放进项目级 skills 目录
git clone https://github.com/<your-account>/huashu-excel ./.claude/skills/huashu-excel
```

脚本也可以脱离 agent 单独用：

```bash
python3 scripts/profile_table.py 你的表.xlsx
python3 scripts/verify_numbers.py 你的表.xlsx     # 退出码 1 = 有对不上的
```

**依赖**：只需要 `openpyxl`（读写 .xlsx 时）。CSV 路径连它都不用，纯标准库。
不用 pandas、不用 LibreOffice、不联网、不依赖任何 agent 平台特性
（不需要 subagent、不需要沙盒）。

skill 会在开工时探测自己所处环境的能力，选择该环境下的最佳工作流——
能起并行子任务就并行分析，只能串行就串行轮转视角，跑不了脚本就转成
「给你 Excel 公式和操作步骤」。

---

## 方法论出处

这份 skill 的判断力不是凭空来的，每条都有出处：

| 来源 | 用在哪 |
|---|---|
| Hadley Wickham, *Tidy Data* | 清洗的目标形态与五类脏数据分类 |
| John Tukey, *Exploratory Data Analysis* (1977) | 五数概括、抗差统计、箱线图 |
| ICAEW, *Financial Modelling Code* (2024) | master check、可追溯引用、Excel 工程约定 |
| Cleveland & McGill (1984) | 视觉编码的感知精度阶梯 |
| Gene Zelazny, *Say It With Charts* | 先定信息再选图、比较类型 → 图表形式 |
| Barbara Minto, *The Pyramid Principle* | 结论先行、MECE、SCQA |
| Ronny Kohavi et al., *Trustworthy Online Controlled Experiments* | Twyman's Law、常见陷阱 |
| Cassie Kozyrkov | 探索与推断的分工红线 |
| Panko (1998)、EuSpRIG | 电子表格错误的实证规模 |

引用的规范文本均为提炼转述并注明出处，不含受版权保护的原文复制。

---

## License

MIT

---

## English

**huashu-excel** is a data-analysis and Excel skill for AI agents. It works in
Claude Code, Doubao, CodeBuddy, or any agent that can run Python.

It exists to solve one problem: **making the numbers an AI produces survive scrutiny.**

Measured on a typical messy spreadsheet (title in row 1, two-level header,
merged cells, thousands separators, subtotal rows mixed into the detail),
the mainstream approach overstated a simple sum by **161%** — with no error,
no NaN, and no warning.

Six-step workflow: profile → clean → declare caliber → analyze → reconcile → deliver.

Four things it does differently: it reads raw cells before touching pandas;
it treats the spreadsheet's own total rows as a free checksum; it reports
five-number summaries instead of mean/std by default (business data is skewed);
and it draws a hard line between what a machine can verify (arithmetic) and
what only a human can judge (whether the analysis means anything).

Chart selection follows Zelazny's "message first, chart form second," corrected
by the Cleveland–McGill perception hierarchy — which is why it will refuse to
generate a pie chart beyond three categories.

Dependencies: `openpyxl` only, and not even that for CSV. No pandas, no
LibreOffice, no network, no platform-specific agent features.
