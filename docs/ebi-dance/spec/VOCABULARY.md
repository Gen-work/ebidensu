# VOCABULARY — ebi-dance 中性词汇表

> 状态:**草案**,P0 阶段定稿。
> 所有标识符、字段名、step 名一律英文;说明文字用中文。

## 0. 这份文档要解决的问题

旧工具把具体系统名写死在代码和文件名里(`GiftHmSnap` / `GIFT_MQ` /
`Correl_ID_S`)。换一份工作 —— 页面不同、下载渠道不同、Excel 格式不同 ——
就要全局改名,等于重写。

**规则:工作流(workflow JSON)和模块(step)里只允许出现本文左列的中性名。
具体系统名只允许出现在 profile 数据里。**

判断标准很简单:*把这份工作流原样拿给另一个项目,除了 profile 之外还需要改
东西吗?* 需要,就说明有名字漏进了不该在的地方。

---

## 1. 核心名词

| 中性名 | 含义 | 旧名对照 |
|--------|------|----------|
| `worklist` | 工作清单 CSV,一行 = 一件要做的事 | `mapping_<Owner>.csv` |
| `item` | 清单里的一行 | 一个 correl 行 |
| `key` | item 的主键。profile 声明它是哪一列 | `Correl_ID_S` |
| `group` | item 的分组属性,用于合并页面访问 | `JOB_NAME` / `TO_code` |
| `owner` | 清单的归属人(用于分工筛选) | `Owner` |
| `deliverable` | 交付物工作簿 | `Excel_NAME` 指向的证据簿 |
| `side` | 对照面。profile 声明有哪几面及显示名 | `GIFT` / `GFIX` |
| `page` | **命名页面实例**:一个具体画面。pages.json 按 page 名声明,每个 page 有一个 role(结构)和显示名 | HM 処理結果 / MQ転送状態 / Jenkins ファイル一覧(各是一个 page) |
| `capture` | 一次证据抓取的产物(PNG + 页面文本) | `snap\<folder>\<id>.png` / `.txt` |
| `verdict` | 一次判定的结论 | `GIFT_MQ_snap` 的 0/1/2 |
| `profile` | 一个项目的全部数据化知识 | `verify_config.json` + 散落的硬编码 |
| `workflow` | 一条声明式流程 | 一个 phase |
| `step` | 一个正交能力原语 | 散落的函数 |
| `run` | 一次工作流执行 | — |
| `trace` | 一次 run 的完整机器可读记录 | `progress.jsonl`(粗粒度版) |

### 1.1 `side` 的用法

`side` 是「同一件事的几个对照面」。迁移类工作通常是两面:

```json
{ "sides": { "before": "移行前", "after": "移行後" } }
```

工作流里写 `{{vars.side}}`,profile 决定它显示成什么。
**不要**在工作流里写 `gift` / `gfix`。

非迁移类工作可以只有一面,或者三面以上 —— `sides` 是个字典,不限定数量。

### 1.2 `verdict` 的取值(固定三值,不许扩展)

| 值 | 含义 | 后续处理 |
|----|------|----------|
| `ok` | 明确通过 | 标记完成 |
| `ng` | 明确不通过 | 标记为 NG,**下次运行仍会被选中重做**,并列入结束时的汇总 |
| `unknown` | **判定不出来** | 触发人工关卡。**绝不允许自动降级成 `ok`** |

`unknown` 是这套设计里最重要的一个值。旧工具的教训是:判定不确定时偷偷挑一个
"看起来对的",结果把正确的数据改错了。宁可停下来问人。

---

## 2. 页面角色(page role)

> **v2 修订。** 初版把 role 定义成「你要从这页拿什么」,这是错的:
> 一页上经常要取好几份证据 —— 既要截图存档,又要点某行的链接下载,
> 文档页也可能既要截图又要下载。用途不唯一,所以不能拿用途当分类轴。
>
> **修正:role 描述页面的「形状」,即"怎么在这页上找到东西";
> 在一页上做几件事是 action 的事,可以有任意多个。**

> **v3 修订(开工前评审,P0-R1)**:role 不能当唯一键。一个项目**同一侧可以有
> 多个同型页面** —— 当前工作 before 侧就同时有 MQ転送状態和 Jenkins ファイル
> 一覧两个 `list` 页。所以引入 **page(命名页面实例)**:pages.json /
> grammar.json / rules.json / 工作流 id / capture 目录一律按 **page 名**索引;
> role 只是 page 的一个属性,决定「用哪套定位 / 解析机制」。

### 2.1 role = 页面的结构

| role | 结构特征 | 定位方式 | 例子 |
|------|----------|----------|------|
| `entry` | 没有目标数据,只是入口 / 导航 | 不需要定位 | 各系统首页、登录后首屏 |
| `form` | 有输入框,要填要提交 | 焦点序列(Tab / 点击) | 検索画面、查询表单 |
| `record` | 单条记录,「标签: 值」结构 | 按标签取值 | 処理結果画面 |
| `list` | 多行表格 | **解析全表 → 定位目标行** | 転送状態一覧、バッチ処理一覧、文件列表、作业列表 |
| `document` | 排版好的成品,不是数据结构 | 整页 / 按区域 | 帳票 preview、PDF 预览 |

**5 个,不是 6 个。** 初版的 `artifact`(可下载产物入口)被删掉了 ——
下载链接是**长在某一页上的一个元素**,不是一种页面形状。一个 `list` 页的每行
可能都有下载链接;一个 `document` 页可能也能下载。所以下载是 action。

### 2.2 action = 在这一页上做什么(可以有多个)

| action | 含义 | 依赖的 role 能力 |
|--------|------|------------------|
| `read` | 取页面文本并解析 | 任意 |
| `locate` | 定位到目标行 / 目标字段 | `list` / `record` |
| `capture` | 截图取证(**同一页可多次**,不同区域/滚动位置) | 任意 |
| `download` | 点链接 / 按钮下载文件 | 任意 |
| `input` | 填写并提交 | `form` |
| `navigate` | 跳转到下一页 | 任意 |

一页上典型的组合:

```
list 页:  read → locate(找到我那一行) → capture(截图) → download(点那行的链接)
                                       └→ capture(再截一张别的区域)
document 页: capture(整页截图) → download(下载 PDF)
form 页:  input → navigate
```

### 2.3 为什么不设 `monitor` / `status` 这类 role

「転送状態一覧」的结构就是一张多行表格,和「文件列表」「作业一览」**完全同型**,
用同一套定位机制(`verify.parse_text` → `verify.match_record`)。

按系统叫法分 role,会让三个同型的东西变成三套代码 —— 这正是旧工具犯的错。

但**同型 ≠ 同一**:它们是三个不同的 **page**(`transferStatus` / `fileList` /
`jobList`),共享 role `list` 的机制,各有自己的 pages.json 条目、grammar 和
rules。role 管机制,page 管身份。

### 2.4 拿不准怎么办

问的**不是**「我要拿什么」(那是 action),而是:

> **「这一页上的东西是怎么排列的?我要靠什么找到目标?」**

- 靠 Tab 找到输入框 → `form`
- 靠标签找到值 → `record`
- 靠在表里找我那一行 → `list`
- 它就是一张排好版的图,没有可定位的结构 → `document`
- 上面根本没有我要的东西 → `entry`

### 2.5 一页多份证据的命名

同一页多次 `capture` 时,产物加 tag 区分:

```
capture/<side>_<page>/<keySafe>.png            默认(单张)
capture/<side>_<page>/<keySafe>__<tag>.png     多张时,tag 由工作流指定
```

文件名一律用 `keySafe`(复合键 + 非法字符的文件名安全形,定义见
PROFILE-SCHEMA §6.1b),不用裸 key。

例:`capture/before_transferStatus/ABC123__row.png`、
`capture/before_transferStatus/ABC123__total.png`

目录按 **page** 分,不按 role 分 —— before 侧的転送状態截图和ファイル一覧
截图是两个目录(`before_transferStatus` / `before_fileList`),按 role 分
它们会挤进同一个 `before_list` 互相覆盖。

---

## 3. 工作流动词(workflow verb)

**注意和 §2.2 的 action 区分**:action 是「在一页上做的一个动作」(粒度小,
一页可有多个);verb 是「一条工作流整体在干什么」(粒度大,给工作流起名用)。

| verb | 含义 | 旧名对照 |
|------|------|----------|
| `capture` | 抓证据(截图 + 页面文本 + 判定) | `*Snap` 系列 |
| `collect` | 下载文件并归档 | `GfixLogDownload` / Jenkins 下载 |
| `compose` | 把证据组装进交付物工作簿 | `Replace*` |
| `annotate` | 在工作簿上画框 / 标注 | `Mark*` |
| `review` | 人工复核 + 记录决定 | `Review*` |
| `deliver` | 交付(文件 / 邮件 / 检查表) | `Deliver*` / `CheckSheet` |
| `sync` | 与基线对比 / 同步 | `Align` |
| `derive` | 从外部表生成工作清单 | `Generate-HostOpenMapping` |

### 3.1 这张表是开放的

这 8 个覆盖目前见过的全部场景,但**不是封闭枚举**。加一个新 verb 的成本很低:

- verb 只用于**工作流命名**,没有代码依赖它
- 加的时候在本表补一行,说清楚它和已有 verb 的区别
- 只有一条判断标准:**它是不是真的和已有 8 个都不同?**
  如果只是「同一件事在另一个系统上做」,那不是新 verb,是新 profile

### 3.2 工作流 id 的组成

```
<side>.<page>.<verb>       页面相关的
<side>.<verb>              工作簿 / 文件相关的(没有特定页面)
```

用 **page 名**,不用 role —— 否则 `GiftMqSnap` 和 `GiftJenkins` 都叫
`before.list.capture`,直接撞名。page 名要中性(说结构 / 职能,不说系统名):
`transferStatus` 而不是 `mq`,`fileList` 而不是 `jenkins` —— 系统名只出现在
pages.json 的 `title`(显示名)里。

| 新 id | 旧 phase |
|-------|----------|
| `before.procResult.capture` | `GiftHmSnap` |
| `before.transferStatus.capture` | `GiftMqSnap` |
| `before.fileList.capture` | `GiftJenkins` |
| `after.procResult.capture` | `GfixHmSnap` |
| `after.jobList.collect` | `GfixLogDownload`(在 jobList 页上点下载) |
| `before.compose` | `ReplaceGift` |
| `before.annotate` | `MarkGift` |

只有一面时省略 `side`:`transferStatus.capture`。

一条工作流可以同时做几件事(在 list 页上既截图又下载),这时按**主要目的**
命名,或者拆成两条工作流 —— 拆开的好处是可以分别重跑。

### 3.2 目录命名

```
worklist.csv                              工作清单
capture/<side>_<page>/<keySafe>.png       截图
capture/<side>_<page>/<keySafe>.txt       页面文本(与截图同时归档)
run/<runId>/trace.jsonl               本次运行的完整记录
run/<runId>/ledger.jsonl              断点续跑用的完成台账
.ebi/                                 本机状态(不进 git)
```

旧的 `snap/GIFT_MQ/<id>.png` → 新的 `capture/before_transferStatus/<key>.png`。

---

## 4. 副作用等级(step 的 `effects` 字段)

用来让 `ebi explain` 一眼看出一条工作流有多危险。

| 等级 | 含义 | 例子 |
|------|------|------|
| `pure` | 纯函数,无 I/O | `verify.parse_text` |
| `read` | 只读外部状态 | `table.load`、`file.stat` |
| `ui` | 操作前台窗口 / 键鼠 | `browser.fill`、`browser.submit` |
| `write` | 写文件 / 写清单 | `screen.save`、`table.set` |
| `destructive` | **覆盖或删除已有数据** | `excel.replace_sheet`、`file.cleanup` |

`destructive` 的 step,runner **默认**在执行前插入一个 `human.gate`,
除非工作流显式声明 `confirm: false`。

---

## 5. 换一份工作时要改什么

这是本文的验收标准。换工作时的完整清单:

| 要改 | 不用改 |
|------|--------|
| `profiles/<新名>/vocabulary.json`(side 名、role 显示名、列名映射) | `modules/**` 全部 step |
| `profiles/<新名>/pages.json`(每个 page 的 role / URL / 页面指纹 / Tab 序列) | `kernel/**` |
| `profiles/<新名>/rules.json`(判定规则表) | `docs/**` |
| `profiles/<新名>/worklist.json`(清单列 schema、主键、位定义) | 多数 `workflows/*.json`(能直接抄) |
| `profiles/<新名>/layout.json`(工作簿位置、画框坐标) | |

**如果换工作时需要动 `modules/` 或 `kernel/`,说明抽象漏了。**
把漏掉的东西补成新 step 记入 catalog —— 这是正常的成长,但每次都要问一句:
*这个新 step 是通用能力,还是我把业务知识塞进代码了?*

---

## 6. 旧名 → 新名完整对照

迁移期用,迁完即弃。

| 旧 | 新 |
|----|----|
| `mapping_<Owner>.csv` | `worklist.csv` |
| `Correl_ID_S` | `key`(profile 声明实际列名) |
| `JOB_NAME` | `group` |
| `Excel_NAME` | `deliverable` |
| `GIFT` / `GFIX` | `before` / `after`(profile 声明) |
| `HM` 画面 | page `procResult`(role `record`) |
| `MQ` 転送状態 | page `transferStatus`(role `list`) |
| `Jenkins` 文件列表 | page `fileList`(role `list`) |
| `GoAnywhere` 作业一览 | page `jobList`(role `list`) |
| Jenkins 下载 / GFIX 日志下载 | 在 `fileList` / `jobList` 页上的 `download` action |
| 検索画面 | role `form` |
| `snap/` | `capture/` |
| phase | workflow |
| `isReplaced` / `isMarked` / `isReviewed` | profile 声明的 checkpoint 位 |
| `SnapVerify` 判定 | `verify.assert` + profile 规则表 |
| `progress.jsonl` | `run/<runId>/trace.jsonl`(细化) |
