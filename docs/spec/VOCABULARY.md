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

**按用途分,不按系统名分。** 工作流写 `page: list`,profile 把 `list` 绑定到
真实的 URL / 页面指纹 / 解析规则。

| role | 定义 | 判断方法 | 例子 |
|------|------|----------|------|
| `entry` | 起点页,通常需要人工打开或登录 | 「这页本身没有信息,只是入口」 | 各系统首页 |
| `query` | 输入 key、提交查询 | 「这页有输入框,我要填东西」 | 检索画面、查询表单 |
| `record` | 单条记录的详情,判定的主要依据 | 「这页只讲一件事」 | 处理结果画面 |
| `list` | 多行表格,需要在里面**定位到目标行** | 「这页有很多行,我要找出我那一行」 | 転送状態一覧、バッチ処理一覧、文件列表、作业列表 |
| `artifact` | 可下载产物的入口 | 「这页上有个链接,点了会下载东西」 | 下载页、日志下载 |
| `document` | 排版好的文档预览 | 「这页是给人看的成品,不是数据表」 | 帳票 preview |

### 2.1 为什么不设 `monitor` / `status` 这类角色

「転送状態一覧」本质上就是一个**要在里面找到目标行**的表格,和「Jenkins 文件
列表」「作业一览」是同一种东西,应该用同一套 step 处理(`verify.parse_text` →
`verify.match_record` → `verify.assert`)。

按系统叫法分角色,会让三个一模一样的东西变成三套代码 —— 这正是旧工具犯的错。

### 2.2 一个系统可以有多个角色页

```
系统 A: entry → query → record
系统 B: entry → list → artifact
系统 C: list(バッチ処理一覧) → document(帳票preview)
```

角色是**这一页在流程里干什么**,不是**它属于谁**。

### 2.3 拿不准怎么办

问一句:**「我在这页上要做的最主要的动作是什么?」**

- 填东西 → `query`
- 找我那一行 → `list`
- 读一件事的结论 → `record`
- 点下载 → `artifact`
- 看排版 → `document`
- 什么都不做,只是路过 → `entry`

如果一页同时是 `query` 和 `list`(填了就在同一页出结果),按**你要从它身上拿
什么**来定:要拿结果行 → `list`。

---

## 3. 动作(工作流命名)

| 中性动词 | 含义 | 旧名对照 |
|---------|------|----------|
| `capture` | 抓证据(截图 + 页面文本 + 判定) | `*Snap` 系列 |
| `collect` | 下载文件并归档 | `GfixLogDownload` / Jenkins 下载 |
| `compose` | 把证据组装进交付物工作簿 | `Replace*` |
| `annotate` | 在工作簿上画框 / 标注 | `Mark*` |
| `review` | 人工复核 + 记录决定 | `Review*` |
| `deliver` | 交付(文件 / 邮件 / 检查表) | `Deliver*` / `CheckSheet` |
| `sync` | 与基线对比 / 同步 | `Align` |
| `derive` | 从外部表生成工作清单 | `Generate-HostOpenMapping` |

### 3.1 工作流 id 的组成

```
<side>.<role>.<action>
```

| 新 id | 旧 phase |
|-------|----------|
| `before.record.capture` | `GiftHmSnap` |
| `before.list.capture` | `GiftMqSnap` |
| `after.record.capture` | `GfixHmSnap` |
| `after.artifact.collect` | `GfixLogDownload` |
| `before.compose` | `ReplaceGift` |
| `before.annotate` | `MarkGift` |

只有一面时省略 `side`:`list.capture`。

### 3.2 目录命名

```
worklist.csv                          工作清单
capture/<side>_<role>/<key>.png       截图
capture/<side>_<role>/<key>.txt       页面文本(与截图同时归档)
run/<runId>/trace.jsonl               本次运行的完整记录
run/<runId>/ledger.jsonl              断点续跑用的完成台账
.ebi/                                 本机状态(不进 git)
```

旧的 `snap/GIFT_MQ/<id>.png` → 新的 `capture/before_list/<key>.png`。

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
| `profiles/<新名>/pages.json`(每个 role 绑哪个 URL / 页面指纹 / Tab 序列) | `kernel/**` |
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
| `HM` 画面 | role `record` |
| `MQ` 転送状態 | role `list` |
| `Jenkins` 文件列表 | role `list` |
| `GoAnywhere` 作业一览 | role `list` |
| Jenkins 下载 / GFIX 日志下载 | role `artifact` |
| `snap/` | `capture/` |
| phase | workflow |
| `isReplaced` / `isMarked` / `isReviewed` | profile 声明的 checkpoint 位 |
| `SnapVerify` 判定 | `verify.assert` + profile 规则表 |
| `progress.jsonl` | `run/<runId>/trace.jsonl`(细化) |
