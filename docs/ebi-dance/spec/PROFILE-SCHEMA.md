# PROFILE-SCHEMA — 项目档案规格

> 状态:**草案**,P0 阶段定稿。
> profile = **一个项目的全部数据化知识**。换一份工作 = 写一个新 profile。

## 0. 定位

| 层 | 放什么 | 谁改 |
|----|--------|------|
| `kernel/` `modules/` | 通用能力 | 只有加新能力时改 |
| `workflows/*.json` | 流程连线 | 换工作时**基本能原样抄** |
| **`profiles/<name>/`** | **具体系统知识** | **换工作时全改这里** |
| `<WorkDir>/ebi.local.json` | 本机 / 本次作业的临时覆盖 | 操作员随手改 |

**优先级(高 → 低)**:CLI 参数 > 工作目录 overlay > profile > 工具默认值。

这套分层沿用旧仓库已验证的 `ConfigOverlay` 深合并机制,不重新发明。

---

## 1. 文件构成

```
profiles/<name>/
  vocabulary.json     side 名、role 显示名、列名映射
  pages.json          每个 page(命名页面实例)声明 role + URL / 指纹 / 导航 / 超时
  grammar.json        页面文本 → records 的解析规则
  rules.json          判定规则表
  worklist.json       清单列 schema、主键、变体规则、checkpoint 位
  layout.json         工作簿位置、画框坐标(有 compose/annotate 才需要)
  calibration.json    降级层的阈值和有效期(用到 fallback 才需要)
  fixtures/           脱敏后的页面文本样本,用于单测
  ocr-truth/          OCR 校准样本(用到 fallback 才需要)
```

全部 **UTF-8 无 BOM**。日文/中文可以直接写在 JSON 里 —— 这正是把它们从 `.ps1`
挪出来的原因(`.ps1` 里的非 ASCII 在 JP locale 主机上会乱码,JSON 不会)。

---

## 2. `vocabulary.json`

```jsonc
{
  "sides": {
    "before": "移行前",
    "after":  "移行後"
  },
  "roles": {
    "entry":    "入口",
    "form":     "検索画面",
    "record":   "結果画面",
    "list":     "一覧",
    "document": "帳票"
  },
  "columns": {
    "group":       "JOB_NAME",
    "owner":       "Owner",
    "deliverable": "Excel_NAME"
  }
}
```

`sides` / `roles` 的值只用于**显示**。`roles` 只有 5 个键(`VOCABULARY.md` §2.1
的 5 个 role,穷举,不能多也不能少),给的是**角色的通用显示名**——注意
`list` 这里只能填「一覧」这种泛称,**不能**填某一个具体页面的名字(比如
「転送状態一覧」)。原因见 §3:同一个 role 下**可以有多个 page**,每个
page 各有各的具体名字,那个名字属于 `pages.json` 的 `label` 字段,不属于
这里。`ebi explain` 的 `list(転送状態一覧)` 这种输出,`list` 来自这里、
`転送状態一覧` 来自对应 page 的 `label`,两者拼在一起显示,不是同一个字段。

工作流里永远只出现左边的中性名(`role` 的 5 个枚举值,或者具体的 `page` 名)。

`columns` 把中性名映射到清单 CSV 的实际列名 —— **不包含 `key`**。key 只
在 `worklist.json` 里声明一次(见 §6.1,P0-R4),这里重复声明会导致两处
漂移,已经删掉。

---

## 3. `pages.json`(P0-R1:按 page 名为键,不按 role 为键)

**每个 page(命名页面实例)一个条目,不是每个 role 一个条目。**

一个项目同一侧经常有**好几个结构相同的页面**——当前工作里 before 侧同时
有 MQ 転送状態一覧(role `list`)和 Jenkins 文件列表(role `list`),after
侧还有 GoAnywhere 作业一览(也是 role `list`)。三个页面结构相同、用同一套
定位/解析机制,但 URL、指纹、字段完全不同,是**三个不同的 page**。如果拿
role 当 `pages.json` 的键,三个 page 会互相覆盖(同一个 `"list"` 键只能存
一份数据)—— 这正是本卡要堵的洞。

每个 page 条目里的 `role` 属性**只决定用哪套定位/解析机制**(见
`VOCABULARY.md` §2.1 的 5 个 role),不代表这个 page 的身份。

```jsonc
{
  "transferStatus": {
    "role":       "list",
    "label":      "転送状態一覧",
    "url":        "https://<host>/path/index.html",
    "openHint":   "ブラウザで転送状態ページを開いてください",

    "fingerprint": {
      "ok":      ["転送状態", "Correlid"],
      "loading": ["読み込み中"],
      "empty":   ["No Data"],
      "expired": ["ログイン", "セッション"]
    },

    "tabsToForm":  1,
    "tabsToInput": 4,
    "timeoutSec":  12,
    "pollMs":      800,

    "crop": { "left": 6, "top": 6, "right": 6, "bottom": 6 }
  },

  "fileList": {
    "role":  "list",
    "label": "ファイル一覧",
    "url":   "https://<host>/jenkins/files.html",
    "fingerprint": {
      "ok":      ["ファイル名", "更新日時"],
      "expired": ["ログイン"]
    },
    "timeoutSec": 12
  }
}
```

`label` 是这个 page 的**具体显示名**(`ebi explain` 拼成
`list(転送状態一覧)` 时,`list` 来自 `vocabulary.json` 的 role 通用名,
`転送状態一覧` 来自这里的 `label`)。`crop` 是这个 page 专属的截图裁剪
边距(不同页面的窗口边框/内容区可能不一样,同role 的两个 page 未必能共用
一份裁剪参数)。

### 3.0 当前工作的 5 个 page(验证:role 相同也不冲突)

这是 P0-R1 的验收:把当前 Host→Open 迁移工作里全部会用到的页面按新规格
逐个列出 page 名 + role,确认没有 id 冲突、没有 capture 目录冲突。

| page 名 | role | 现在对应 | capture 目录 |
|---------|------|----------|---------------|
| `hmSearch` | `form` | HM 検索画面 | (不截图,见下方"两页流程") |
| `hmResult` | `record` | HM 処理結果画面 | `capture/<side>_hmResult/` |
| `transferStatus` | `list` | MQ 転送状態一覧(自带搜索框,见下方) | `capture/<side>_transferStatus/` |
| `fileList` | `list` | Jenkins 文件列表 | `capture/<side>_fileList/` |
| `jobList` | `list` | GoAnywhere 作业一览 | `capture/<side>_jobList/` |
| `reportPreview` | `document` | 帳票 preview | `capture/<side>_reportPreview/` |

6 个 page 名互不相同 → 5 个 capture 目录互不相同(`hmSearch` 是入口/表单
页,不出证据,不落 capture 目录),即使其中 3 个都是 role `list` 也不会
互相覆盖。这就是 R1 要修的洞:旧设计下 `transferStatus` / `fileList` /
`jobList` 会全部落到同一个 `capture/<side>_list/` 目录,互相覆盖对方的
截图。

**两页流程 vs 单页自带表单 —— 这条容易漏,一并验证清楚。** `VOCABULARY.md`
§6 说「検索画面 → role `form`」,但不是每个目标系统都有一个独立的检索
画面:

- **MQ(单页自带表单)**:`transferStatus` 这一个 page 上既有输入框
  (Tab 到字段、填 key、提交)又有结果表格 —— 提交后**同一个 URL**
  刷新出匹配的行。这种情况下 `pages.json` 的 `transferStatus` 条目自己
  带 `tabsToForm`/`tabsToInput`(`WORKFLOW-SCHEMA.md` §8 的示例就是这种,
  `page` 顶层只绑 `transferStatus` 一个 page,够用)。
- **HM(两页)**:先到 `hmSearch`(独立 URL,只有输入框)填 key、提交,
  浏览器跳转到**另一个 URL**才是 `hmResult`。这种情况下工作流仍然只
  `"page": "hmResult"` 绑定一个(要截图、要判定的是 `hmResult`),但
  `each` 段开头的导航/填表步骤引用 `hmSearch` 的数据时,写**完整路径**
  `{{profile.pages.hmSearch.url}}` / `{{profile.pages.hmSearch.tabsToInput}}`
  ——不是 `{{page.X}}`(那只指向已绑定的 `hmResult`)。**这不需要新机制**:
  `{{page.X}}` 是"当前绑定 page"的简写,不代表工作流只能引用绑定的那一个
  page;引用别的 page 的数据,原来的完整路径写法永远可用,只是没有简写。

判断走哪种的问题(访谈时问,见 `INTERVIEW.md` §4 的 2.2):填完搜索表单
提交后,地址栏 / 页面指纹变了吗?变了 → 两页;没变(只是同一页面刷新出
结果)→ 单页自带表单。

### 3.1 `fingerprint` — 页面指纹

**整套容错的地基。** 没有它,工具会在错误的页面上截图并判定成功。

| 键 | 含义 | 不匹配任何一个时 |
|----|------|------------------|
| `ok` | 正确页面独有的字符串,**全部**都要出现 | |
| `loading` | 加载中 | 继续轮询 |
| `empty` | 查询结果为空 | → `verdict: unknown` + gate |
| `expired` | 会话超时 / 跳回登录 | → gate,提示重新登录 |
| (以上都不匹配) | 未知页面 | → gate,**绝不截图** |

指纹的松紧由页面改版频率决定(见 `INTERVIEW.md` §7 的 5.11):
改版频繁 → 只放最核心的一两个词;多年不变 → 可以放严一点,更早发现异常。

---

## 4. `grammar.json`(P0-R1:按 page 名为键)

页面文本(`Ctrl+A` 得到的整页纯文本)→ 结构化 records。**键是 page 名,
和 `pages.json` 一一对应** —— 旧示例这里混用过 `list`(role)和
`fileList`(page 名)两种键,已经修正为清一色的 page 名。

支持三种解析器,够覆盖旧仓库里全部三套手写解析:

```jsonc
{
  "transferStatus": {
    "parser": "delimited",
    "delimiter": "\t",
    "rowWhen": { "field": 0, "matches": "^\\d+$" },
    "fields": ["jobNo", "key", "status", "recvTime", "count", "rtncd"]
  },

  "hmResult": {
    "parser": "labeled",
    "pairs": {
      "status":   { "after": "状態", "take": "line" },
      "endTime":  { "after": "終了時刻", "take": "token" }
    }
  },

  "fileList": {
    "parser": "columns",
    "headerLine": { "contains": ["ファイル名", "更新日時"] },
    "columns": { "name": [0, 60], "time": [60, 80] }
  }
}
```

| parser | 适用 | 旧实现 |
|--------|------|--------|
| `delimited` | tab / 空白分隔的表格,靠某列的正则识别数据行 | `ConvertFrom-GfixJobListText` |
| `labeled` | 「标签: 值」形式的详情页 | `ConvertFrom-HmPageText` |
| `columns` | 固定列宽的等宽表格 | `ConvertFrom-JenkinsListText` |
| `regex` | **逃生舱**:一条正则,命名捕获组即字段 | — |

### 4.1 不要指望这四种能猜准 —— 用 `ebi grammar tune` 调

四种解析器覆盖不了所有页面,而且**即使覆盖得了,参数(分隔符、列宽、行识别
正则)也没人能一次猜对**。所以配套一个交互式调试器,这比多加几种解析器有用
得多:

```
ebi grammar tune capture/before_transferStatus/ABC123.txt --profile host-open --page transferStatus
```

循环:

```
  ┌ 当前 grammar: delimited, delimiter="\t", rowWhen=field[0] matches ^\d+$
  │
  │  解析结果(前 5 行 / 共 12 行):
  │  ┌────────┬──────────┬──────────┬─────────────────────┬───────┐
  │  │ jobNo  │ key      │ status   │ recvTime            │ count │
  │  ├────────┼──────────┼──────────┼─────────────────────┼───────┤
  │  │ 100234 │ ABC123   │ 正常終了 │ 2026/08/24 9:50:03  │ 1,204 │
  │  │ 100235 │ ABC124   │ 正常終了 │ 2026/08/24 10:02:11 │ 0     │
  │  └────────┴──────────┴──────────┴─────────────────────┴───────┘
  │  ⚠ 3 行未被识别为数据行(按 u 查看)
  │
  └ d=改分隔符 r=改行识别正则 c=改列名 p=换解析器 u=看未识别行
    a=让 AI 提议  s=存进 profile + 存成 fixture  q=退出   > _
```

要点:

1. **立刻可见。** 改一个参数马上重新解析、重新渲染表格,不用跑整条工作流
2. **未识别的行要显式报出来。** 旧工具最恶劣的一个 bug 就是「静默丢行」——
   单位数小时的行被正则漏掉,页面上明明有,判定却说"文件不在列表里"。
   调试器必须把「这 3 行我没认出来」摆在脸上
3. **`a` 让 AI 参与是可选的。** 把这份文本 + 当前解析结果交给 AI,让它提议
   grammar。**AI 的提议同样要在这个循环里跑一遍给人看**,不能直接采信
4. **`s` 同时做两件事**:写进 profile,并把这份文本存成 fixture + 期望结果。
   于是**每次调 grammar 都自动积累一个回归测试**

### 4.2 时间格式的一个坑

解析时间时**必须**允许单位数小时。

> **真实事故**:某系统把上午的时间戳渲染成 `9:50:03`(无前导零),
> 解析器要求 `\d{2}:\d{2}:\d{2}`,于是所有 10 点前的行被静默丢弃,
> 判定报告「文件不在列表里」—— 而文件明明就在页面上。

grammar 里的时间格式一律用 `H:mm:ss` 这种单字符说明符(.NET 的 `ParseExact`
接受 1 或 2 位),不要 `HH`。

---

## 5. `rules.json`(P0-R1:按 page 名为键)

判定规则表。`verify.assert` 按顺序跑,**第一条不满足的决定结论**。**键是
page 名**,和 `pages.json` / `grammar.json` 一致 —— 不同的 `list` 页判定
规则通常完全不同(MQ 転送状態的规则和 Jenkins 文件列表的规则没有理由一样),
按 role 为键就会强迫它们共用一份规则表。

```jsonc
{
  "transferStatus": {
    "rules": [
      { "field": "status",   "op": "equals", "value": "正常終了", "else": "ng",
        "message": "状態が正常終了ではない" },
      { "field": "rtncd",    "op": "equals", "value": "0",        "else": "ng",
        "message": "リターンコードが0以外" },
      { "field": "recvTime", "op": "within", "value": "{{run.window}}", "else": "ng",
        "message": "受信時刻が実行時間帯の外" },
      { "field": "count",    "op": "present",                      "else": "unknown",
        "message": "件数が読み取れない" }
    ],
    "default": "ok"
  }
}
```

`{{run.window}}` 是合法引用:`run` 作用域正式包含 `window`(`runId` /
`startedAt` / `operator` / `workDir` / `window`,`WORKFLOW-SCHEMA.md`
§4.1),由 `human.input` 或 CLI `--window` 写入(P2-07 接线,P0-R6)。
`rules.json` 里的模板和 `pages.json` 一样,在传给 `verify.assert` 之前会
被**递归求值一次**(`WORKFLOW-SCHEMA.md` §4.3)。

### 5.1 `op` 一览(**穷举**)

| op | 含义 |
|----|------|
| `equals` / `notEquals` | 相等 / 不等 |
| `in` / `notIn` | 值在数组里 |
| `matches` | 正则 |
| `present` / `empty` | 有值 / 无值 |
| `within` | 时间落在窗口内 |
| `gt` / `lt` / `gte` / `lte` | 数值比较 |

需要更复杂的判断 → **不要扩展 op**,写一个新的 `verify.*` step。
规则表要保持"能念给人听"。

### 5.2 `else` 只能是 `ng` 或 `unknown`

- 明确知道这是错的 → `ng`
- **读不出来 / 拿不准** → `unknown`(触发人工关卡)

**`else` 永远不能是 `ok`。** 见 `INTERVIEW.md` 铁律三。

### 5.3 `message` 是给人看的

出现在 `human.gate` 面板和结束汇总里。用操作员的语言写,不是技术语言。

---

## 6. `worklist.json`

```jsonc
{
  "file": "worklist.csv",
  "encoding": "utf8-bom",

  "key": {
    "columns": ["Correl_ID_S", "JOB_NAME"],
    "confirmedRules": [
      { "kind": "suffix", "pattern": "\\.\\d{6}\\.\\d{8}$",
        "note": "転送バッチのタイムスタンプ付き表記",
        "confirmedAt": "2026-08-24", "confirmedBy": "operator" },
      { "kind": "fullwidth" },
      { "kind": "case-insensitive" }
    ],
    "ambiguityPolicy": "listAndAsk"
  },

  "columns": [
    { "name": "Correl_ID_S",  "role": "key" },
    { "name": "JOB_NAME",     "role": "key" },
    { "name": "Excel_NAME",   "role": "deliverable" },
    { "name": "before_transferStatus", "role": "verdict", "default": "" },
    { "name": "before_hmResult",       "role": "verdict", "default": "" },
    { "name": "composed",     "role": "bitmask",
      "bits": { "before": 1, "after": 2, "compare": 4 } },
    { "name": "note",         "role": "text" }
  ]
}
```

### 6.1 主键可以是复合的

`key.columns` 是**数组**,不是单列。

> **真实场景**:当前的工作里,`Correl_ID_S` 相同但 `JOB_NAME` 不同 ——
> 单靠 correl id 分不开两件事。

- 一列 → `["Correl_ID_S"]`
- 复合 → `["Correl_ID_S", "JOB_NAME"]`,按顺序拼接比较
- **不同工作流可以用不同的键**:工作流可在 `source` 里写
  `"keyColumns": ["Correl_ID_S"]` 覆盖 profile 的默认。
  用于「这一步只按 correl 去重就够了」这类情况

`ebi lint` 检查:所有 `role: key` 的列都出现在 `key.columns` 里,反之亦然。

### 6.2 变体规则是**长出来的**,不是一次写死的

这是这一节最重要的设计,来自一个明确的现场判断:**没人能预先声明全部变体规则。**

> **兔子洞长什么样**:这个系统的 key 带时间戳,那个系统不带;同名文件一大堆,
> 时间戳还对不上(创建时间 ≠ 接收时间);中间还混着完全没有时间戳的文件 ——
> 但那些文件可能在别的表或列表里有自己的时间记录。
>
> 试图一次写全规则,只会写出一套自己都不敢信的规则。

所以:**规则分两层。**

| 层 | 内容 | 谁写 |
|----|------|------|
| `confirmedRules` | 已经被人确认过的规则,自动应用 | 从下面的歧义处理里**自动追加** |
| (运行时) | 匹配不上 / 匹配到多个 → 走歧义流程 | 人当场判断 |

### 6.3 歧义流程:摊开全部候选 + 全部证据,让人判

`ambiguityPolicy: "listAndAsk"` 时,任何 key 匹配拿不准都触发:

```
  ⚠ 找不到唯一匹配:key = ABC123 / JOB_A

  已应用的规则:后缀时间戳、全角、大小写不敏感
  找到 4 个候选,没有一个能靠现有规则确定:

  ┌───┬──────────────────────────────┬──────────┬──────────┬──────────┬────────┐
  │ # │ 候选                         │ 来源     │ 创建时间 │ 接收时间 │ 大小   │
  ├───┼──────────────────────────────┼──────────┼──────────┼──────────┼────────┤
  │ 1 │ ABC123.260824.10515511.dat   │ 下载目录 │ 09:51:02 │ 09:50:03 │ 1.2 MB │
  │ 2 │ ABC123.260824.10515533.dat   │ 下载目录 │ 09:53:40 │ 09:53:12 │ 1.2 MB │
  │ 3 │ ABC123.dat                   │ 下载目录 │ 08:12:00 │ (无)     │ 1.2 MB │
  │ 4 │ ABC123_old.dat               │ 下载目录 │ 昨天     │ (无)     │ 0.9 MB │
  ├───┴──────────────────────────────┴──────────┴──────────┴──────────┴────────┤
  │ 我的建议:#2                                                                │
  │   理由:接收时间 09:53:12 落在本次运行窗口内且最新;#1 也在窗口内但更早;   │
  │         #3 无接收时间,创建时间早于本次运行;#4 是昨天的                   │
  │ ⚠ 不确定的地方:#1 和 #2 只差 3 分钟,如果本次是重跑,可能两个都是本次的   │
  ├────────────────────────────────────────────────────────────────────────────┤
  │ 1-4=选一个  v=看某个的详情  n=都不对(标 unknown)  q=中止                  │
  │ 选完之后:要不要把这次的判断存成规则?(y/n)                                │
  └────────────────────────────────────────────────────────────────────────────┘
```

四条设计要求,缺一不可:

1. **摊开全部候选,不许只显示"最佳"。** 旧工具的做法是静默选最新的,只打一行
   `[WARN] N 个候选,选了最新的` —— 人根本不会去看那行
2. **每个候选带上全部可用证据**:文件名、来源、创建时间、接收时间、大小、
   在哪张表的哪一行。**证据不够就说证据不够**,不要假装能判断
3. **给建议 + 给理由 + 给不确定点。** 「我建议 #2,因为…;但如果是重跑,
   #1 也可能对」—— 让人在 3 秒内能验证我的推理,而不是盲信
4. **选完之后问"要不要固化成规则"。** 人说 y → 追加一条 `confirmedRules`,
   下次同型的情况自动处理。**规则库随使用长大,而不是靠一开始想全**

### 6.4 规则统一走一个地方

所有涉及 key 比较的 step(`verify.match_record`、`file.find`、
`excel.find_anchor`、`file.newest`……)**必须**调 `table.key` 提供的规范化 +
候选排序函数,不许自己写比较。

> **真实事故**:旧代码在**七个不同的地方**各自用 `-eq` 精确比较,于是同一个 ID
> 在有些环节匹配得上、有些匹配不上:找不到记录、下错文件、解压出来的文件因为
> 把时间戳后缀当成扩展名而打不开。修了三个版本才修干净。

`ebi lint` 会检查:profile 声明了 key 规则,但工作流里有 step 绕过了它 → 报警。

### 6.5 `role` 一览

| role | 含义 |
|------|------|
| `key` | 主键的一部分(**至少一个**,可多个组成复合键) |
| `group` | 分组列 |
| `owner` | 归属人 |
| `deliverable` | 交付物工作簿名 |
| `verdict` | 三值判定列(`ok` / `ng` / `unknown` / 空) |
| `bitmask` | 位掩码列,`bits` 声明位名 → 位值 |
| `text` | 自由文本 |
| `time` | 时间列 |

启动时按此 schema 自动补齐缺失的列(沿用 `Ensure-MappingColumns` 的行为)。

### 6.6 `{{item.key}}` / `{{item.keySafe}}`、规则落盘、标准候选形状(P0-R4)

三个之前没定义、会互相放大的洞,在这里一次定死。

**(a) `{{item.key}}` 是显示形,`{{item.keySafe}}` 是文件名安全形,不是
同一个东西:**

| 模板 | 定义 | 用途 |
|------|------|------|
| `{{item.key}}` | `key.columns` 按声明顺序取值,单列键就是该列原值;复合键用固定分隔符 `" / "` 拼接(`ABC123 / JOB_A`) | 人看的地方:`human.gate` 面板、歧义列表、trace 里的标签 |
| `{{item.keySafe}}` | 每列值先做**全角转半角**规范化(和 `table.key` 规范化候选用同一套函数,不另造),替换 Windows 文件名非法字符(`\ / : * ? " < > \|` 及控制字符)为 `_`,多列用 `_` 拼接 | **文件名 / 目录名一律用它**,`{{item.key}}` 不许出现在路径模板里 |

**二者都不是系统交互的入参。** 需要往搜索框里填、往 Ctrl+F 里塞的是**某一
列的原始值**,直接用 `{{item.<列名>}}`(如 `{{item.Correl_ID_S}}`)—— 拼
出来的显示形字符串(`"ABC123 / JOB_A"`)打进目标系统的输入框大概率是错的。

**(b) `confirmedRules` 的落盘位置:** 运行时新学到一条规则(§6.3 的歧义
面板问完"要不要固化"、人选 y 之后),**先写 `<WorkDir>/ebi.local.json`**
(§0 已有的"本机/本次作业临时覆盖"层),不是直接改 `profiles/<name>/
worklist.json`(那份文件在 git 里,office PC 上的部署副本不一定能同步
回去)。面板同时提示:

```
已学到 1 条规则,记得回填 profile 并提交:
  confirmedRules += { kind: "suffix", pattern: "...", note: "..." }
```

`ebi profile check` 检测 `<WorkDir>/ebi.local.json` 里有 `worklist.key.
confirmedRules` 但对应 profile 的 `worklist.json` 里没有同款规则的情况,
报「N 条学到的规则还没回填」。

**(c) 候选列表的标准形状。** `table.key` / `file.find` / `file.newest` /
`verify.match_record` 遇到歧义时**返回同一个形状**,`human.choose` 只写
**一份**渲染逻辑:

```jsonc
{
  "candidates": [
    { "id": "c1", "candidate": "ABC123.260824.10515511.dat",
      "evidence": { "source": "下载目录", "createdAt": "09:51:02",
                     "receivedAt": "09:50:03", "size": "1.2 MB" } },
    { "id": "c2", "candidate": "ABC123.260824.10515533.dat",
      "evidence": { "source": "下载目录", "createdAt": "09:53:40",
                     "receivedAt": "09:53:12", "size": "1.2 MB" } }
  ],
  "suggestion": { "id": "c2", "reason": "接收时间最新且落在本次运行窗口内" },
  "doubts": "#1 和 #2 只差 3 分钟,如果本次是重跑,可能两个都是本次的"
}
```

`evidence` 是自由字段的 map —— 证据不够就少填几个字段,**不要**编一个
假值。`suggestion` / `doubts` 都是可选的(证据实在不够时可以不给建议,
只摊开候选)。§6.3 的歧义面板就是这个形状的 ASCII 渲染。

---

## 7. `layout.json`

只有 `compose` / `annotate` 类工作流需要。

```jsonc
{
  "sheets": {
    "before": "移行前結果",
    "after":  "移行後結果"
  },
  "anchor": { "column": "A", "matchesKey": true },
  "pictures": {
    "before_transferStatus": { "offsetX": 0, "offsetY": 20, "scale": 1.0 }
  },
  "boxes": {
    "before_transferStatus": [
      { "offsetX": 167.9, "offsetY": 176.9, "width": 528.8, "height": 63,
        "baseRow": 2, "rowHeight": 63.8 }
    ]
  }
}
```

坐标单位是 Excel 的 point。这些数字**只能在办公 PC 上用 `ebi probe` 量出来**,
无法在开发环境推导 —— 所以它们必须是 profile 数据,不能是代码常量。

---

## 8. `calibration.json`

只有用到 `tier: fallback` 的 step 才需要。

```jsonc
{
  "ocr": {
    "minAccuracy": 0.99,
    "validDays":   90,
    "truthDir":    "ocr-truth",
    "language":    "ja"
  }
}
```

见 `STEP-CONTRACT.md` §5。不满足 → runner 拒绝运行。

---

## 9. `fixtures/` —— 让规则能在家里测

```
fixtures/
  list-ok.txt          正常页面的 Ctrl+A 文本(已脱敏)
  list-ng-rtncd.txt    返回码非 0
  list-empty.txt       查询结果为空
  list-expired.txt     会话超时
  list-multi-row.txt   同一 key 多行
```

每个 fixture 配一个期望结论,组成单测:

```jsonc
// fixtures/expected.json
{
  "list-ok.txt":        { "verdict": "ok" },
  "list-ng-rtncd.txt":  { "verdict": "ng", "message": "リターンコードが0以外" },
  "list-empty.txt":     { "verdict": "unknown" },
  "list-multi-row.txt": { "verdict": "ok", "matchedRow": 3 }
}
```

**这是整套设计里性价比最高的一环**:grammar + rules 是纯数据,fixture 是纯文本,
所以**判定逻辑可以完全在没有办公环境的机器上验证**。旧仓库已经用这招保住了
`SnapVerify` / `GfixLog`,新设计把它变成 profile 的标准组成部分。

fixture 必须先过 `ebi mask` 脱敏才能提交。

---

## 10. 新建一个 profile

```
ebi profile new <name>      # 生成骨架 + 每个字段的说明注释
ebi profile check <name>    # schema 校验 + fixture 跑一遍
ebi profile diff <a> <b>    # 两个 profile 的差异(迁移时看要改什么)
```

推荐做法:**复制一个已有的 profile 改**,而不是从空白开始 ——
`ebi profile diff` 会告诉你还有哪些字段没改。

### 10.1 换工作时的顺序

1. `ebi profile new <新名>`(或复制旧的)
2. 按 `INTERVIEW.md` 做访谈,边问边填 `vocabulary` → `pages` → `grammar` → `rules`
3. 收集 fixture,`ebi profile check` 跑绿 —— **这一步在家里就能做完**
4. 改 workflow 的 `profile` 字段指向新 profile,`ebi lint` + `ebi explain`
5. 办公 PC 上按 `INTERVIEW.md` §9.3 的五级顺序跑通

**第 3 步跑绿之前不要去办公 PC。** 判定逻辑的错误在家里发现最便宜。
