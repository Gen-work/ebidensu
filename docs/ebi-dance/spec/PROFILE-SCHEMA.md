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
  vocabulary.json     side 名(显示名)、列名映射
  pages.json          命名页面实例(page):role、显示名、URL / 指纹 / 导航 / 超时
  grammar.json        页面文本 → records 的解析规则(按 page 名)
  rules.json          判定规则表(按 page 名)
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
  "columns": {
    "key":         "Correl_ID_S",
    "group":       "JOB_NAME",
    "owner":       "Owner",
    "deliverable": "Excel_NAME"
  }
}
```

`sides` 的值只用于**显示**。工作流里永远只出现中性名。

> **v3 修订(P0-R1)**:初版还有一个 `roles` 显示名映射,已删 —— 页面的显示名
> 归 pages.json 每个 **page** 条目自己的 `title`(`ebi explain` 写成
> `transferStatus(MQ転送状態一覧)`);role 是结构枚举,不需要每个项目起显示名。

`columns` 把中性名映射到清单 CSV 的实际列名。**`{{item.key}}` 就是靠这个解析的。**

---

## 3. `pages.json`

每个 **page(命名页面实例)** 一个条目,**键是 page 名**。

```jsonc
{
  "transferStatus": {
    "role":     "list",
    "title":    "MQ転送状態一覧",

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
    "crop":        { "left": 6, "top": 6, "right": 6, "bottom": 6 }
  },

  "fileList": {
    "role":  "list",
    "title": "受信ファイル一覧",
    "url":   "https://<host2>/files/",
    "fingerprint": { "ok": ["ファイル名", "更新日時"] }
  }
}
```

- `role` 决定用哪套定位 / 解析机制(VOCABULARY §2.1 的 5 个枚举);
  `title` 是显示名,系统名只出现在这里。
- **同一个 role 可以有任意多个 page**(上例两个都是 `list`)—— 这正是引入
  page 的原因:role 当键会让同侧的第二个同型页面无处安放,而当前工作
  before 侧就同时有転送状態和ファイル一覧。
- `crop` 这类**页面级参数**也放在 page 条目里,随 page 一起换,
  工作流经模板引用(见 WORKFLOW-SCHEMA §4)。

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

## 4. `grammar.json`

页面文本(`Ctrl+A` 得到的整页纯文本)→ 结构化 records。
**键是 page 名**,与 pages.json 一一对应。

支持三种解析器,够覆盖旧仓库里全部三套手写解析:

```jsonc
{
  "transferStatus": {
    "parser": "delimited",
    "delimiter": "\t",
    "rowWhen": { "field": 0, "matches": "^\\d+$" },
    "fields": ["jobNo", "key", "status", "recvTime", "count", "rtncd"]
  },

  "procResult": {
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

## 5. `rules.json`

判定规则表。**键是 page 名。**`verify.assert` 按顺序跑,**第一条不满足的决定结论**。

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
    { "name": "before_procResult",     "role": "verdict", "default": "" },
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
  transferStatus-ok.txt          正常页面的 Ctrl+A 文本(已脱敏)
  transferStatus-ng-rtncd.txt    返回码非 0
  transferStatus-empty.txt       查询结果为空
  transferStatus-expired.txt     会话超时
  transferStatus-multi-row.txt   同一 key 多行
```

文件名前缀是 **page 名**。每个 fixture 配一个期望结论,组成单测:

```jsonc
// fixtures/expected.json
{
  "transferStatus-ok.txt":        { "verdict": "ok" },
  "transferStatus-ng-rtncd.txt":  { "verdict": "ng", "message": "リターンコードが0以外" },
  "transferStatus-empty.txt":     { "verdict": "unknown" },
  "transferStatus-multi-row.txt": { "verdict": "ok", "matchedRow": 3 }
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
