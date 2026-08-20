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
  pages.json          每个 role 绑哪个页面:URL / 指纹 / 导航 / 超时
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
    "query":  "検索画面",
    "record": "処理結果画面",
    "list":   "転送状態一覧",
    "artifact": "ファイル一覧"
  },
  "columns": {
    "key":         "Correl_ID_S",
    "group":       "JOB_NAME",
    "owner":       "Owner",
    "deliverable": "Excel_NAME"
  }
}
```

`sides` / `roles` 的值只用于**显示**(`ebi explain` 会写成 `list(転送状態一覧)`)。
工作流里永远只出现左边的中性名。

`columns` 把中性名映射到清单 CSV 的实际列名。**`{{item.key}}` 就是靠这个解析的。**

---

## 3. `pages.json`

每个 role 一个条目。

```jsonc
{
  "list": {
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
    "pollMs":      800
  }
}
```

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

支持三种解析器,够覆盖旧仓库里全部三套手写解析:

```jsonc
{
  "list": {
    "parser": "delimited",
    "delimiter": "\t",
    "rowWhen": { "field": 0, "matches": "^\\d+$" },
    "fields": ["jobNo", "key", "status", "recvTime", "count", "rtncd"]
  },

  "record": {
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

### 4.1 时间格式的一个坑

解析时间时**必须**允许单位数小时。

> **真实事故**:某系统把上午的时间戳渲染成 `9:50:03`(无前导零),
> 解析器要求 `\d{2}:\d{2}:\d{2}`,于是所有 10 点前的行被静默丢弃,
> 判定报告「文件不在列表里」—— 而文件明明就在页面上。

grammar 里的时间格式一律用 `H:mm:ss` 这种单字符说明符(.NET 的 `ParseExact`
接受 1 或 2 位),不要 `HH`。

---

## 5. `rules.json`

判定规则表。`verify.assert` 按顺序跑,**第一条不满足的决定结论**。

```jsonc
{
  "list": {
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
    "column": "Correl_ID_S",
    "aliases": [
      { "kind": "suffix", "pattern": "\\.\\d{6}\\.\\d{8}$",
        "note": "転送バッチのタイムスタンプ付き表記" },
      { "kind": "fullwidth" },
      { "kind": "case-insensitive" }
    ]
  },

  "columns": [
    { "name": "Correl_ID_S",  "role": "key" },
    { "name": "JOB_NAME",     "role": "group" },
    { "name": "Excel_NAME",   "role": "deliverable" },
    { "name": "before_list",  "role": "verdict", "default": "" },
    { "name": "before_record","role": "verdict", "default": "" },
    { "name": "composed",     "role": "bitmask",
      "bits": { "before": 1, "after": 2, "compare": 4 } },
    { "name": "note",         "role": "text" }
  ]
}
```

### 6.1 `key.aliases` —— 时间戳变体等

这是旧仓库最痛的一个坑。

> **真实事故**:主键在某些场景带批次时间戳后缀(`ABC123` vs
> `ABC123.260815.10515511`)。旧代码在**七个不同的地方**各自用 `-eq` 精确
> 比较,于是同一个 ID 在有些环节匹配得上、有些匹配不上:找不到记录、
> 下错文件、解压出来的文件因为把后缀当成扩展名而打不开。修了三个版本。

新设计:**变体规则只在这里声明一次**,所有涉及 key 比较的 step
(`verify.match_record`、`file.find`、`excel.find_anchor`……)统一走
`table.key` 提供的规范化函数。

`ebi lint` 会检查:profile 声明了 aliases,但工作流里有 step 绕过了它 → 报警。

### 6.2 `role` 一览

| role | 含义 |
|------|------|
| `key` | 主键(**必须恰好一个**) |
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
    "before_list": { "offsetX": 0, "offsetY": 20, "scale": 1.0 }
  },
  "boxes": {
    "before_list": [
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
