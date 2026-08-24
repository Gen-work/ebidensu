# WORKFLOW-SCHEMA — 工作流 JSON 规格

> 状态:**草案**,P0 阶段定稿。
> 读者:写工作流的人和 Agent。step 的实现契约见 `STEP-CONTRACT.md`。

## 0. 一条设计铁律

**workflow JSON 只做"连线",不做"计算"。**

具体禁止:

- 没有表达式(不能写 `{{item.count > 0}}`)
- 没有函数定义
- 没有循环变量之外的变量赋值
- 没有字符串拼接以外的运算

任何需要「判断」的地方,都必须落到一个 `verify.*` step 的 PowerShell 纯函数 +
profile 里的规则表。

理由:一旦 JSON 里能写表达式,它就开始长成一门自制编程语言 —— 难调试、难
diff、Agent 容易写出微妙的错,而且最终会比直接写 PowerShell 还难懂。

---

## 1. 顶层结构

```jsonc
{
  "id":      "before.list.capture",
  "title":   "転送状態ページの証跡取得",
  "version": "1.0.0",
  "profile": "host-open",

  "vars":    { "side": "before", "role": "list" },

  "source":  { ... },      // 遍历什么,见 §3
  "onError": { ... },      // 默认容错策略,见 §6

  "setup":    [ ... ],     // 整个 run 开始前跑一次
  "each":     [ ... ],     // 每个 item 跑一次
  "teardown": [ ... ]      // 整个 run 结束后跑一次(含失败退出)
}
```

| 字段 | 必填 | 说明 |
|------|------|------|
| `id` | ✓ | 全局唯一。命名见 `VOCABULARY.md` §3.1 |
| `title` | ✓ | 给人看的标题,可用日文/中文 |
| `version` | ✓ | 语义化版本,改动时手工 bump |
| `profile` | ✓ | 用哪个 profile |
| `vars` | | 工作流级常量,可被 CLI `--var k=v` 覆盖 |
| `source` | | 没有则不遍历,只跑 `setup` + `teardown` |
| `onError` | | 默认 `{ "policy": "ask" }` |
| `setup` / `each` / `teardown` | | step 数组,都可省略 |

---

## 2. Step 调用

```jsonc
{
  "id":   "shot",                    // 可选,但被别人引用时必填
  "use":  "screen.capture_window",   // step id
  "with": { "crop": 6 },             // 参数
  "when": "...",                     // 可选条件,见 §5
  "onError": { "policy": "retry", "times": 3 }   // 可选,覆盖默认
}
```

| 字段 | 说明 |
|------|------|
| `use` | **必填**,必须存在于 `catalog.json` |
| `id` | 步骤局部 id。同一段(setup/each/teardown)内唯一。要被 `{{steps.X.out.Y}}` 引用就必须有 |
| `with` | 参数。按 step manifest 的 `inputs` 校验 |
| `when` | 条件,见 §5 |
| `onError` | 覆盖顶层策略 |
| `label` | 可选,`ebi explain` 里显示的说明文字 |

---

## 3. `source` — 遍历什么

```jsonc
"source": {
  "table":  "worklist",
  "select": {
    "field":       "before_list",
    "pendingWhen": "!= ok"
  },
  "groupBy": "group",       // 可选:按此列分组,见 §7
  "orderBy": "key",         // 可选
  "limit":   0              // 0 = 不限;CLI --limit 会覆盖
}
```

### 3.1 `pendingWhen` 的允许写法(**穷举,不许扩展**)

| 写法 | 含义 |
|------|------|
| `"empty"` | 空 或 `0` |
| `"!= ok"` | 值不是 `ok` |
| `"== ng"` | 值等于 `ng` |
| `"bit !3"` | 位掩码:3 号位组合未全置(位定义在 profile) |
| `"always"` | 全部,不管状态 |

**这不是表达式语言,是五个固定枚举。** 需要更复杂的筛选 → 用
`table.select` step 加参数,或者在 profile 里定义一个命名筛选器。

### 3.2 `ng` 仍算 pending

`"!= ok"` 会把上次判为 `ng` 的行重新选中。这是刻意的:NG 需要人来处理,
不能因为"已经跑过"就消失。

---

## 4. 模板语法

参数值里的 `{{...}}` 会被 runner 求值。

### 4.1 可引用的作用域

| 前缀 | 含义 | 可用范围 |
|------|------|----------|
| `{{vars.X}}` | 工作流常量 | 全部 |
| `{{profile.X.Y}}` | profile 数据 | 全部 |
| `{{run.X}}` | 运行元数据(`runId` / `startedAt` / `operator` / `workDir`) | 全部 |
| `{{item.X}}` | 当前行的某列 | 仅 `each` |
| `{{item.key}}` | 当前行的主键(profile 声明是哪列) | 仅 `each` |
| `{{steps.<id>.out.<field>}}` | 同段内先前 step 的输出 | 同段内,且被引用的 step 必须在前面 |

### 4.2 规则

- **只有取值和字符串拼接**,没有运算:
  `"capture/{{vars.side}}_{{vars.role}}/{{item.key}}.png"` ✓
  `"{{item.count + 1}}"` ✗
- 引用不存在的路径 → `ebi lint` **静态报错**(不是运行时才发现)
- 引用了尚未执行的 step → `ebi lint` 报错
- 整个值就是一个 `{{}}` 时,保留原类型(不会被转成字符串)。
  `"count": "{{steps.parse.out.total}}"` 得到的是 int
- 要输出字面的 `{{`,写 `\{\{`

---

## 5. `when` — 条件

和 `pendingWhen` 一样,是**固定枚举**,不是表达式:

```jsonc
{ "use": "human.gate", "when": "steps.verdict.out.code == unknown" }
```

允许的形式**只有**:

```
<path> == <literal>
<path> != <literal>
<path> exists
<path> empty
```

`<path>` 是 §4.1 的任意引用路径,`<literal>` 是裸字符串或数字。

需要「且 / 或」→ 拆成多个 step,或者做成一个 `verify.*` step 返回一个布尔字段。
**不要在 JSON 里发明布尔代数。**

---

## 6. `onError` — 容错策略

```jsonc
"onError": {
  "policy":   "ask",     // retry | ask | skip | fail
  "times":    3,         // policy=retry 时
  "backoffMs": 800       // policy=retry 时,每次翻倍
}
```

| policy | 行为 |
|--------|------|
| `retry` | 重试 `times` 次,退避递增。用尽后降级为 `ask` |
| `ask` | **默认。** 停下,渲染 `human.gate` 面板:失败原因、上下文、证据路径,让人选 r=重试 / s=跳过这条 / q=中止 |
| `skip` | 记录后跳过这条 item,继续下一条 |
| `fail` | 中止整个 run |

### 6.1 为什么默认是 `ask`

参见 `INTERVIEW.md` 铁律三。一个从不停下但偶尔悄悄记错的工具,比一个经常停下
的工具危险得多 —— 因为没人会去复查它。

先 `ask`,跑够数据、确认某步从不出问题,再改 `retry` 或 `skip`。
反方向(先自动、出错再收紧)代价大得多。

### 6.2 `destructive` 的额外保护

`effects = destructive` 的 step,runner **自动**在前面插一个确认关卡,
不管 `onError` 设成什么。要关掉必须显式写:

```jsonc
{ "use": "excel.replace_sheet", "with": {...}, "confirm": false }
```

`ebi explain` 会把所有 `confirm: false` 高亮出来。

---

## 7. `flow.*` 构造

### 7.1 `flow.foreach` — 隐式

`each` 段本身就是对 `source` 的遍历,不需要显式写 `flow.foreach`。

### 7.2 `flow.group_by` — 分组访问

`source.groupBy` 设了之后,runner 提供额外作用域:

```jsonc
"each": [
  { "use": "browser.navigate", "with": { "url": "{{group.url}}" },
    "once": "group" }                          // 每组只跑一次
]
```

`"once": "group"` 表示这一步在同组的第一条 item 上跑,后续跳过。
用于「一个页面上能查同组的多条」这种场景(旧工具里 HmSnap 的 per-appl 分组)。

### 7.3 `flow.checkpoint` — 标记完成

```jsonc
{ "use": "flow.checkpoint",
  "with": { "field": "before_list", "value": "{{steps.verdict.out.code}}" } }
```

写工作清单 + 写 ledger。**这是断点续跑的唯一依据。**

位掩码形式:

```jsonc
{ "use": "flow.checkpoint", "with": { "field": "composed", "bit": "before" } }
```

`bit` 的名字 → 位值映射在 profile 的 `worklist.json` 里声明。

### 7.4 `flow.call` — 子工作流

```jsonc
{ "use": "flow.call", "with": { "workflow": "shared/refocus-and-search.json",
                                 "vars": { "term": "{{item.key}}" } } }
```

子工作流只有 `each` 段的内容会被内联。用于抽出重复片段。

---

## 8. 完整示例

```jsonc
{
  "id": "before.list.capture",
  "title": "転送状態ページの証跡取得",
  "version": "1.0.0",
  "profile": "host-open",

  "vars": { "side": "before", "role": "list" },

  "source": {
    "table": "worklist",
    "select": { "field": "before_list", "pendingWhen": "!= ok" }
  },

  "onError": { "policy": "ask" },

  "setup": [
    { "use": "human.prepare",
      "with": { "message": "{{profile.pages.list.openHint}}",
                "url":     "{{profile.pages.list.url}}" } },
    { "use": "browser.ensure" },
    { "use": "screen.fit_window",
      "with": { "width":  "{{profile.window.width}}",
                "height": "{{profile.window.height}}" } }
  ],

  "each": [
    { "use": "browser.focus_body" },

    { "use": "browser.tab_to", "with": { "count": "{{profile.pages.list.tabsToForm}}" } },
    { "use": "browser.submit" },
    { "use": "browser.tab_to", "with": { "count": "{{profile.pages.list.tabsToInput}}" } },
    { "use": "browser.fill",   "with": { "text": "{{item.key}}" } },
    { "use": "browser.submit" },

    { "id": "page", "use": "browser.wait_for",
      "with": { "contains":   "{{item.key}}",
                "timeoutSec": "{{profile.pages.list.timeoutSec}}",
                "archiveTo":  "capture/{{vars.side}}_{{vars.role}}/{{item.key}}.txt" } },

    { "use": "browser.assert_page",
      "with": { "text": "{{steps.page.out.text}}",
                "fingerprint": "{{profile.pages.list.fingerprint}}" } },

    { "id": "shot", "use": "screen.capture_window",
      "with": { "saveAs": "capture/{{vars.side}}_{{vars.role}}/{{item.key}}.png" } },

    { "use": "screen.crop",
      "with": { "path":  "{{steps.shot.out.path}}",
                "left":  "{{profile.crop.list.left}}",
                "top":   "{{profile.crop.list.top}}",
                "right": "{{profile.crop.list.right}}",
                "bottom":"{{profile.crop.list.bottom}}" } },

    { "id": "rec", "use": "verify.parse_text",
      "with": { "text":    "{{steps.page.out.text}}",
                "grammar": "{{profile.grammar.list}}" } },

    { "id": "row", "use": "verify.match_record",
      "with": { "records": "{{steps.rec.out.records}}",
                "key":     "{{item.key}}",
                "aliases": "{{profile.key.aliases}}",
                "tieBreak":"newest" } },

    { "id": "verdict", "use": "verify.assert",
      "with": { "record": "{{steps.row.out.record}}",
                "rules":  "{{profile.rules.list}}" } },

    { "use": "human.gate",
      "when": "steps.verdict.out.code == unknown",
      "with": { "reason":   "{{steps.verdict.out.message}}",
                "evidence": "{{steps.shot.out.path}}" } },

    { "use": "flow.checkpoint",
      "with": { "field": "before_list", "value": "{{steps.verdict.out.code}}" } }
  ],

  "teardown": [
    { "use": "progress.status", "with": { "field": "before_list" } }
  ]
}
```

**注意这份 JSON 里没有一个具体系统的名字。** 全部在 `profile` 和 `vars` 里。
换一份工作,这份文件基本能原样抄。

---

## 9. `ebi lint` 检查什么

静态,不运行:

- [ ] `use` 的 step 都存在于 catalog
- [ ] `with` 的参数都在 step 的 `inputs` 里声明过,类型对得上
- [ ] `required` 的参数都给了
- [ ] 所有 `{{}}` 引用都解析得到(profile 路径存在、step id 存在且在前面)
- [ ] `when` / `pendingWhen` 是允许的枚举形式
- [ ] `id` 在同段内唯一
- [ ] 用到 `tier: fallback` 的 step → **警告**,提示需要校准
- [ ] 有 `destructive` 且 `confirm: false` → **警告**,列出位置
- [ ] `profile` 字段指向的 profile 存在且能加载

---

## 10. 版本与兼容

- workflow 的 `version` 手工维护,改了语义就 bump
- schema 本身的版本记在本文件顶部;schema 有破坏性改动时,runner 会拒绝加载
  过旧的 workflow 并提示怎么迁移
- profile 和 workflow 分开演进:改 profile 不需要动 workflow,反之亦然。
  **这是整个设计的目的**
