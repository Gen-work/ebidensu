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
  "id":      "before.transferStatus.capture",
  "title":   "転送状態ページの証跡取得",
  "version": "1.0.0",
  "profile": "host-open",
  "page":    "transferStatus",

  "vars":    { "side": "before" },

  "source":  { ... },      // 遍历什么,见 §3
  "onError": { ... },      // 默认容错策略,见 §6

  "setup":    [ ... ],     // 整个 run 开始前跑一次
  "each":     [ ... ],     // 每个 item 跑一次
  "teardown": [ ... ]      // 整个 run 结束后跑一次(含失败退出)
}
```

| 字段 | 必填 | 说明 |
|------|------|------|
| `id` | ✓ | 全局唯一。命名见 `VOCABULARY.md` §3.2(`<side>.<page>.<verb>`) |
| `title` | ✓ | 给人看的标题,可用日文/中文 |
| `version` | ✓ | 语义化版本,改动时手工 bump |
| `profile` | ✓ | 用哪个 profile |
| `page` | | 绑定到 `profile.pages` 里的哪个 page(P0-R6)——通常是这条工作流**要截图/要判定**的那一个。省略则 `{{page.X}}` 作用域不可用。**只是 `{{page.X}}` 简写的绑定,不是"整条工作流只能碰这一个 page"的限制**:流程里要经过别的 page(比如先填一个独立的检索画面表单,再跳到结果页)时,那个未绑定的 page 依旧可以用完整路径 `{{profile.pages.<名>.X}}` 引用,只是没有简写(`PROFILE-SCHEMA.md` §3.0"两页流程") |
| `vars` | | 工作流级常量,可被 CLI `--var k=v` 覆盖 |
| `source` | | 没有则不遍历,只跑 `setup` + `teardown` |
| `onError` | | 默认 `{ "policy": "ask" }` |
| `setup` / `each` / `teardown` | | step 数组,都可省略 |

---

## 2. Step 调用

```jsonc
{
  "id":   "shot",                    // 必填(P0-R3 起):ledger 的记账键
  "use":  "screen.capture_window",   // step id
  "with": { "crop": 6 },             // 参数
  "when": "...",                     // 可选条件,见 §5
  "onError": { "policy": "retry", "times": 3 }   // 可选,覆盖默认
}
```

| 字段 | 说明 |
|------|------|
| `use` | **必填**,必须存在于 `catalog.json` |
| `id` | **必填**,步骤局部 id。同一段(setup/each/teardown)内唯一 |
| `with` | 参数。按 step manifest 的 `inputs` 校验 |
| `when` | 条件,见 §5 |
| `onError` | 覆盖顶层策略 |
| `label` | 可选,`ebi explain` 里显示的说明文字 |

### 2.1 `id` 为什么是必填的(P0-R3)

`STEP-CONTRACT.md` §6.1 起,ledger 按 **(item, step)** 记账,`step` 就是
这里的 `id` —— 断点续跑靠它判断"这一步做过没有"。如果 `id` 可以省略,
runner 只能退而求其次拿 `use` 当键,而**同一段里出现两次同一个 `use`
是完全合法的**(比如先 `browser.tab_to` 到表单区、填完再 `browser.tab_to`
到别处),两次调用会共用一条 ledger 记录 —— 第二次会被误判为"已完成"而
在 resume 时跳过。所以 `id` 改成必填,不再是"被引用时才需要"。

`ebi lint` 检查:每个 step 调用都有非空 `id`,且同段内不重复。

---

## 3. `source` — 遍历什么

```jsonc
"source": {
  "table":  "worklist",
  "select": {
    "field":       "before_transferStatus",
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
| `{{page.X}}` | 顶层 `page` 绑定的那个 page 的数据(`profile.pages[<page>].X` 的简写),`{{page.id}}` 是 page 名本身 | 全部,且顶层声明了 `page` 字段时才可用(P0-R6) |
| `{{run.X}}` | 运行元数据(`runId` / `startedAt` / `operator` / `workDir` / `window`) | 全部 |
| `{{item.X}}` | 当前行的某列 | 仅 `each` |
| `{{item.key}}` | 当前行的主键显示形(复合键按声明顺序用 `" / "` 拼接) | 仅 `each` |
| `{{item.keySafe}}` | 当前行的主键**文件名安全形**,文件/目录名一律用它,见 `PROFILE-SCHEMA.md` §6.6 | 仅 `each` |
| `{{item.group}}` | 当前行的分组列值(`vocabulary.json` 的 `columns.group` 声明是哪一列,和 `key` 同一种映射方式)。`source.groupBy` 设了之后,**没有单独的 `group` 作用域** —— `"once":"group"` 的 step 就是靠这个访问代表整组的那条 item 的分组值,见 §7.2 | 仅 `each` |
| `{{steps.<id>.out.<field>}}` | 同段内先前 step 的输出 | 同段内,且被引用的 step 必须在前面 |

`{{page.grammar}}` / `{{page.rules}}` 是两个特别的 `{{page.X}}` 路径:它们
不解析进 page 自己的数据,而是解析到 `grammar.json` / `rules.json` 里和
当前 page **同名**的条目(`profile.grammar[<page>]` / `profile.rules[<page>]`
的简写)—— 这两份文件本来就是按 page 名为键的(`PROFILE-SCHEMA.md` §4、§5,
P0-R1),`{{page.grammar}}` 只是省去重复写一遍 page 名。

`run.window` 由 `human.input` 或 CLI `--window` 写入(P2-07 接线),用于
`rules.json` 里 `op: "within"` 的时间窗判定,例如
`{{run.window}}`(`PROFILE-SCHEMA.md` §5)。

### 4.2 规则

- **只有取值和字符串拼接**,没有运算:
  `"capture/{{vars.side}}_{{page.id}}/{{item.keySafe}}.png"` ✓
  `"{{item.count + 1}}"` ✗
- 引用不存在的路径 → `ebi lint` **静态报错**(不是运行时才发现)
- 引用了尚未执行的 step → `ebi lint` 报错
- 整个值就是一个 `{{}}` 时,保留原类型(不会被转成字符串)。
  `"count": "{{steps.parse.out.total}}"` 得到的是 int
- 要输出字面的 `{{`,写 `\{\{`
- **`$Ctx.Session` 里的资源(窗口句柄、Excel COM 对象)不能被模板引用。**
  只有 `outputs` 声明的、JSON-可序列化的字段才能进 `{{steps.X.out.Y}}`
  (`STEP-CONTRACT.md` §3.4,P0-R2)。需要用到某个 step 注册的窗口/工作簿,
  在消费方 `type='session'` 的参数里填**这条工作流自己起的名字**(某个
  更早的 `with.as` 用过的字符串字面量),不走模板 —— manifest 的
  `provides`/`needs` 声明的是资源*种类*(如 `window`),具体叫什么名字
  永远由工作流决定,不是 manifest 写死的
- **禁止嵌套模板**(`{{profile.pages.{{vars.page}}.url}}` 这种写法不合法)
  ——这正是为什么需要 `page` 顶层绑定 + `{{page.X}}` 简写,而不是让
  `vars` 里存一个 page 名再拼进路径

### 4.3 `profile` / `page` 子树的求值(P0-R6)

`{{profile...}}` / `{{page...}}` 取出的**可能是一整棵子树**,不是标量
——比如 `{{page.fingerprint}}` 取出的是 `{ ok:[...], loading:[...], ... }`
整个对象,`{{page.grammar}}` 取出的是一整份 grammar 条目。这些子树内部
如果本身含有 `{{...}}`(比如 `rules.json` 里的 `{{run.window}}`),
**在传给 step 之前递归求值一次**——不会递归到第二层(取出来的结果里
不再解析新的 `{{...}}`),源 JSON 里手写嵌套模板依旧不合法(见 §4.2)。

---

## 5. `when` — 条件

和 `pendingWhen` 一样,是**固定枚举**,不是表达式:

```jsonc
{ "id": "gate", "use": "human.gate", "when": "steps.verdict.out.code == unknown" }
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
  "backoffMs": 800,      // policy=retry 时,每次翻倍
  "byFailure": {         // 按失败 id 覆盖顶层策略,见 §6.0(P0-R5)
    "timeout":   { "policy": "retry", "times": 3 },
    "not_found": { "policy": "ask" }
  }
}
```

| policy | 行为 |
|--------|------|
| `retry` | 重试 `times` 次,退避递增。用尽后降级为 `ask` |
| `ask` | **默认。** 停下,渲染 `human.gate` 面板:失败原因、上下文、证据路径,让人选 r=重试 / s=跳过这条 / q=中止 |
| `skip` | 记录后跳过这条 item,继续下一条 |
| `fail` | 中止整个 run |

### 6.0 按失败种类覆盖:`byFailure`(P0-R5)

顶层 `policy` 是一刀切的,但失败种类天差地别:`browser.wait_for` 超时
(`timeout`)重试是合理的,`not_found`(页面已经加载完、内容确实没有)
重试只会对着同一个错误页面再敲三遍键盘。`manifest.failures` 已经把每种
失败标成了 `transient`(`STEP-CONTRACT.md` §2.1),`onError.byFailure`
把这个信息用起来:

- `byFailure` 里没列出的失败 id,走顶层 `policy`
- **`policy: retry` 只对 `transient = $true` 的失败 id 生效。** 对
  `transient = $false` 的 id 在 `byFailure` 里写 `retry` 是配置错误,
  `ebi lint` 报错(§9,呼应 P1-08)
- `internal_error`(保留 id)默认不重试,除非显式在 `byFailure` 里覆盖

**`warnings`(`STEP-CONTRACT.md` §3.1)不受 `onError` 影响。** 带
`warnings` 的返回值仍然是 `ok = $true`,`onError` 只处理 `ok = $false`
的情形 —— `warnings` 走的是"写进 trace + 计入汇总 + `--guided` 即时显示"
那条路,不会触发关卡或重试。

### 6.1 为什么默认是 `ask`

参见 `INTERVIEW.md` 铁律三。一个从不停下但偶尔悄悄记错的工具,比一个经常停下
的工具危险得多 —— 因为没人会去复查它。

先 `ask`,跑够数据、确认某步从不出问题,再改 `retry` 或 `skip`。
反方向(先自动、出错再收紧)代价大得多。

### 6.2 `destructive` 的额外保护

`effects = destructive` 的 step,runner **自动**在前面插一个确认关卡,
不管 `onError` 设成什么。要关掉必须显式写:

```jsonc
{ "id": "replaceSheet", "use": "excel.replace_sheet", "with": {...}, "confirm": false }
```

`ebi explain` 会把所有 `confirm: false` 高亮出来。

---

## 7. `flow.*` 构造

### 7.1 `flow.foreach` — 隐式

`each` 段本身就是对 `source` 的遍历,不需要显式写 `flow.foreach`。

### 7.2 `flow.group_by` — 分组访问

`source.groupBy` 设了之后,`each` 段仍然只有 `item` 一个作用域(见
§4.1)——**没有单独的 `group` 前缀**。`"once": "group"` 的 step 在组内
第一条 item 上执行,此时 `{{item.X}}` 自然就是这条(代表整个组的)item
的数据,包括 `{{item.group}}`(分组列的值,和 `{{item.key}}` 一样是
profile 声明哪一列的派生访问,见 §4.1):

```jsonc
"each": [
  { "id": "navigate", "use": "browser.navigate",
    "with": { "url": "{{profile.pages.hmResult.url}}?appl={{item.group}}" },
    "once": "group" }                          // 每组只跑一次
]
```

`"once": "group"` 表示这一步在同组的第一条 item 上跑,后续跳过。
用于「一个页面上能查同组的多条」这种场景(旧工具里 HmSnap 的 per-appl 分组)。

### 7.3 `flow.checkpoint` — 标记完成

```jsonc
{ "id": "checkpoint", "use": "flow.checkpoint",
  "with": { "field": "before_transferStatus", "value": "{{steps.verdict.out.code}}" } }
```

写工作清单 + 写 ledger。**这是断点续跑的唯一依据。**

位掩码形式:

```jsonc
{ "id": "checkpoint", "use": "flow.checkpoint", "with": { "field": "composed", "bit": "before" } }
```

`bit` 的名字 → 位值映射在 profile 的 `worklist.json` 里声明。

### 7.4 `flow.call` — 子工作流

```jsonc
{ "id": "refocusAndSearch", "use": "flow.call",
  "with": { "workflow": "shared/refocus-and-search.json",
            "vars": { "term": "{{item.key}}" } } }
```

子工作流只有 `each` 段的内容会被内联。用于抽出重复片段。

### 7.5 断点续跑推演例子(P0-R3)

「重跑跳过已完成的 (item, step)」没说清楚跳过之后,后续 step 引用它的
输出该从哪来 —— 这个例子把 `STEP-CONTRACT.md` §6 的三条规则(ledger 记
outputs、setup 每次重跑、`once: group` 的 ledger 键)在一次真实中断里
过一遍,覆盖**同段引用**、**跨 item**、**`once: group`** 三种情况。

三个 item:`A`、`B` 同组 `G1`,`C` 单独一组 `G2`。`each` 段依次是
`navigate`(`once: group`)→ `shot` → `crop` → `verdict` → `checkpoint`。

跑到 A 做完 `shot`、还没跑 `crop` 时进程被打断(比如 Ctrl+C)。此刻
`run/<runId>/ledger.jsonl` 的内容:

```json
{"runId":"r1","group":"G1","step":"navigate","status":"ok","outputs":{"url":"https://.../G1"},"ts":"..."}
{"runId":"r1","item":"A","step":"shot","status":"ok","outputs":{"path":"capture/before_transferStatus/A.png","width":1280,"height":800},"ts":"..."}
```

（B、C 的任何 step 都还没有记录 —— runner 是按 item 顺序遍历的,还没轮到
它们。）

重跑(`ebi run` 同一个 `runId`)时依次发生:

1. **`setup` 完整重跑**(`STEP-CONTRACT.md` §6.2),`$Ctx.Session` 里的
   浏览器窗口句柄重新注册好 —— ledger 完全不影响这一步。
2. 遍历到 item A,段内第一个 step `navigate`:ledger 里有
   `(group=G1, step=navigate)` 的记录 → **跳过执行**,把它的 `outputs`
   重放进 `steps.navigate.out.*`(这条记录接下来对 B 也生效,见第 6 步)。
3. 仍在 item A,`shot`:ledger 里有 `(item=A, step=shot)` → **跳过执行**,
   重放 `steps.shot.out.path = "capture/before_transferStatus/A.png"` 等
   字段。**这就是"跳过 `shot` 之后 `crop` 引用的输出从哪来"的答案** ——
   从 ledger 重放,不是重新截一次图。
4. 仍在 item A,`crop`:ledger 里**没有** `(item=A, step=crop)` →
   **真正执行**,拿到的 `{{steps.shot.out.path}}` 就是上一步重放出来的
   值 —— 对 `crop` 这个 step 完全透明,它不知道自己上游是重放还是真跑。
5. `verdict`、`checkpoint` 对 A 正常执行,写入新的 ledger 记录。
6. 遍历到 item B(同组 G1),`navigate`:ledger 里已经有
   `(group=G1, step=navigate)` → 跳过,重放**同一条**记录的 `outputs`
   (跨 item 复用,不需要 B 自己再有一条 navigate 记录)。B 的
   `shot`/`crop`/`verdict` 之前从未跑过 → 全部真正执行。
7. 遍历到 item C(组 G2),`navigate` 从未为 G2 跑过(ledger 里没有
   `(group=G2, step=navigate)`)→ 真正执行;后续也全部真正执行。

**不重复截图**(A 的 `shot` 没有再跑一次)、**crop 仍然拿得到正确路径**
(靠重放,不是靠重新执行上游)、**B 不用重新打开页面**(navigate 靠
`once: group` 的重放跨 item 复用)——三点都验证了 P2-06 的断点续跑
验收标准。

---

## 8. 完整示例

```jsonc
{
  "id": "before.transferStatus.capture",
  "title": "転送状態ページの証跡取得",
  "version": "1.0.0",
  "profile": "host-open",
  "page": "transferStatus",

  "vars": { "side": "before" },

  "source": {
    "table": "worklist",
    "select": { "field": "before_transferStatus", "pendingWhen": "!= ok" }
  },

  "onError": { "policy": "ask" },

  "setup": [
    { "id": "prepare", "use": "human.prepare",
      "with": { "message": "{{page.openHint}}",
                "url":     "{{page.url}}" } },
    { "id": "ensure", "use": "browser.ensure", "with": { "as": "mainWindow" } },
    { "id": "fit", "use": "screen.fit_window",
      "with": { "window": "mainWindow",
                "width":  "{{profile.window.width}}",
                "height": "{{profile.window.height}}" } }
  ],

  "each": [
    { "id": "focus", "use": "browser.focus_body", "with": { "window": "mainWindow" } },

    { "id": "tabToForm", "use": "browser.tab_to", "with": { "count": "{{page.tabsToForm}}" } },
    { "id": "submitForm", "use": "browser.submit" },
    { "id": "tabToInput", "use": "browser.tab_to", "with": { "count": "{{page.tabsToInput}}" } },
    { "id": "fill", "use": "browser.fill", "with": { "text": "{{item.Correl_ID_S}}" } },
    { "id": "submitQuery", "use": "browser.submit" },

    { "id": "wait", "use": "browser.wait_for",
      "with": { "contains":   "{{item.Correl_ID_S}}",
                "timeoutSec": "{{page.timeoutSec}}",
                "archiveTo":  "capture/{{vars.side}}_{{page.id}}/{{item.keySafe}}.txt" } },

    { "id": "assert", "use": "browser.assert_page",
      "with": { "text": "{{steps.wait.out.text}}",
                "fingerprint": "{{page.fingerprint}}" } },

    { "id": "shot", "use": "screen.capture_window",
      "with": { "window": "mainWindow",
                "saveAs": "capture/{{vars.side}}_{{page.id}}/{{item.keySafe}}.png" } },

    { "id": "crop", "use": "screen.crop",
      "with": { "path":  "{{steps.shot.out.path}}",
                "left":  "{{page.crop.left}}",
                "top":   "{{page.crop.top}}",
                "right": "{{page.crop.right}}",
                "bottom":"{{page.crop.bottom}}" } },

    { "id": "rec", "use": "verify.parse_text",
      "with": { "text":    "{{steps.wait.out.text}}",
                "grammar": "{{page.grammar}}" } },

    { "id": "row", "use": "verify.match_record",
      "with": { "records": "{{steps.rec.out.records}}",
                "key":     "{{item.key}}",
                "tieBreak":"newest" } },

    { "id": "verdict", "use": "verify.assert",
      "with": { "record": "{{steps.row.out.record}}",
                "rules":  "{{page.rules}}" } },

    { "id": "gate", "use": "human.gate",
      "when": "steps.verdict.out.code == unknown",
      "with": { "reason":   "{{steps.verdict.out.message}}",
                "evidence": "{{steps.shot.out.path}}" } },

    { "id": "checkpoint", "use": "flow.checkpoint",
      "with": { "field": "before_transferStatus", "value": "{{steps.verdict.out.code}}" } }
  ],

  "teardown": [
    { "id": "status", "use": "progress.status", "with": { "field": "before_transferStatus" } }
  ]
}
```

**注意这份 JSON 里没有一个具体系统的名字。** 全部在 `profile`、`page`、
`vars` 里。换一份工作,这份文件基本能原样抄 —— 而且现在换 **page**(比如
从 `transferStatus` 换到 `fileList`)也只改一行:顶层的 `"page":
"transferStatus"`。之前(P0-R1 刚落地、P0-R6 还没做时)`page` 名被写死在
六七个 `profile.pages.transferStatus.X` 路径段里,换页面要全文替换 ——
`{{page.X}}` 就是为了消掉这个代价。

其余几处值得注意:

- `browser.ensure` 的 manifest 声明 `provides = @('window')`(资源
  **种类**,不是名字);这条工作流用 `with.as: "mainWindow"` 把它注册
  进 `$Ctx.Session['mainWindow']`——`mainWindow` 是这条工作流自己起的
  名字,换一条工作流完全可以叫别的。后续需要这个窗口的 step
  (`screen.fit_window`、`browser.focus_body`、`screen.capture_window`)
  在自己 `type='session', sessionKind='window'` 的 `window` 参数里填同一个
  名字字符串,不是各自重新去找前台窗口(`STEP-CONTRACT.md` §3.4,P0-R2)。
- `browser.fill` / `browser.wait_for` 的 `contains` 用的是
  `{{item.Correl_ID_S}}`(具体列),不是 `{{item.key}}`——要打进搜索框、
  要在页面里找的是这一列的原始值,不是复合键拼出来的显示字符串
  (`PROFILE-SCHEMA.md` §6.6)。`verify.match_record` 的 `key` 输入用的
  仍是 `{{item.key}}`,因为它经 `kernel/Key.ps1` 处理,吃的是显示形。
- `archiveTo` / `saveAs` 的文件名用 `{{item.keySafe}}`,不是
  `{{item.key}}`——文件名安全形,见 `PROFILE-SCHEMA.md` §6.6。
- `verify.match_record` 没有 `aliases` 输入 —— 复合键的变体规则
  (`confirmedRules`)由它 dot-source 的 `kernel/Key.ps1` 直接从已加载的
  worklist 数据里读,不是工作流显式传进去的参数(P0-R4)。

---

## 9. `ebi lint` 检查什么

静态,不运行:

- [ ] `use` 的 step 都存在于 catalog
- [ ] `with` 的参数都在 step 的 `inputs` 里声明过,类型对得上
- [ ] `required` 的参数都给了
- [ ] 所有 `{{}}` 引用都解析得到(profile 路径存在、step id 存在且在前面)
- [ ] `when` / `pendingWhen` 是允许的枚举形式
- [ ] 每个 step 调用都有非空 `id`,且在同段内唯一(P0-R3:ledger 记账键,§2.1)
- [ ] 用到 `tier: fallback` 的 step → **警告**,提示需要校准
- [ ] 有 `destructive` 且 `confirm: false` → **警告**,列出位置
- [ ] `profile` 字段指向的 profile 存在且能加载
- [ ] `page` 字段(有的话)指向的 page 在 `pages.json` 里存在(P0-R6)
- [ ] 每个 `type='session'` 参数的字面量名字,都能在它前面找到一个
      `with.as` 等于这个名字、且目标 step 的 `provides` 含匹配种类的调用
      (种类配平,不是名字硬编码在 manifest 里比对,见 `STEP-CONTRACT.md`
      §3.4,P0-R2)
- [ ] `onError.byFailure` 里 `policy: retry` 只用在 `transient = $true` 的失败 id 上(P0-R5)

---

## 10. 版本与兼容

- workflow 的 `version` 手工维护,改了语义就 bump
- schema 本身的版本记在本文件顶部;schema 有破坏性改动时,runner 会拒绝加载
  过旧的 workflow 并提示怎么迁移
- profile 和 workflow 分开演进:改 profile 不需要动 workflow,反之亦然。
  **这是整个设计的目的**
