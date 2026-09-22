# WORKFLOW-SCHEMA — 工作流 JSON 规格

> 状态:**草案**,P0 阶段定稿。**schema: 1**(P0-R15,见 §1 的 `schema` 字段和 §10)。
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
  "schema":  1,
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
  "teardown": [ ... ]      // 整个 run 结束后跑一次;哪些退出路径有这个保证见 §1.1
}
```

| 字段 | 必填 | 说明 |
|------|------|------|
| `schema` | ✓ | 整数,本文件顶部声明的 schema 版本(现在是 `1`)。缺失或大于 runner 支持的版本 → `ebi lint` / runner 拒绝加载(P0-R15,§10) |
| `id` | ✓ | 全局唯一。命名见 `VOCABULARY.md` §3.2(`<side>.<page>.<verb>`) |
| `title` | ✓ | 给人看的标题,可用日文/中文 |
| `version` | ✓ | 语义化版本,改动时手工 bump |
| `profile` | ✓ | 用哪个 profile |
| `page` | | 绑定到 `profile.pages` 里的哪个 page(P0-R6)——通常是这条工作流**要截图/要判定**的那一个。省略则 `{{page.X}}` 作用域不可用。**只是 `{{page.X}}` 简写的绑定,不是"整条工作流只能碰这一个 page"的限制**:流程里要经过别的 page(比如先填一个独立的检索画面表单,再跳到结果页)时,那个未绑定的 page 依旧可以用完整路径 `{{profile.pages.<名>.X}}` 引用,只是没有简写(见 `PROFILE-SCHEMA.md` §3.0) |
| `vars` | | 工作流级常量,可被 CLI `--var k=v` 覆盖 |
| `source` | | 没有则不遍历,只跑 `setup` + `teardown` |
| `onError` | | 默认 `{ "policy": "ask" }` |
| `setup` / `each` / `teardown` | | step 数组,都可省略。`teardown` 的执行保证(哪些退出路径算数)见 §1.1 |

### 1.1 `teardown` 的执行保证(P0-R10)

「整个 run 结束后跑一次」不是无条件的。穷举 runner 能走到的每一条退出
路径:

| 退出路径 | `teardown` 保证跑吗 |
|---------|-------------------|
| 正常跑完(所有 item 处理完) | 保证 |
| `onError.policy=fail` 触发中止 | 保证 |
| step 抛出未预期异常(`internal_error`,`STEP-CONTRACT.md` §3.1) | 保证 |
| 关卡上人选了 `q`(`human.gate` 输出 `action='quit'`,或 `onError.policy=ask` 面板的 q;runner 记为保留失败 `cancelled`,走和 `policy=fail` 同一条路,P0-R13) | 保证 |
| Ctrl+C / 进程被杀 / 终端被关 / 系统重启 | **不保证** |

前三种都发生在同一个 PowerShell 进程的正常控制流里(跑到头,或者被
runner 自己的 `catch` 接住),runner 用 `try { ... } finally { 跑
teardown }` 包住整条执行路径就能保证。Ctrl+C 这类硬中断不一样:
PowerShell 5.1 默认直接终止进程,不触发 `finally`;就算 runner 注册
`Console.CancelKeyPress` 尽力兜底,中断到达时线程完全可能正卡在一次
还没返回的 COM 调用里,`teardown` 连开始跑的机会都没有。

**不保证发生时,资源会怎样**:接受泄漏,不做孤儿检测——这就是这个
项目里 Excel COM 场景一直以来的真实运维方式(操作员手动在任务管理器
里杀多余的 `EXCEL.EXE`),规格如实写清楚这一点,不假装有一套自动回收
机制。哪些资源需要显式释放、由谁释放,见 `STEP-CONTRACT.md` §3.4 第 5
点;一次真实的 Ctrl+C 场景下 Session 资源怎么处理,见 §7.5 的推演例子。

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
| `with` | 参数。按 step manifest 的 `inputs` 校验,唯一例外是 `as`(见 `STEP-CONTRACT.md` §3.4)——runner 保留字段,不进 `inputs`,`provides` 非空的 step 才能用 |
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
  "table":  "wl",           // setup 里某个 table.load 用 with.as 注册的 worklist 实例名(P0-R11)
  "select": {
    "field":       "before_transferStatus",
    "pendingWhen": "!= ok"
  },
  "groupBy": "group",       // 可选:按此列分组,见 §7
  "orderBy": "key",         // 可选
  "limit":   0              // 0 = 不限;CLI --limit 会覆盖
}
```

**`table` 的值是 Session 实例名,不是文件名(P0-R11)。** 工作清单是
种类 `worklist` 的 Session 资源(`STEP-CONTRACT.md` §3.2、§3.4 第 6 点):
`setup` 里 `{ "id": "load", "use": "table.load", "with": { "path": ..., "as": "wl" } }`
把它注册进 `$Ctx.Session`,`source.table` 填同一个名字,runner 遍历的和
`flow.checkpoint` / `progress.status` 写读的就是**同一个内存表**——不存在
两份副本。`ebi lint` 用 §9 的种类配平检查它(必须有一个 `provides` 含
`worklist` 的 `setup` 调用注册过这个名字)。`table.load` `provides` 非空,
所以每次 resume 都真执行(`STEP-CONTRACT.md` §6.2),拿到的永远是磁盘上
最新的表。

CLI `--only <key>[,<key>...]`(P0-R16)按 `{{item.key}}` 的显示形在
`select` 之上再筛一层;`--force` 才忽略 `pendingWhen`。

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

**比较前先经 profile 的 `values` 映射翻译(P0-R11)。** `verdict` 列可以在
`worklist.json` 里声明存储编码(`PROFILE-SCHEMA.md` §6.5,如
`{ "ok": "1", "ng": "2", "unknown": "", "pending": "0" }`),`pendingWhen`
的 `ok` / `ng` 指的是**逻辑值**:runner 把列里读到的存储值先翻译回逻辑值
再比较,`flow.checkpoint` 写入时反向翻译。这是新旧工具在同一份清单上
混跑的前提——工作流 JSON 里永远只写 `ok` / `ng`。

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
| `{{run.X}}` | 运行元数据(`runId` / `startedAt` / `operator` / `workDir` / `timeWindow`)。**持久化在 `run/<runId>/run.json`**(P0-R16):runner 启动时写入,`human.input` / `--time-window` 改 `timeWindow` 时同步更新,`--resume` 时从它恢复,不再问人 | 全部 |
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

**`item.key` / `item.keySafe` / `item.group` 也是保留键,和 `pages.json` 的
`grammar`/`rules`/`id`(`PROFILE-SCHEMA.md` §3)是同一类问题**:它们是派生
访问,不是 worklist 里真实存在的列名。如果 worklist 恰好有一列**字面**就叫
`key`、`keySafe` 或 `group`,`{{item.key}}` 解析到的是派生形,不是那一列的
原始值——原始值仍然可以按列名单独访问(`{{item.key}}` 这个具体列名恰好和
保留名相同时无法区分,建议 worklist 设计阶段避开这三个作为真实列名)。

`run.timeWindow` 由 `human.input` 或 CLI `--time-window` 写入(P2-07 接线),
形状是 `{ "from": "<ISO8601>", "to": "<ISO8601>" }`,用于 `rules.json` 里
`op: "within"` 的时间窗判定,例如 `{{run.timeWindow}}`
(`PROFILE-SCHEMA.md` §5)。**故意不叫 `run.window`**:同一份 §8 示例里
还有 `profile.window.width/height`(浏览器窗口尺寸)和字面量
`"mainWindow"`(`$Ctx.Session` 里的窗口句柄名,§3.4)——三个概念都叫
"window" 太容易读串,`timeWindow` 消掉这一个。

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
  `provides`(注册方)和 `inputs` 里 `type='session'` 参数的 `sessionKind`
  (消费方,P0-R10:`needs` 里不重复声明这件事,见 `STEP-CONTRACT.md` §4)
  声明的都是资源*种类*(如 `window`),具体叫什么名字永远由工作流决定,
  不是 manifest 写死的
- **禁止嵌套模板**(`{{profile.pages.{{vars.page}}.url}}` 这种写法不合法)
  ——这正是为什么需要 `page` 顶层绑定 + `{{page.X}}` 简写,而不是让
  `vars` 里存一个 page 名再拼进路径

### 4.3 `profile` / `page` 子树的求值(P0-R6)

`{{profile...}}` / `{{page...}}` 取出的**可能是一整棵子树**,不是标量
——比如 `{{page.fingerprint}}` 取出的是 `{ ok:[...], loading:[...], ... }`
整个对象,`{{page.grammar}}` 取出的是一整份 grammar 条目。这些子树内部
如果本身含有 `{{...}}`(比如 `rules.json` 里的 `{{run.timeWindow}}`),
**在传给 step 之前递归求值一次**——不会递归到第二层(取出来的结果里
不再解析新的 `{{...}}`),源 JSON 里手写嵌套模板依旧不合法(见 §4.2)。

---

## 5. `when` — 条件

和 `pendingWhen` 一样,是**固定枚举**,不是表达式:

```jsonc
{ "id": "checkpoint", "use": "flow.checkpoint", "when": "steps.gate.out.action != skip" }
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

### 5.1 被 `when` 跳过的 step,它的输出是什么(P0-R13)

`when` 是运行期才知道真假的,所以"引用了没执行的 step"这件事 lint 判不了。
定死:runner 为被跳过的 step 在作用域里放 `@{ ok=$true; skipped=$true }`,
manifest `outputs` 声明的**每个字段都存在、值为 `null`**;引用它不是
错误。`when` 的四种形式对 `null` 的判法:`exists` 为假、`empty` 为真、
`==` / `!=` 按字面比较(`null` 不等于任何字面量)。于是"上一步跳过了
就也跳过"写成 `when: "steps.x.out.skipped != true"`。ledger 里记一条
`status='skipped'`,resume 时照常重放(`STEP-CONTRACT.md` §6.1)。

### 5.2 关卡的结论怎么进 checkpoint(P0-R13)

旧工具的判定流是三态收口到两态:`ok` → 写完成,`ng` → 写 NG,`ask` →
问人,人答 o / n / s / q → 完成 / NG / 留 pending / 中止。要让"人在关卡
上答了什么"覆盖 `verify.assert` 的结论,而 JSON 里又不许写"如果 gate 跑了
取 gate 的、否则取 verdict 的"这种表达式,做法是**让关卡 step 自己做
条件,不靠 `when`**:

- `human.gate` **总是执行**,输入 `code`(`verify.assert` 的结论)和
  `askWhen`(默认 `["unknown"]`)。`code` 不在 `askWhen` 里 → 静默直通,
  输出 `code` 原值、`action='pass'`;在里面 → 渲染面板问人:`Enter` →
  `code='ok'`、`n` → `code='ng'`、`s` → `code=''` + `action='skip'`、
  `q` → `action='quit'`(runner 记 `cancelled`,走 §1.1 的中止路径)。
- `flow.checkpoint` 统一写 `{{steps.gate.out.code}}`,并带
  `when: "steps.gate.out.action != skip"`——`s` 是"留 pending",不写盘。
- `onError.policy=ask` 的面板复用 `human.gate` 的渲染和 r / s / q 语义,
  不另写一套。

§8 的完整示例就是这个形状。

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
- `STEP-CONTRACT.md` §3.1 的全部保留 id(`needs_unmet` / `session_missing` /
  `session_invalid` / `session_conflict` / `cancelled`)都可以在 `byFailure`
  里引用;`ebi lint` 对它们不查"manifest 里有没有列"(P0-R15)。
  `session_invalid` 是其中唯一 `transient` 的——重跑 `setup` 就能重建

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

**`"once": "groupEnd"` —— 和 `"once": "group"` 对偶,组尾钩子(P0-R10
第四轮)。** 在同组**最后一条** item 处理完之后跑一次,用来释放
`once:"group"` 在组开头注册进 `$Ctx.Session` 的组级资源。

为什么需要它:按交付物分组、每组开一个工作簿的 compose 类工作流,长这样

```jsonc
"source": { "table": "worklist", "groupBy": "deliverable" },

"setup": [
  { "id": "app", "use": "excel.ensure_app", "with": { "as": "xl" } }
],
"each": [
  { "id": "open", "use": "excel.open",
    "with": { "app": "xl", "path": "{{item.group}}.xlsx", "as": "wb" },
    "once": "group" },
  { "id": "insert", "use": "excel.insert_picture",
    "with": { "workbook": "wb",
              "path": "capture/{{vars.side}}_{{page.id}}/{{item.keySafe}}.png" } },
  { "id": "close", "use": "excel.close",
    "with": { "workbook": "wb" },
    "once": "groupEnd" }              // 组尾释放,不是 teardown
],
"teardown": [
  { "id": "quit", "use": "excel.quit_app", "with": { "app": "xl" } }
]
```

两种作用域在这一份里同时出现:Application 一次 run 一个(`setup` 注册 /
`teardown` 释放),Workbook 一组一个(`once:"group"` / `once:"groupEnd"`)。
Excel 的生命周期必须是这四个 step,不能让一个 `excel.open` 同时产出
Application 和 Workbook——原因见 `STEP-CONTRACT.md` §3.4 第 6 点。

如果没有 `once:"groupEnd"`、只能靠 `teardown` 释放:`open` 每组都用**同一个**
`with.as` 名字 `wb` 重新注册(`with.as` 是字符串字面量,不走 `{{}}` 模板,见
`STEP-CONTRACT.md` §3.4 第 3 点,**不能**按组变化)——第 2 组的 `open` 一跑,
`$Ctx.Session['wb']` 就被覆盖成新工作簿的引用,第 1 组那个 Workbook/Application
COM 对象当场变成孤儿:没有名字能传给 `excel.close`,谁也关不掉。`teardown`
只跑一次、只关得掉最后一组,前面所有组全部泄漏成挂起的 `EXCEL.EXE`
——`ReplaceEvidence.ps1:148` 的 `Group-Object Excel_NAME`、以及它在这里的
替代品 `P4-20 workflows/*.compose.json`,就是这个按交付物分组开工作簿的
真实场景。`STEP-CONTRACT.md` §3.4 第 3 点新增的"同名重复注册在名字仍然
活着时是运行期失败"这条规则,就是为了让这种覆盖在第 2 组一开始就报错,
而不是悄悄泄漏。`once:"groupEnd"` 把释放挪到组尾,和组头的注册配对,
`teardown` 只管 `setup` 里注册的东西——这正是 `STEP-CONTRACT.md` §3.4
第 5 点"释放点必须紧跟注册点所在的作用域"的落地。

规则:
- ledger 键和 `once:"group"` 一样是 **(group, step)**(`STEP-CONTRACT.md`
  §6.3)
- `source.groupBy` 没设时用 `once:"groupEnd"` 是配置错误,`ebi lint` 报错
  (见 §9)
- runner 怎么知道"最后一条":`source` 按 `groupBy` 排序后遍历,组切换时
  (或整个 `source` 遍历结束时)触发上一组的 `groupEnd` step
- **resume 时 `open`/`close` 总是真执行**,不按 ledger 跳过(它们
  `provides`/`releases` 非空,见 `STEP-CONTRACT.md` §6.2)。代价是一个上次
  已整组跑完的组,resume 时会被多开关一次(不写数据);不付这个代价,
  `wb` 在新进程里就永远不会被重新注册,这条工作流被 Ctrl+C 之后**永久**
  跑不完
- 组内最后一条 item 被 `onError` 的 `skip` 跳过时,`groupEnd` **照跑**
  ——资源已经注册了,不释放就泄漏;`policy: fail` 中止时走 §1.1 的
  `teardown` 保证(硬中断不保证,接受泄漏)

### 7.3 `flow.checkpoint` — 标记完成

```jsonc
{ "id": "checkpoint", "use": "flow.checkpoint",
  "when": "steps.gate.out.action != skip",
  "with": { "worklist": "wl",
            "field": "before_transferStatus", "value": "{{steps.gate.out.code}}" } }
```

写工作清单 + 写 ledger。**这是断点续跑的唯一依据。** `worklist` 是
`type='session'; sessionKind='worklist'` 的输入(P0-R11),填 `setup` 里
`table.load` 注册的名字;`value` 是逻辑值(`ok` / `ng` / `unknown` /
`''`),写入前经 profile 的 `values` 映射翻译成存储编码(§3.1),写完
原子落盘。`value` 取 `steps.gate.out.code` 而不是 `steps.verdict.out.code`,
配 `when` 跳过 `action == skip`——理由见 §5.2。

位掩码形式:

```jsonc
{ "id": "checkpoint", "use": "flow.checkpoint", "with": { "field": "composed", "bit": "before" } }
```

`bit` 的名字 → 位值映射在 profile 的 `worklist.json` 里声明。

### 7.4 `flow.call` — 子工作流 ⚠ 暂未排期

```jsonc
{ "id": "refocusAndSearch", "use": "flow.call",
  "with": { "workflow": "shared/refocus-and-search.json",
            "vars": { "term": "{{item.key}}" } } }
```

子工作流只有 `each` 段的内容会被内联。用于抽出重复片段。**这张卡不在
BACKLOG 里,P1 的 32 个 MVP step 也没排它** —— 实现前先读下面的命名空间
规则,不要假设内联的 id 会自动避让。

**内联 id 必须加前缀,否则会撞上 P0-R3 刚堵上的 ledger 键冲突。**
`flow.call` 的用途写明是"抽出重复片段"(预期被多处调用,或同一处循环
调用),而 `each` 内联进父工作流后,子工作流自己的 step id(如
`shot`/`crop`)会和父工作流里同名的 step,或者**同一个子工作流被调用
两次**产生的两份 `shot`/`crop`,直接重名 —— 违反 §2.1 的"同段内唯一"、
也会让 §6.1 的 ledger 记账把两次调用的 outputs 记成一条,resume 时把
从没跑过的那次误判为"已完成"跳过。这正是 P0-R3 把 `id` 改成必填想消灭
的那类 bug,不能让 `flow.call` 从后门带回来。

规则:内联时,子工作流每个 step 的 id 自动加上
`<flow.call 调用点的 id>.` 前缀(如 `refocusAndSearch.shot`)。父工作流
引用子工作流内某个 step 的输出时,写全前缀:
`{{steps.refocusAndSearch.shot.out.path}}`。

**子工作流自己内部的引用也要一起改写。** 子工作流是独立文件,写的时候
按自己的 step id 写引用(比如 `crop` 引用同一个子工作流里的
`{{steps.shot.out.path}}`)——内联时如果只改写 id 本身、不改写子工作流
内部这些引用,`shot` 变成了 `refocusAndSearch.shot`,但 `crop` 里那句
`{{steps.shot.out.path}}` 还在找一个不存在的 `shot`,变成未定义引用。
内联必须同时改写子工作流内部对自己 step 的所有引用,加上同一个前缀。

**id 里出现 `.` 之后,`{{steps.<id>.out.<field>}}` 的解析不能再朴素地按
`.` 切分。** `refocusAndSearch.shot` 本身就带一个 `.`,
`{{steps.refocusAndSearch.shot.out.path}}` 这个字符串里有三个 `.`——
解析器必须专门找 `.out.` 这个边界(`steps` 和 `out` 之间的一切都是 id,
`out` 之后的一切都是字段路径),而不是假设"第二段就是 id、第三段就是
`out`"这种位置写死的切法。

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

**这次 Ctrl+C 没有跑 `teardown`**(§1.1,P0-R10:硬中断不保证)——如果
这条工作流的 `setup` 里有一步用 `with.as` 注册过浏览器窗口(常见情况,
`browser.ensure` 一类 step 通常这么做),那个实例名和它背后的真实窗口
就跟着这次中断一起停在半空,没有任何释放 step 被调用。下面的重跑是
全新进程,`$Ctx.Session` 从空表开始——第 1 步 `setup` 重跑会重新
`ensure` 一个窗口并重新注册**同一个名字**,旧的那个窗口不会被专门找
出来处理:是被动接受的泄漏(浏览器窗口泄漏无害,人工关掉即可),不是
这个例子要验证的东西——这个例子验证的是 ledger 重放,不是资源回收。
换成 Excel 场景(`with.as` 注册的是工作簿/COM 对象)结论完全一样,只是
泄漏的代价从"多一个窗口"变成"多一个挂起的 `EXCEL.EXE` 进程",这正是
`STEP-CONTRACT.md` §3.4 第 5 点要求 `teardown` 里显式调用释放 step 的
理由——但这条要求管的是"跑到 teardown 时会不会释放",管不了"这次
根本没跑到 teardown"。

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

### 7.6 第二个推演例子:组级资源在 resume 后怎么回来(P0-R10 第五轮)

§7.5 的例子里 `navigate` 不注册任何 Session 资源,所以它只考了 ledger 重放。
把它换成 §7.2 那条 compose 工作流(`open` 用 `once:"group"` 注册 `wb`),
中断/重跑的走法就完全不同了——这一节把它单独走一遍,因为**照 §6.1 的
跳过规则字面理解会得到一条永远跑不完的工作流**。

组 `G1` 有 5 条 item。`open`(`once:"group"`)执行并记进 ledger,item 1-3
做完,Ctrl+C(§1.1:硬中断不跑 `teardown`,上一进程的 Excel 泄漏,人工清)。

重跑(同一个 `runId`,**新进程,`$Ctx.Session` 空**):

1. runner 读 `ledger.jsonl` 构建"已完成集合"时,**把 `provides`/`releases`
   非空的 step 的历史记录排除在外**(§6.2)——所以 `(G1, open)` 和
   `(G1, close)` 不算数,item 1-3 的 `insert` 等记录照常算数。
2. `setup` 完整重跑,`excel.ensure_app` 重新注册 `xl`。
3. 遍历到 item 1(属于 G1),`open`:`once:"group"` 看的是**本次进程**的
   ledger,组内还没有它的新记录 → **真正执行**,`wb` 重新进入
   `$Ctx.Session`。**这就是"跳过 `open` 之后 `insert` 引用的 `wb` 从哪来"
   的答案** —— 不是从 ledger 重放(句柄/COM 对象根本不进 ledger),是靠
   这一步真的重新注册。
4. item 1-3 的 `insert`:ledger 里有记录 → 跳过。**代价在这里**:这一组前
   三条其实白开了一次工作簿。这是让"注册-释放"在新进程里重新配平必须付的
   钱(§6.2)。
5. item 4、5 的 `insert`:ledger 里没有 → 真正执行,`with.workbook: "wb"`
   查得到,不再报"名字未注册"。
6. G1 最后一条 item 处理完 → `close`(`once:"groupEnd"`)同样因为
   `releases` 非空而不被历史 ledger 跳过 → 真正执行,`wb` 释放。
7. 进到 G2,`open` 再次注册 `wb` —— 因为第 6 步已经释放过,不触发 §3.4
   第 3 点的"同名重复注册非法"。**这就是 `provides` 和 `releases` 必须
   成对豁免的原因**:只豁免 `open` 而跳过 `close`,`wb` 会一直占着名字,
   G2 在这一步当场失败。

对照 §7.5:那个例子验证的是**输出重放**,这个例子验证的是**资源重建**——
两条路互不替代,ledger 能重放的只有 JSON 值,Session 资源永远只能靠重新
执行注册步骤拿回来。

---

## 8. 完整示例

```jsonc
{
  "schema": 1,
  "id": "before.transferStatus.capture",
  "title": "転送状態ページの証跡取得",
  "version": "1.0.0",
  "profile": "host-open",
  "page": "transferStatus",

  "vars": { "side": "before" },

  "source": {
    "table": "wl",
    "select": { "field": "before_transferStatus", "pendingWhen": "!= ok" }
  },

  "onError": { "policy": "ask" },

  "setup": [
    { "id": "load", "use": "table.load",
      "with": { "path": "{{profile.worklist.file}}", "as": "wl" } },
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

    { "id": "tabToForm", "use": "browser.tab_to",
      "with": { "window": "mainWindow", "count": "{{page.tabsToForm}}" } },
    { "id": "submitForm", "use": "browser.submit", "with": { "window": "mainWindow" } },
    { "id": "tabToInput", "use": "browser.tab_to",
      "with": { "window": "mainWindow", "count": "{{page.tabsToInput}}" } },
    { "id": "fill", "use": "browser.fill",
      "with": { "window": "mainWindow", "text": "{{item.Correl_ID_S}}", "verifyChange": true } },
    { "id": "submitQuery", "use": "browser.submit", "with": { "window": "mainWindow" } },

    { "id": "wait", "use": "browser.wait_for",
      "with": { "window":     "mainWindow",
                "contains":   "{{item.Correl_ID_S}}",
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

    { "id": "meta", "use": "file.write_json",
      "with": { "path": "capture/{{vars.side}}_{{page.id}}/{{item.keySafe}}.meta.json",
                "data": { "row": { "index": "{{steps.row.out.index}}",
                                   "count": "{{steps.row.out.count}}" } } } },

    { "id": "verdict", "use": "verify.assert",
      "with": { "record": "{{steps.row.out.record}}",
                "rules":  "{{page.rules}}" } },

    { "id": "gate", "use": "human.gate",
      "with": { "code":     "{{steps.verdict.out.code}}",
                "askWhen":  ["unknown"],
                "reason":   "{{steps.verdict.out.message}}",
                "evidence": "{{steps.shot.out.path}}" } },

    { "id": "checkpoint", "use": "flow.checkpoint",
      "when": "steps.gate.out.action != skip",
      "with": { "worklist": "wl",
                "field": "before_transferStatus", "value": "{{steps.gate.out.code}}" } }
  ],

  "teardown": [
    { "id": "status", "use": "progress.status",
      "with": { "worklist": "wl", "field": "before_transferStatus" } }
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
- `"schema": 1` 是必填的版本标记(P0-R15,§1、§10)。
- `table.load` 在 `setup` 里用 `as: "wl"` 注册工作清单,`source.table`、
  `flow.checkpoint` 的 `worklist`、`progress.status` 的 `worklist` 都填这个
  名字——runner 遍历的和 step 写的是同一个内存表(P0-R11,§3)。
- **每个发键的 step 都带 `window: "mainWindow"`**(`tab_to` / `submit` /
  `fill` / `wait_for`,P0-R12):发键前把这个窗口拉到前台并核对,而不是
  对着"当下前台"发——`gate` 一问完人,前台就在控制台了。`fill` 开了
  `verifyChange`,填完对比页面文本,没变化就 `no_effect`。
- `gate` **没有 `when`**,而是拿 `code` + `askWhen` 自己决定问不问
  (P0-R13,§5.2);`checkpoint` 写的是 `steps.gate.out.code`,并用 `when`
  跳过 `action == skip`(留 pending)。
- `meta` 把 `match_record` 算出的行号 / 条数写进
  `capture/<side>_<page>/<keySafe>.meta.json` 侧车(P0-R14)——annotate
  工作流靠它平移红框,ledger 按 run 隔离,它读不到这里的 outputs。

---

## 9. `ebi lint` 检查什么

静态,不运行:

- [ ] `schema` 存在、是整数、不大于 runner 支持的版本(P0-R15,§1)
- [ ] `source.table` 填的名字,在 `setup` 里有一个 `provides` 含 `worklist`
      的调用用同名 `with.as` 注册过(P0-R11,§3;是下面种类配平检查的
      一个特例,单列出来因为它不是某个 step 的 `type='session'` 输入,
      而是 runner 自己的消费)
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
- [ ] `setup` 段里的每个 step 都是 `idempotent = $true`(`STEP-CONTRACT.md` §6.2,P0-R3——`setup` 每次 resume 都重跑,非幂等 step 出现在这里是契约违反)
- [ ] `"once": "groupEnd"` 只在 `source.groupBy` 有值时合法,没有 `groupBy`
      的 `each` 段里出现 `once:"groupEnd"` → 报错(§7.2,P0-R10 第四轮)
- [ ] 每个通过 `with.as` 注册的资源名,如果它的种类在 `STEP-CONTRACT.md`
      §3.4 第 6 点的种类表里标了 `mustRelease = $true`,必须存在一个同名、
      种类匹配的释放调用(`releases` 覆盖该种类的 step),且这个释放调用
      在**同一段的更后面**(`each` 里 `once:"group"` 注册 → `each` 里
      `once:"groupEnd"` 释放)或**更后的段**(`setup` 注册 → `teardown`
      释放);配不上就报错。种类标了 `mustRelease = $false`(比如
      `window`——窗口句柄泄漏无害)→ 不要求。**判据是种类表,不是
      "catalog 里现在有没有恰好带 `releases` 的 step"**——后者会让新增
      一个释放 step 使所有已有工作流集体变红,而它们一行都没改
      (`STEP-CONTRACT.md` §3.4 第 6 点,P0-R10 第四轮)
- [ ] `provides` 或 `releases` 非空的 step,`idempotent` 必须是 `$true`
      (resume 时它们总是真执行,`STEP-CONTRACT.md` §6.2,P0-R10 第五轮)
- [ ] worklist 里所有 `role: key` 的列都出现在 `key.columns` 里,反之亦然(`PROFILE-SCHEMA.md` §6.1)
- [ ] 涉及 key 比较的 step(`verify.match_record`/`file.find`/`file.newest`/`excel.find_anchor`……)都走 `table.key` 的规范化,没有 step 自己写比较绕过它(`PROFILE-SCHEMA.md` §6.4)

> 这四项(连同上面 `page`/`session`/`onError.byFailure` 三项)散落声明在
> `STEP-CONTRACT.md` 和 `PROFILE-SCHEMA.md` 各自的章节里,**这份清单是
> 唯一汇总处**——P1-08 只读这一份,散在别处的规则如果没抄过来会被
> 实现直接漏掉。新增任何一条 lint 规则,定义它的章节和这份清单都要改。

---

## 10. 版本与兼容

- workflow 的 `version` 手工维护,改了语义就 bump
- schema 本身的版本记在本文件顶部(现在是 `1`),每条 workflow 用顶层
  `schema` 字段声明它按哪一版写的(§1,P0-R15);schema 有破坏性改动时
  这里的数字 bump,runner 对 `schema` 缺失或大于自己支持的版本拒绝加载并
  提示怎么迁移。加字段的那天所有已有工作流都得补一行——所以从第一条
  工作流起就带着它
- profile 和 workflow 分开演进:改 profile 不需要动 workflow,反之亦然。
  **这是整个设计的目的**
