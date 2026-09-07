# STEP-CONTRACT — 模块步骤契约

> 状态:**草案**,P0 阶段定稿。
> 这是整个项目的地基。改这份文档 = 改所有 step。

## 0. 一个 step 是什么

一个 step 是**一个正交能力原语**:做一件事,不知道自己在什么业务里。

判断一个东西该不该是 step:

| 是 step | 不是 step |
|---------|-----------|
| 「用 Ctrl+F 查找一个字符串」 | 「查找相关 ID」← 业务知识漏进来了 |
| 「截图并按四边裁剪」 | 「截 HM 画面的图」← 同上 |
| 「按规则表判定一条记录」 | 「判断批处理是否异常终了」← 规则应该是数据 |
| 「等待页面文本匹配某模式」 | 「等 MQ 页面加载完」← 模式应该是参数 |

**检验方法**:把这个 step 的说明念一遍,如果里面出现了任何具体系统/画面/
字段的名字,它就设计错了。

---

## 1. 文件形态

```
modules/<group>/<group>.<verb>.ps1
```

例:`modules/browser/browser.find.ps1`

每个文件导出**两样东西,不多不少**:

```powershell
$Manifest = @{ ... }                            # 见 §2
function Invoke-Step { param($In, $Ctx) ... }   # 见 §3
```

### 1.1 PS 5.1 硬性约束

沿用现有仓库的铁律:

- **文件里不许有 `param()` 块**(step 文件是被 dot-source 的)
- **源码必须是纯 ASCII**。任何日文/中文字面量用 `[char]` 构造,或者从 profile
  读。理由:JP locale 的主机上,无 BOM 的 `.ps1` 里的非 ASCII 会乱码 —— 这个
  bug 在旧仓库里静默破坏过 owner 匹配
- 编码 UTF-8 **无 BOM**,由 `Check-Encoding.ps1` 强制
- 不用 PowerShell `class`(跨 dot-source 有作用域坑),用 hashtable +
  `[pscustomobject]`
- **禁止 `@($hashtable[$key])` 这个包装形状**。PS 5.1 的 binder 在
  `List[object]` 上会抛「参数类型不匹配」。旧仓库两次同类事故都是这个模式

---

## 2. Manifest

```powershell
$Manifest = @{
  # --- 身份 ---
  id       = 'browser.find'      # 必须等于文件名(去掉 .ps1)
  group    = 'browser'
  summary  = 'Ctrl+F search for an exact string; report whether it hit'
  tier     = 'core'              # core | fallback   见 §5

  # --- 契约 ---
  effects  = 'ui'                # pure | read | ui | write | destructive
  needs    = @('foreground')     # 运行前置条件 / Session 资源依赖,见 §4、§3.4
  provides = @()                 # 本 step 能注册进 $Ctx.Session 的资源种类,见 §3.4
  releases = @()                 # 本 step 会释放 $Ctx.Session 里哪个种类的资源,见 §3.4 第 5 点(P0-R10)
  idempotent = $true             # 重复执行是否安全

  # --- 输入 ---
  inputs   = @{
    term       = @{ type='string'; required=$true;  desc='exact string to search for' }
    closeAfter = @{ type='bool';   default=$true;   desc='press Esc afterwards' }
  }

  # --- 输出 ---
  outputs  = @{
    hit  = @{ type='bool' }
    rect = @{ type='rect'; desc='pixel rect of the active match, null when no hit' }
  }

  # --- 失败模式(P0-R5:每项标 transient) ---
  failures = @(
    @{ id = 'not_found';            transient = $false }
    @{ id = 'no_foreground_window'; transient = $true  }
  )

  # --- 给 Agent 的示例 ---
  example  = @{ use='browser.find'; with=@{ term='{{item.key}}' } }
}
```

### 2.1 字段说明

| 字段 | 必填 | 说明 |
|------|------|------|
| `id` | ✓ | 全局唯一,`<group>.<verb>` 形式,必须与文件名一致 |
| `group` | ✓ | CATALOG 的 **9 组**:`browser` / `screen` / `file` / `excel` / `table` / `verify` / `human` / `progress` / `flow`,和 `modules/` 下的目录一一对应。注意 `flow` 组只装真的用 `use` 调的 step(`flow.checkpoint` / `flow.call`);`flow.foreach` / `flow.group_by` 是 runner 构造,不是 step,见 `WORKFLOW-SCHEMA.md` §7.1、§7.2 |
| `summary` | ✓ | **一句话,英文,不超过 80 字符**。这是 Agent 挑 step 的主要依据 |
| `tier` | ✓ | `core` 或 `fallback`(见 §5) |
| `effects` | ✓ | 副作用等级,见 VOCABULARY §4 |
| `needs` | | 前置条件数组,见 §4。**不写** `session:<种类>` 这一类(P0-R10):消费某种 Session 资源这件事,已经由 `inputs` 里那个 `type='session'` 参数的 `sessionKind` 说清楚了,见 §3.4 |
| `provides` | | 本 step 能注册进 `$Ctx.Session` 的资源**种类**数组(不是实例名),见 §3.4。空则 `@()` |
| `releases` | | 本 step 会释放 `$Ctx.Session` 里哪个种类的资源(与 `provides` 对称,声明的也是**种类**不是实例名),见 §3.4 第 5 点(P0-R10)。空则 `@()` |
| `idempotent` | ✓ | `$false` 的 step,runner 在续跑时不会自动重放 |
| `inputs` | ✓ | 参数定义。空则 `@{}` |
| `outputs` | ✓ | 返回字段定义。空则 `@{}` |
| `failures` | ✓ | **穷举**所有可能的失败标识,每项 `@{ id=; transient=<bool> }`(P0-R5)。runner 用它校验返回值,`transient=$true` 是 `onError.policy=retry` 能重试的必要条件(见 `WORKFLOW-SCHEMA.md` §6);`internal_error` 是保留失败 id(§3.1),不需要在这里列出 |
| `example` | ✓ | 至少一个可运行的调用例 |
| `notes` | | 补充说明,支持多行。会渲染进 CATALOG.md |

### 2.2 `inputs` 的字段类型

| type | PS 类型 | 说明 |
|------|---------|------|
| `string` | `[string]` | |
| `int` | `[int]` | |
| `bool` | `[bool]` | |
| `path` | `[string]` | 会被 runner 做路径规范化 |
| `rect` | hashtable | `@{ X=; Y=; W=; H= }` |
| `list` | `[object[]]` | |
| `map` | hashtable | |
| `session` | `[string]` | `$Ctx.Session` 里某个已注册资源的名字(字面量,不是 `{{}}` 模板);必须带 `sessionKind`,见 §3.4 |
| `any` | — | 尽量避免 |

`type = 'session'` 的参数额外必填 `sessionKind`(字符串,资源种类,如
`'window'`),`ebi lint` 用它做 §3.4 的种类配平检查。

每个参数可有:`required`(bool)、`default`、`desc`(英文一句话)、
`enum`(允许值数组)。

**`required=$true` 和 `default` 互斥**,`ebi lint` 会检查。

---

## 3. `Invoke-Step`

```powershell
function Invoke-Step {
    param($In, $Ctx)
    # $In  : hashtable,已经过 schema 校验和模板求值的参数
    # $Ctx : 只读上下文,见 §3.2
    ...
    return @{ ok = $true; hit = $true; rect = $r }
}
```

### 3.1 返回值约定

**必须返回一个 hashtable,必须含 `ok` 键。**

成功:
```powershell
@{ ok = $true;  <manifest.outputs 里声明的字段...> }
```

失败:
```powershell
@{ ok = $false; failure = 'not_found'; message = 'no row matched' }
```

规则:
- `failure` 的值**必须**等于 `$Manifest.failures` 某一项的 `id`,否则 runner 判为契约违反
- `message` 是给人看的一句话,英文,可选
- 失败时**也可以**带 outputs 字段(比如部分结果),runner 不会用,但会进 trace
- **不许抛异常表达业务失败**。异常只用于「代码写错了」这类真正的意外;
  runner 捕获后统一记为 `failure = 'internal_error'`(这是**保留**失败
  id,不需要出现在 `$Manifest.failures` 里)

**可选的 `warnings`(P0-R5,非致命异常的标准通道)。** 成功或失败都可以带:

```powershell
@{ ok = $true; hit = $true; rect = $r;
   warnings = @(
     @{ code = 'unrecognized_line'; message = '3 lines did not match any row pattern';
        data = @{ lines = @(4, 9, 12) } }
   ) }
```

`warnings` 不影响 `ok` 的值,但 runner **必须**:1) 写进 trace;2) 计入
run 结束时的汇总;3) `--guided` 模式下立即显示在面板上。**不要为"需要
上报但不算失败"的情况发明私有输出字段**——这正是旧工具"静默丢行"事故
(`SnapVerify`/`GfixJobList` 解析时,页面上明明有的行被正则漏掉、返回值
没有任何字段能表达"我漏了几行")的根因:数据返回了,但没人看得见。

### 3.2 `$Ctx` 提供什么(只读)

| 字段 | 说明 |
|------|------|
| `$Ctx.WorkDir` | 工作目录绝对路径 |
| `$Ctx.RunId` | 本次运行 id |
| `$Ctx.Profile` | 已加载并合并好的 profile(hashtable) |
| `$Ctx.Log` | `$Ctx.Log.Info('...')` / `.Warn(...)` / `.Debug(...)` |
| `$Ctx.DryRun` | `$true` 时,有副作用的 step **必须**只打印不执行 |

**`$Ctx` 里没有别的 step 的输出。** step 之间只通过 workflow JSON 的模板引用
传值 —— 这是正交性的保证。

### 3.3 `DryRun` 的处理

`effects` 为 `write` / `destructive` / `ui` 的 step **必须**处理 `DryRun`:

```powershell
if ($Ctx.DryRun) {
    $Ctx.Log.Info(('would save to {0}' -f $In.path))
    return @{ ok = $true; path = $In.path }
}
```

`pure` / `read` 的 step 可以照常执行。

### 3.4 `$Ctx.Session` —— 跨 step 的资源通道(P0-R2)

`{{steps.X.out.Y}}` 只能传 JSON 值,而句柄(浏览器窗口)、COM 对象
(Excel app、打开的工作簿)既不能塞进模板,也不该被 trace 序列化。但下一个
step 经常就是需要**用同一个**窗口/工作簿,不是重新找一个。旧工具的解法是
`$Global:Shell` 这类全局变量 —— 这正是要消灭的东西。

`$Ctx.Session` 是运行期的**命名资源注册表**(一个 hashtable),规则:

1. **句柄 / COM 对象只进 Session,永不进 `outputs`。** `outputs` 必须
   JSON-可序列化(`ConvertTo-Json` 不报错、不丢字段)—— 这条由 P0-06 的
   契约检查器强制,也是 P0-R3 断点续跑"ledger 重放 outputs"的前提:
   outputs 进不了 ledger 的东西,重放就无从谈起。

2. **manifest 声明的是资源*种类*(kind),不是实例名。** `provides` 是
   本 step 能注册的资源种类数组,如 `provides = @('window')`(不是
   `@('mainWindow')`——`mainWindow` 是某一条工作流给它起的名字,manifest
   是通用的,不能替所有工作流预先决定叫什么)。消费方**不**在 `needs`
   里重复声明种类:一个 `type='session'` 的 `inputs` 参数,连同它必填的
   `sessionKind`,本身就是"这个 step 需要某种 Session 资源"的完整声明
   ——`needs` 里再写一遍 `session:<kind>` 是同一件事说两次,两处必然会
   漂移(某天改了 `sessionKind` 却忘了同步改 `needs`,或者反过来),
   和 P0-R4 修的 `columns.key`/`key.columns` 双写是同一类洞。**`needs`
   数组不出现 `session:<kind>` 这个形式**(P0-R10,详见 §4);runner
   / `ebi lint` 一律从 `inputs` 里的 `type='session'` 参数推导"这个
   step 消费哪个种类的资源"。**实例名永远由工作流决定**,manifest 里
   不出现任何具体名字。

3. **实例名怎么产生和引用**:
   - **`as` 是 runner 保留字段,不是 step 的 `inputs`。** 任何 `provides`
     非空的 step,它的调用点(不是 manifest!)都可以在 `with` 里加
     `as: "<名>"` 完成注册,如
     `{ "id": "ensure", "use": "browser.ensure", "with": { "as": "mainWindow" } }`
     ——`as` 不需要、也不允许出现在该 step 的 `$Manifest.inputs` 里,
     runner 在按 `inputs` 校验 `with` 之前就把它摘出来单独处理(§2 的
     "`with` 的参数按 `inputs` 校验"这条规则,`as` 是唯一的例外)。
     一个 step 调用最多注册一个资源,种类是其 `provides` 唯一的一项(本
     项目目前的 step 都只 `provides` 一种;需要注册多种资源的 step 请拆
     成多个 step,不要在一次调用里塞两个;`provides` 元素数 > 1 的 step
     不合法,P0-06 的契约检查器强制)。
   - **`provides` 非空但调用没写 `as` 是合法的**:资源正常产出、正常
     被这次调用内部使用,只是不注册进 `$Ctx.Session` 给后面的 step 引用
     ——适用于"这个窗口/工作簿只用这一次,不需要跨 step 复用"的场景。
     没有默认名这回事:不写 `as` 就是不注册,不是注册成某个隐含名字。
     **未注册的资源必须在产生它的这次 step 调用返回之前,由该 step 自己
     完成释放**(比如函数内部对临时 `New-Object -ComObject` 出来的对象
     做 `ReleaseComObject`)——它没有名字,后面没有任何 step 有机会引用
     到它,`$Ctx.Session` 之外没有第二条路能找回它(P0-R10)。**任何需要
     跨 step 存活、或者需要显式释放的资源都必须注册**:窗口句柄这类
     "泄漏了也无害"的资源可以不注册;但 COM 对象(Excel `Application`/
     `Workbook` 一类)"泄漏了会累积、需要人工杀进程",不注册就等于把
     自己锁死在无法释放的状态——这正是这条约束存在的原因。释放侧完整
     规则见下面第 5 点。
   - 消费方在自己的 `inputs` 里声明一个 `type = 'session'` 的参数,并带
     `sessionKind`(比如 `screen.capture_window` 的 `window` 参数是
     `@{ type='session'; sessionKind='window'; required=$true }`)。工作流
     调用时把这个参数的值填成某个 `with.as` 用过的名字字符串(**不走
     `{{}}` 模板** —— 这是给 runner 做资源查找用的字面量,不是数据)。
   - 名字的作用域是一次 `run`(`setup` 里注册的资源,`each`/`teardown`
     都能用)。
   - **同一个名字重复注册,如果它当前仍然"活着"(已注册、还没被释放),
     是运行期失败**(P0-R10 第四轮)。runner 在真正执行
     `with.as: "<名>"` 这次注册之前,先查 `$Ctx.Session` 里这个名字是否
     已经存在——存在就直接失败,不进 `Invoke-Step`(报"名字 `<名>` 已被
     占用,先释放再注册")。理由:静默覆盖等于静默泄漏——旧的那个实例
     (如果是 COM 对象)从此没有名字可以传给释放 step,谁也关不掉;这个
     项目的铁律是宁可停下来问人,不悄悄接受一个可能错的状态
     (`VOCABULARY.md` §1.2 的 `unknown`)。**不选"runner 自动释放旧的再
     注册新的"**——这是第 5 点已经否掉的"runner 自动清理"那条路,理由
     同样成立:不同种类资源的释放方式不同,runner 不该替工作流做这个
     决定。典型触发场景:`groupBy` 分组遍历里,`once:"group"` 在每组开头
     都调一次同一个 `with.as`(比如 `wb`)——不解决这条,第二组开始就会
     覆盖第一组还没关的工作簿,详见 `WORKFLOW-SCHEMA.md` §7.2 的
     `once:"groupEnd"`。

4. **`ebi lint` 怎么配平**:静态走一遍 `setup` → `each` → `teardown`(按
   写在 JSON 里的顺序,不模拟 `once`/`foreach` 的运行期分支),维护一张
   `名字 -> 种类` 表:
   - 遇到 `with.as: "<名>"` 且目标 step 的 `provides` 含种类 `K` → 表里
     记 `<名> -> K`
   - 遇到某个 `type='session'` 参数的字面量值 `<名>`,要求它已经在表里,
     且对应的种类等于该参数声明的 `sessionKind`,否则报错(要么"这个
     名字从没被 `with.as` 注册过",要么"注册的是别的种类的资源")
   这就是"用了 browser 资源但没人 `ensure` 过"的检测(P1-08),现在是
   **种类匹配 + 名字存在**两道检查,不是名字本身写死在 manifest 里比对。

5. **释放:显式调用 `releases` 非空的 step,不是 runner 自动清理
   (P0-R10)**。manifest 新增可选字段 `releases = @(<kind>, ...)`(与
   `provides` 对称,默认 `@()`),声明这个 step 会释放 `$Ctx.Session` 里
   哪个种类的资源——它照样通过自己 `inputs` 里某个 `type='session'` 的
   参数拿到具体实例名(和普通消费方用同一套机制,`releases` 只是多一条
   "这次调用等于释放"的元信息,不改变参数怎么传)。**释放点必须紧跟
   注册点所在的作用域,不能一律拖到 `teardown`**(P0-R10 第四轮):
   `setup` 里注册的资源在 `teardown` 释放;`each` 里注册的资源——典型是
   `once:"group"` 开的组级资源(比如每个交付物一个工作簿)——在**同组
   结束时**用新引入的 `once:"groupEnd"` 释放(`WORKFLOW-SCHEMA.md`
   §7.2)。拖到 `teardown` 才关是错的:`once:"group"` 每组都用同一个
   `with.as` 名字重新注册,第二组开始就会覆盖第一组还没关的实例(见上面
   第 3 点新增的"同名重复注册"约束),`teardown` 跑到时只剩最后一组的
   名字还查得到,前面几组全部泄漏。

   **不选"runner 在 run 结束时自动清空 Session"的理由**:不同种类资源
   的释放顺序/方式完全不同(工作簿要先 `Close` 再让 App `Quit`,窗口
   句柄什么都不用做)——让 runner 内置一套按种类分发的释放逻辑,等于
   让 runner 替 workflow 做"计算",违反 `WORKFLOW-SCHEMA.md` §0 的设计
   铁律("workflow JSON 只做连线,不做计算");这也是本仓库
   `ExcelHelpers.ps1` 一直以来的写法——`Close-Workbook` / `Close-ExcelApp`
   从来是调用方显式调用,从没有框架自动挡在中间。

   `ebi lint` 的配平检查(某种类要不要求释放调用,判据是下面第 6 点的
   `mustRelease` 种类表——**不是**"catalog 里现在有没有恰好带 `releases`
   的 step",见第四轮的修正)见 `WORKFLOW-SCHEMA.md` §9;`teardown` 每次
   resume 都重跑、释放 step 必须能安全面对"目标资源本来就不存在",见
   §6.2。

6. **哪些种类必须释放:`mustRelease`,种类表里显式声明,不从 catalog 里
   有没有对应 step 推出来(P0-R10 第四轮)。**

   | 种类 | `mustRelease` | 说明 |
   |------|---------------|------|
   | `window` | `$false` | 窗口句柄,泄漏无害,人工关掉即可(见 §1.1) |
   | `workbook` | `$true` | Excel `Workbook` COM 对象,泄漏会累积,需要人工杀 `EXCEL.EXE` |
   | `excelApp` | `$true` | Excel `Application` COM 对象,同上 |

   **表里的每个种类都必须有 step 产出它,不许有死条目**(P0-R10 第五轮)。
   这条不是形式主义:`excelApp` 和 `workbook` 分成两个种类,就意味着
   `excel.open` **不能**一个 step 同时产出 Application 和 Workbook——
   §3.4 第 3 点规定「一个 step 调用最多注册一个资源」「`provides` 元素数
   > 1 不合法」,而没被注册的那个又必须「在本次调用返回之前由 step 自己
   释放完」,可 Application 必须活到整个 run 结束,直接矛盾。所以 Excel
   的生命周期是**四个 step**:`excel.ensure_app`(`provides=@('excelApp')`)
   / `excel.open`(`provides=@('workbook')`,用 `type='session'` 输入消费
   app)/ `excel.close`(`releases=@('workbook')`)/ `excel.quit_app`
   (`releases=@('excelApp')`)。这不是为了迁就规则硬拆——`ExcelHelpers.ps1`
   现有代码本来就是 `New-ExcelApp` / `Open-Workbook` / `Close-Workbook` /
   `Close-ExcelApp` 四个独立函数,拆开反而更贴近要抄的代码;两者的作用域
   也天然不同:Application 一次 run 一个(`setup` 注册 / `teardown` 释放),
   Workbook 一组一个(`once:"group"` / `once:"groupEnd"`)。见 `BACKLOG.md`
   的 P4-01。

   **不选"某个种类在 catalog 里有没有 `releases` 覆盖"当判据的理由**:
   那是拿"当前 catalog 恰好长什么样"反推"这种资源要不要释放",依赖方向
   是反的——哪天有人往 catalog 里新增一个 `browser.close`
   (`releases = @('window')`),所有已经注册过 `window`、从来没写释放
   调用的老工作流会在同一天集体变红,而它们一行都没改。`mustRelease` 是
   资源**种类自身**的属性(泄漏会不会累积、需不需要人工清理),和 catalog
   里现在有没有恰好写出释放它的 step 无关,必须独立声明,不能派生。

   **放在这里(`STEP-CONTRACT.md` §3.4),不放 `catalog.json`**:
   `catalog.json` 是 `kernel/Docs.ps1`(P1-06)从 `modules/**` 的 manifest
   扫描**自动生成**的产物(见文件地图,标注"自动生成"),`mustRelease`
   恰恰不能从 manifest 扫描出来——`provides`/`releases` 只声明"哪个 step
   产出/消费/释放哪种资源",不声明"这种资源泄漏了要不要紧",从这两个
   字段反推 `mustRelease` 正是上一段要避免的循环。放进一份需要独立
   人工维护的表,和这份契约本身其它人工声明(比如 §2.1 `failures[].
   transient`)是同一类东西——不是从代码扫出来的,是写契约的人对这个
   世界的事实判断。**新增一个此前没出现过的资源种类时,连同它第一次
   出现的 `provides`/`releases` 一起,把它加进这张表**——这是 §7 底下
   元规则要求的"新增契约事实,定义处和汇总处一起改"的又一个例子;§7
   的 Run-Tests.ps1 清单有一条静态检查兜底这条纪律(见下方)。

`$Ctx.Session` 本身**不持久化**、**不写进 ledger**、**不出现在 trace 里**
(trace 只记 `outputs`)。它在每次进程启动时都是空的 —— 断点续跑时怎么
重建它,见 §6.2(P0-R3)。

---

## 4. `needs` — 前置条件

runner 在调用前检查,不满足直接失败,不进 step。

| need | 含义 |
|------|------|
| `foreground` | 需要一个前台窗口 |
| `browser` | 需要浏览器进程在运行 |
| `excel` | 需要 Excel COM 可用 |
| `worklist` | 需要工作清单已加载 |
| `calibrated:ocr` | 需要本机 OCR 校准通过(见 §5) |

`needs` 的意义是**把失败提前到运行前**,并且让 `ebi lint` 能在不运行的情况下
警告「这条工作流需要 Excel,你确定这台机器有吗」。

**`needs` 里没有 `session:<种类>` 这一项(P0-R10)。** 早先的草案里这个
形式和 §3.4/§2.2 的 `type='session'` + `sessionKind` 是同一件事的两处
声明——一个 step 需要哪种 Session 资源,只由它 `inputs` 里那个
`type='session'` 参数的 `sessionKind` 决定,不再额外写进 `needs`。理由
见 §3.4 点 2:两处声明会漂移(改了其中一处忘了改另一处,`ebi lint` 也
无从检查两处到底对不对得上,因为压根不存在第二份独立事实)。「用了这种
资源但没有任何 step 用 `with.as` 注册过」的检查(P1-08)因此**只**读
`inputs` 的 `type='session'`/`sessionKind`,不读 `needs`——具体算法见
§3.4 点 4。`releases` 非空的 step(§3.4 第 5 点,P0-R10)同样如此:它
消费哪个种类,一样只由它自己 `inputs` 里的 `sessionKind` 决定,`needs`
里不重复声明。

---

## 5. `tier` — 核心层与降级层

| tier | 含义 | 门禁 |
|------|------|------|
| `core` | 确定性的,跨机器结果一致 | 无 |
| `fallback` | **结果依赖本机环境**(OCR 引擎版本、字体、DPI、屏幕缩放) | 必须校准 |

### 5.1 为什么要这个区分

页面文本(`Ctrl+A`)是精确的、跨机器一致的。OCR 不是 —— 同一张图,不同机器的
OCR 引擎版本可能给出不同结果。

> **真实事故**:日文 OCR 把 MS Gothic 的 `9` 读成 `3`,而且读出来的是一个
> **格式完全合法**的时间戳,任何后置的格式检查都抓不住。围绕这个问题写的
> 一整套猜测/修复规则最后成了负债:规则本身可能把**正确**的读数改错。

### 5.2 规则

1. **文本永远优先。** 只要能拿到 `Ctrl+A` 文本,就绝不用 OCR。
   截图类 step 必须**同时**归档页面文本。
2. `tier = 'fallback'` 的 step 默认不可用。引用它的工作流,`ebi lint` 报警告。
3. 这类 step 的 `needs` 必须含 `calibrated:<something>`。
   runner 在运行前检查 `.ebi/calibration/<something>.json`:
   - 存在
   - 机器名匹配当前机器
   - 引擎版本匹配
   - 未过期(profile 声明有效期,默认 90 天)
   - 准确率达标(profile 声明阈值,默认 99%)

   任何一项不满足 → **拒绝运行**,提示跑 `ebi calibrate`。

### 5.3 降级层清单(初始)

`screen.ocr`、`screen.preprocess`、`screen.diff`、`verify.ocr_read`、`verify.pixel`

---

## 6. 幂等、checkpoint 与断点续跑(P0-R3)

「重跑按 ledger 跳过已完成的 (item, step)」这句话本身没说清楚一件事:跳过
`shot` 之后,下一个 `crop` 引用的 `{{steps.shot.out.path}}` 从哪来?下面
三条定死这个问题。

### 6.1 ledger 记录连同 outputs 一起持久化,resume 时重放

`idempotent = $true` 的 step,runner 在断点续跑时**可以**安全重新执行 ——
但绝大多数情况下**不需要真的重跑**:断点续跑的意义就是跳过已完成的
(item, step),而跳过之后,后续 step 如果引用了它的
`{{steps.<id>.out.<field>}}`,这个值必须能拿到,不能因为"跳过了"就变成
未定义引用。

**规则:ledger 每条记录里的 `outputs` 是完整的 step 返回值(不只是
`status`)。** resume 时,runner 按顺序把 ledger 里每条记录的 `outputs`
写回模板作用域(`steps.<id>.out.*`)—— 这叫**重放**。只有 ledger 里
**没有**记录的 (item, step) 才真正调用 `Invoke-Step`。

**一个例外:`provides` 或 `releases` 非空的 step 不适用这条跳过规则**
(P0-R10 第五轮)。它们注册/释放的是 `$Ctx.Session` 里的资源,而 Session
不持久化——跳过它们等于让资源在新进程里永远不再出现,后面引用这个名字的
step 会直接失败。完整规则见 §6.2。

ledger 记录形状(比旧版多了 `outputs`):

```json
{"runId":"...","item":"ABC123","step":"shot","status":"ok",
 "outputs":{"path":"capture/before_transferStatus/ABC123.png","width":1280,"height":800},
 "ts":"..."}
```

`step` 字段的值是这次调用的 `id`(`WORKFLOW-SCHEMA.md` §2.1 起 `id` 是
每个 step 调用的**必填**字段,不再是可选的)。这不是随口选的键 —— 同一段
里出现两次同一个 `use`(比如先后两次 `browser.tab_to`)完全合法,如果
ledger 拿 `use` 当键,第二次调用会和第一次共用一条记录,resume 时被误判
成"已经做过"而跳过。

这就是 §3.4 反复强调"`outputs` 必须 JSON-可序列化"的真正原因:句柄 / COM
对象进 `$Ctx.Session`、永不进 `outputs`,所以 `outputs` 天然能整体塞进
ledger 的一行 JSON 并原样读回。

### 6.2 `setup` / `teardown` 每次 resume 都重跑

`setup` 段负责建立 `$Ctx.Session`(浏览器窗口句柄、Excel COM……)——
这些资源本身不持久化(§3.4),进程重启后 Session 总是空的。所以
**`setup` 和 `teardown` 不走 §6.1 的跳过/重放逻辑,每次 resume 都完整
重跑一遍**,负责把 Session 重新建起来。

推论:**`setup` 里的每个 step 必须是 `idempotent = $true`。**
`browser.ensure` 重复执行的语义是"已经在前台就什么都不做,不是再开一个
浏览器窗口";`idempotent = $false` 的 step 不允许出现在 `setup` 里,
`ebi lint` 检查这一条。

**豁免范围不止这两段:任何 `provides` 或 `releases` 非空的 step,不论在
哪一段,resume 时都总是真执行(P0-R10 第五轮)。** 原来的写法只豁免
`setup`/`teardown`,漏掉了 `each` 里注册的资源——而 §3.4 第 5 点恰恰把
"每个交付物一个工作簿"(`once:"group"` 注册 + `once:"groupEnd"` 释放)
写成了推荐写法。不补这条,那种工作流一旦被 Ctrl+C 就**永久**跑不完:
ledger 里有 `(G1, open)` 的记录,resume 时按 §6.1 跳过,`wb` 从此不再进
Session,组里剩下的 item 一引用这个名字就失败,而 ledger 记录是永久的,
每次重试都撞同一堵墙。

`provides` 和 `releases` **必须成对豁免**,只豁免一边会更糟:如果只豁免
`provides`,resume 后 `open` 重新注册了 `wb`、`close` 却被跳过,`wb` 一直
占着名字,下一组的 `open` 撞上 §3.4 第 3 点的"同名重复注册非法"当场失败。

**豁免的作用点是"加载 ledger 的那一刻",不是"每次遇到这个 step"。**
这条极易实现错,单独说清楚:`once:"group"` 在**同一次运行内**也是靠 ledger
记录让组内后续 item 跳过的(§6.3)。如果把豁免理解成"运行期每次遇到都真
执行",组内第 2 条 item 就会再跑一次 `open`,当场撞上"同名重复注册非法"。
正确做法是——**runner 在进程启动、读 `ledger.jsonl` 构建"已完成集合"时,
把 `provides`/`releases` 非空的 step 的历史记录排除在外**;本次进程内这些
step 照常写新记录,`once:"group"` / `once:"groupEnd"` 的组内跳过语义完全
不受影响。一句话:**跨进程不继承,进程内照旧。**

推论二:**`provides`/`releases` 非空的 step 也必须 `idempotent = $true`**
(理由和上面对 `setup` 的要求同源:它们会被重复执行),`ebi lint` 和
P0-06 的契约检查器各查一侧(§7、`WORKFLOW-SCHEMA.md` §9)。

代价说明白,不假装没有:一个上次已经整组跑完的组,resume 时它的
`open`/`close` 会各自再执行一次(开一次、关一次,不写数据)。这是让
"注册-释放"在新进程里重新配平必须付的钱;不付这个代价的唯一替代是禁止在
`each` 里注册资源,那等于撤掉 `once:"groupEnd"` 和整个"每个交付物一个
工作簿"的写法。**这里选择付这个代价。**

**`teardown` 里负责释放的 step(`releases` 非空)同样每次 resume 都会
真的执行(P0-R10)**——包括"这次进程里这个名字根本没被注册过"的情况:
上一次运行是被 Ctrl+C 打断的(`WORKFLOW-SCHEMA.md` §1.1:硬中断不保证
跑 teardown),注册可能压根没发生;或者发生了,但这次重跑走的是另一次
`setup`,注册到的是另一个实例。这类释放 step 必须能安全处理"目标资源
不存在"(`$Ctx.Session` 里查不到这个名字,或者名字存在但底层句柄/COM
对象已经失效)——照抄本仓库 `ExcelHelpers.ps1` 现有的写法:每个释放
动作包一层 `try { ... } catch {}`,吞掉"本来就没有"这一类异常,不要让
`teardown` 因为清理一个不存在的资源而中止,连累后面该释放的资源也释放
不到。

### 6.3 `once: group` 的 ledger 键是 (group, step)

`"once": "group"` 的 step(`WORKFLOW-SCHEMA.md` §7.2)不是按 item 记账的
——它只在组内第一条 item 上真正跑一次。ledger 键因此是
**(group, step)**,不是 (item, step):

```json
{"runId":"...","group":"JOB_A","step":"navigate","status":"ok",
 "outputs":{"url":"https://.../JOB_A"},"ts":"..."}
```

resume 时,这条记录的 `outputs` 对**同组的全部 item** 重放,不只是当初
触发执行的那一条 —— 组里排在后面的 item 引用 `{{steps.navigate.out.X}}`
时,拿到的是这同一条记录的值。

`"once": "groupEnd"` 的 step(`WORKFLOW-SCHEMA.md` §7.2,P0-R10 第四轮)
——在同组**最后一条** item 处理完之后跑一次,和 `once: "group"` 的开头钩子
对偶——ledger 键**同样是** (group, step),不是 (item, step):跳过/重放的
规则和上面完全一致,只是触发时机在组尾而不是组头。

这两类 step 如果 `provides`/`releases` 非空(组级资源的注册/释放正是
典型),resume 时按 §6.2 的例外**总是真执行**,不按本节的 ledger 记录
跳过;ledger 键规则本身不变。

### 6.4 幂等的一般规则

`idempotent = $false` 的例子:`file.move`(源文件已经不在了)、
`browser.download_link`(会重复下载)。这类 step 必须紧跟一个
`flow.checkpoint`,让 runner 知道它已经做过了 —— checkpoint 走 §6.1 的
ledger + 重放规则,粒度默认 (item, step),`once: group` 时是 §6.3 的
(group, step)。

**完整的"中断发生在 `shot` 与 `crop` 之间"推演例子在
`WORKFLOW-SCHEMA.md` §7.5** —— 断点续跑是 runner 遍历 `source` 和
`flow.*` 构造的行为,例子里同时要用到 `once: group`,放在那份文档里
和 §7 的其余 `flow.*` 说明放在一起更合适。

---

## 7. 测试要求

| tier / effects | 要求 |
|----------------|------|
| `pure` | **必须**有单测,放 `Tests/Test-<Group>.ps1` |
| `read` | 应有单测(用临时目录 fixture) |
| `ui` / `write` / `destructive` | 只做**静态检查**(parse check + manifest 校验)。COM/Edge 路径在当前开发环境无法验证,必须在办公 PC 上冒烟确认 |

`Tests/Run-Tests.ps1` 额外强制:

- 每个 step 文件都能被 dot-source 且不含 `param()`
- `$Manifest.id` == 文件名
- `inputs` 里 `required` 和 `default` 不同时出现
- `failures` 非空,且每一项都有 `id` 和布尔类型的 `transient`(P0-R5)
- `example` 里用到的参数都在 `inputs` 里声明过(`as` 除外 —— 它是 runner
  保留字段,不进 `inputs`,§3.4)
- `outputs` 声明的类型都是 JSON-可序列化的(句柄 / COM 走 `$Ctx.Session`,不进 `outputs`,P0-R2)
- `inputs` 里 `type='session'` 的参数都带了 `sessionKind`(§2.2,P0-R2)
- `provides` 最多一项(§3.4 点 3 —— 一次调用最多注册一个资源;需要多个的 step 拆开写)
- `releases` 声明的种类,必须能在该 step 自己某个 `type='session'` 输入的
  `sessionKind` 里找到(否则声明了要释放,却没有输入能确定释放哪个实例)
  (§3.4 第 5 点,P0-R10)
- `provides`/`releases` 数组里出现的每个种类,都能在 §3.4 第 6 点的
  `mustRelease` 种类表里找到对应声明(否则是引入了一个新种类,却没有
  声明它泄漏了要不要紧)(§3.4 第 6 点,P0-R10 第四轮)
- `provides` 或 `releases` 非空的 step,`idempotent` 必须是 `$true`
  (resume 时它们总是真执行,见 §6.2,P0-R10 第五轮)
- step 文件里除 `Invoke-Step` 以外的每个函数,名字都以该 step id 的
  PascalCase 形式加连字符开头(`browser.find` → `BrowserFind-*`,
  `screen.capture_window` → `ScreenCaptureWindow-*`)。所有 step 会被同一
  runspace 依次 dot-source:`Invoke-Step` 撞名是设计好的,由 runner 逐个
  捕获(P1-02);裸名辅助函数(`Get-Row` 一类)则会互相覆盖,而且没有任何
  地方会报错 —— 后加载的那个静默赢
- 源码纯 ASCII

> 上面这份清单和 `WORKFLOW-SCHEMA.md` §9 的 `ebi lint` 清单是**同一类
> 汇总处的两份**——前者管 manifest 本身写得对不对(P0-06 的契约检查器,
> 静态检查单个 step 文件),后者管一条工作流引用 manifest 引用得对不对
> (`ebi lint`,静态检查 workflow JSON)。**新增任何一条 manifest 侧的
> 契约规则,这份清单要跟着改**,散落在 §2/§3.4 别处的规则不会被
> P0-06 自动捡起来。

---

## 8. 完整示例

```powershell
# modules/screen/screen.crop.ps1
# Crop a PNG by per-side pixel amounts. Pure image I/O, no window involved.

$Manifest = @{
  id         = 'screen.crop'
  group      = 'screen'
  summary    = 'Crop a PNG by per-side pixel amounts and write the result'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  idempotent = $true
  inputs     = @{
    path   = @{ type='path'; required=$true; desc='PNG to crop, modified in place unless out is given' }
    out    = @{ type='path'; desc='write here instead of overwriting path' }
    left   = @{ type='int'; default=0 }
    top    = @{ type='int'; default=0 }
    right  = @{ type='int'; default=0 }
    bottom = @{ type='int'; default=0 }
  }
  outputs    = @{
    path   = @{ type='path'; desc='the file that was written' }
    width  = @{ type='int' }
    height = @{ type='int' }
  }
  failures   = @(
    @{ id = 'file_not_found';     transient = $false }
    @{ id = 'crop_exceeds_image'; transient = $false }
    @{ id = 'image_read_error';   transient = $true  }
  )
  example    = @{
    use  = 'screen.crop'
    with = @{ path='{{steps.shot.out.path}}'; left=6; top=6; right=6; bottom=6 }
  }
  notes      = 'Replaces the four duplicated Invoke-CropPng copies in the old tool.'
}

function Invoke-Step {
    param($In, $Ctx)

    if (-not (Test-Path -LiteralPath $In.path)) {
        return @{ ok = $false; failure = 'file_not_found'; message = $In.path }
    }

    $dest = if ($In.out) { $In.out } else { $In.path }

    if ($Ctx.DryRun) {
        $Ctx.Log.Info(('would crop {0} by L{1} T{2} R{3} B{4} -> {5}' -f `
            $In.path, $In.left, $In.top, $In.right, $In.bottom, $dest))
        return @{ ok = $true; path = $dest; width = 0; height = 0 }
    }

    # ... 实际裁剪 ...

    return @{ ok = $true; path = $dest; width = $w; height = $h }
}
```

---

## 9. 加一个新 step 前的自问

1. 它的说明里有没有出现具体系统 / 画面 / 字段的名字?有 → 设计错了
2. 现有 step 组合起来能不能做到?能 → 不要加
3. 它是 `core` 还是 `fallback`?`fallback` 要有校准方案才能加
4. 它的 `failures` 穷举完了吗?
5. 它 dryrun 时的行为对吗?
