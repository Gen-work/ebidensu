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
  needs    = @('session:browser')  # 运行前置条件,见 §4
  provides = @()                 # 本 step 注册的 Session 资源,见 §3.4
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

  # --- 失败模式 ---
  failures = @('not_found', 'no_foreground_window')

  # --- 给 Agent 的示例 ---
  example  = @{ use='browser.find'; with=@{ term='{{item.key}}' } }
}
```

### 2.1 字段说明

| 字段 | 必填 | 说明 |
|------|------|------|
| `id` | ✓ | 全局唯一,`<group>.<verb>` 形式,必须与文件名一致 |
| `group` | ✓ | 见 CATALOG 的 8 组 |
| `summary` | ✓ | **一句话,英文,不超过 80 字符**。这是 Agent 挑 step 的主要依据 |
| `tier` | ✓ | `core` 或 `fallback`(见 §5) |
| `effects` | ✓ | 副作用等级,见 VOCABULARY §4 |
| `needs` | | 前置条件数组,见 §4 |
| `provides` | | 本 step 会注册的 Session 资源(如 `session:browser`),见 §3.4。`ebi lint` 用它做配平检查 |
| `idempotent` | ✓ | `$false` 的 step,runner 在续跑时不会自动重放 |
| `inputs` | ✓ | 参数定义。空则 `@{}` |
| `outputs` | ✓ | 返回字段定义。空则 `@{}` |
| `failures` | ✓ | **穷举**所有可能的失败标识。runner 用它校验返回值 |
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
| `any` | — | 尽量避免 |

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
- `failure` 的值**必须**在 `$Manifest.failures` 里,否则 runner 判为契约违反
- `message` 是给人看的一句话,英文,可选
- 失败时**也可以**带 outputs 字段(比如部分结果),runner 不会用,但会进 trace
- **outputs 的每个字段必须 JSON-可序列化**(§2.2 的类型即可)。窗口句柄、
  COM 对象一律走 `$Ctx.Session`(§3.4),永不放进返回值 —— outputs 要进
  trace,也要进 ledger 供断点续跑时重放(§6),塞不进 JSON 的东西放进来
  就是把这两条都弄坏
- **不许抛异常表达业务失败**。异常只用于「代码写错了」这类真正的意外;
  runner 捕获后统一记为 `failure = 'internal_error'`

### 3.2 `$Ctx` 提供什么(只读)

| 字段 | 说明 |
|------|------|
| `$Ctx.WorkDir` | 工作目录绝对路径 |
| `$Ctx.RunId` | 本次运行 id |
| `$Ctx.Profile` | 已加载并合并好的 profile(hashtable) |
| `$Ctx.Log` | `$Ctx.Log.Info('...')` / `.Warn(...)` / `.Debug(...)` |
| `$Ctx.DryRun` | `$true` 时,有副作用的 step **必须**只打印不执行 |
| `$Ctx.Session` | 运行期命名资源注册表(**可写**),见 §3.4。唯一允许跨 step 存活的进程内状态 |

**`$Ctx` 里没有别的 step 的输出。** step 之间只通过 workflow JSON 的模板引用
传值 —— 这是正交性的保证。(Session 装的是**资源**,不是输出。)

### 3.3 `DryRun` 的处理

`effects` 为 `write` / `destructive` / `ui` 的 step **必须**处理 `DryRun`:

```powershell
if ($Ctx.DryRun) {
    $Ctx.Log.Info(('would save to {0}' -f $In.path))
    return @{ ok = $true; path = $In.path }
}
```

`pure` / `read` 的 step 可以照常执行。

### 3.4 `$Ctx.Session` — 会话资源(评审修订 P0-R2)

窗口句柄、COM 对象这类**进程内资源**进不了 JSON:不能进模板、不能进 trace、
不能进 ledger(断点续跑要重放 outputs,见 §6)。而 `steps.X.out` 的引用又
**只限同段** —— `setup` 里 `browser.ensure` 拿到的句柄,`each` 里的
`screen.capture_window` 根本引用不到。不给正道,实现者只能回去用全局变量,
也就是旧工具的 `$Global:Shell`。所以:

- `$Ctx.Session` 是一个命名资源注册表:
  `$Ctx.Session['browser'] = @{ hwnd = ...; pid = ... }`
- **句柄 / COM 对象只进 Session,永不出现在 outputs 里**(§3.1;
  契约检查器强制,见 §7)
- 注册资源的 step 声明 `provides = @('session:browser')`;消费的 step 声明
  `needs = @('session:browser')`。`ebi lint` 静态检查:工作流里每个 `needs`
  在它之前(setup 算在前)都有对应的 `provides`
- 同类资源可以有多个实例:`excel.open` 用 `with.as = 'wb1'` 注册
  `session:workbook:wb1`,后续 step 用同一个名字引用
- **Session 不持久化**。断点续跑时它是空的 —— 所以 `setup` 在每次 resume
  都重跑,负责重建资源(见 §6)
- 释放责任:注册了需要释放的资源(COM)的工作流,`teardown` 里必须有对应的
  释放 step;run 结束(含失败退出)时 Session 里仍存活的可释放资源,
  runner 打 `[WARN]`

**为什么不直接用全局变量**:`$Global:Shell` / `$Global:Timing` 就是这样长出来
的 —— 谁都能读写、没人声明依赖、断点续跑后凭空消失。Session 把同一件事变成
**声明过、可 lint、resume 时被系统性重建**的。

### 3.5 标准候选形状(评审修订 P0-R4)

凡是「按 key 找东西可能不唯一」的 step(`table.key`、`file.find`、`file.newest`、
`verify.match_record`、`excel.find_anchor`……)在歧义时**必须**返回同一个形状,
`human.choose` 只有一个渲染器:

```powershell
@{ ok = $false; failure = 'ambiguous'
   candidates = @(
     @{ value    = 'ABC123.260824.10515511.dat'  # 候选本体(文件名/行号/单元格地址)
        source   = 'download-dir'                # 从哪找到的
        evidence = @{ createdAt='09:51:02'; receivedAt='09:50:03'; size='1.2 MB' } }
     # ... 全部候选,一个不少
   )
   suggestion = @{ index  = 2
                   reason = 'received time is newest inside the run window'
                   doubts = '#1 and #2 are 3 minutes apart; both may belong to this run' }
}
```

规则:

- `candidates` 是**全部**候选,不许只给「最佳」
- `evidence` 的键因场景而异,但**有什么证据都要放**;证据不够就让
  `suggestion.doubts` 说出来,不要假装能判断
- `suggestion` 可以是 `$null`(真的没法建议),但不许假装确定
- 这个形状由 `kernel/Key.ps1` 的排序 / 证据函数产生,step 不自己拼 ——
  旧工具七处各写 `-eq` 的教训(PROFILE-SCHEMA §6.4)

---

## 4. `needs` — 前置条件

runner 在调用前检查,不满足直接失败,不进 step。

| need | 含义 |
|------|------|
| `foreground` | 需要一个前台窗口 |
| `session:<name>` | 需要名为 `<name>` 的 Session 资源已被前置 step 注册(如 `session:browser`、`session:workbook:wb1`),见 §3.4 |
| `excel` | 需要 Excel COM 可用(能创建实例;已打开的工作簿用 `session:workbook:*`) |
| `worklist` | 需要工作清单已加载 |
| `calibrated:ocr` | 需要本机 OCR 校准通过(见 §5) |

`needs` 的意义是**把失败提前到运行前**,并且让 `ebi lint` 能在不运行的情况下
警告「这条工作流需要 Excel,你确定这台机器有吗」。

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

## 6. 幂等、ledger 与断点续跑(评审修订 P0-R3)

ledger 是断点续跑的唯一依据,写 `run/<runId>/ledger.jsonl`。
粒度是 **(item, step)**,不是 (item);`once: "group"` 的 step 是 **(group, step)**。
item 字段用 `keySafe`(PROFILE-SCHEMA §6.1b)。

每条记录**连同 outputs 一起持久化**:

```json
{"runId":"...","item":"ABC123__JOB_A","step":"shot","status":"ok",
 "out":{"path":"capture/before_transferStatus/ABC123__JOB_A.png"},"ts":"..."}
```

这是 outputs 必须 JSON-可序列化(§3.1 / §3.4)的第二个原因。

### 6.1 resume 的四条规则

1. **`setup` / `teardown` 每次 resume 都重跑**(Session 是空的,setup 负责重建
   资源,§3.4)→ setup 里的 step 必须幂等
2. `each` 里已完成的 (item, step) 按 ledger **跳过**,其 outputs 从 ledger
   **重放**进模板作用域 —— 后续 step 的 `{{steps.X.out.Y}}` 照常解析,
   不需要真的重新执行 X
3. `once: "group"` 的 step:组内任何 item 上 resume,其 outputs 同样重放
4. `idempotent = $false` 的 step(`file.move`、`browser.download_link`)规则
   相同 —— ledger 里有就跳过;**没有就会重新执行**,所以这类 step 必须紧跟
   `flow.checkpoint`,把「已经做过」尽快落进 ledger

### 6.2 推演例子(中断发生在 shot 与 crop 之间)

item `ABC123__JOB_A` 已完成 `page_text`(留档 txt)和 `shot`(存了 PNG),
操作员 Ctrl+C,晚些时候重跑同一条 workflow:

1. `setup` 重跑:`browser.ensure` 重新注册 `session:browser`(旧句柄早已失效)
2. 前面已 checkpoint 的 item 被 `source.select` 直接过滤(字段已是 `ok`)
3. 轮到 `ABC123__JOB_A`:`page_text`、`shot` 在 ledger 里 → 跳过,不再打键盘、
   不再截图;它们的 outputs(含 `steps.shot.out.path`)从 ledger 重放
4. `crop` 是这个 item 第一个真正执行的 step,`{{steps.shot.out.path}}`
   解析到重放的路径,照常裁剪
5. 之后照常直到 `flow.checkpoint`

**没有这套规则时的两种典型返工**:重放缺失 → crop 解析不到 shot 的输出,只能
整 item 重跑(重复截图,P2-06 验收直接不过);Session 不重建 → 跳过 setup 的
「优化」让 capture 拿着失效句柄截黑屏。

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
- `failures` 非空
- `example` 里用到的参数都在 `inputs` 里声明过
- 源码纯 ASCII
- `outputs` 声明的类型都是 §2.2 的可序列化类型(句柄 / COM 走 §3.4 的 Session)
- 除 `Invoke-Step` 外的辅助函数**必须带 step 前缀**(如 `BrowserFind-*`)——
  所有 step 被同一 runspace 依次 dot-source:`Invoke-Step` 由 Registry 在每次
  dot-source 后**立刻捕获** `${function:Invoke-Step}` 进按 id 索引的表来解决
  同名覆盖;裸名辅助函数则会互相覆盖且无人发现

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
  failures   = @('file_not_found', 'crop_exceeds_image', 'image_read_error')
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
