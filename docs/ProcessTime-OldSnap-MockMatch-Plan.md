# ProcessTime 旧快照 9→3 核对 —— Mock 渲染 + 像素匹配方案

状态：**PARKED（2026-08-04 暂缓）**。设计保留，未实现，不在当前 TODO 内 ——
参见 `docs/Parked-Ideas.md`（为什么暂缓 + 重新拾起时的步骤）。
暂缓原因简述：两代 D2 都需要办公室 PC 现场标定才能启用，而 3↔9 的日常问题
已由确定性检查覆盖（`TimeDigitVerify.ps1`：不可能位修复 + 起止/页面时长的
算术消歧 + 条件格式标红人工复核），无需图像比对。

以下为原设计全文（给将来接手的 agent 的交接说明）。
评审：见文末 §10（2026-07-27 复核结论 + 落地顺序调整）。
先读 `CLAUDE.md`（本文默认遵守其全部约定）与
`docs/ProcessTime-OldSnap-Verify-Plan.md`（上一版方案，D1 与确定性核对已实现并合入 v2.17.0）。

本文**只替换 D2（图像核对）那一层**。D1（相関ID→快照超链接）、`検証` 列、
以及 §4.3 的确定性核对（时长算术 + 3↔9 datestamp 交叉校验）已经上线，不动。

---

## 1. 为什么换方案

v2.17.0 的 D2 用 GDI+ 渲染 MS Gothic 的 `3`/`9` 逐位模板，再和快照裁剪出的
数字做比对。两个硬伤：

1. **光栅化引擎不一致。** 真实快照来自 Edge（DirectWrite），参考图来自 GDI+
   （`Graphics.DrawString`），抗锯齿方式天然不同 —— 原计划 §4.2 自己就写了
   备注："GDI+ rasterises differently from the Edge snap ... or render the
   reference via a headless Edge/WebBrowser instead"。
2. **需要逐位像素坐标标定**，每个数字一个矩形，脆弱且标定量大。

新方案把这两点一次性消掉：

> **参考图不再用 GDI+ 画，而是用办公室 PC 上的 Edge 渲染 mock page 得到** ——
> 同一个浏览器引擎、同一个 MS Gothic 字体、同一套 CSS（`mock-page/templates/
> hm-batch-status.html` 已按真实 IGPXA041 页面校准过）。然后整块做模板匹配，
> 不再逐位切。

### 硬约束

**办公室 PC 不能用 npm，大概率也不能用 node。** 所以整条链路必须是
**纯 PowerShell + Edge**。`mock-page/gen.mjs`、`mock-page/pixeldiff.mjs` 保持
原样，但它们**只在 CI / 开发机跑**，办公室 PC 的流程一行 node 都不碰。
（HTML 填充逻辑要在 PowerShell 里重写一份，见 §4.2；`gen.mjs` 的
`buildRowsHtml` 是现成的参考实现。）

---

## 2. 总体思路（三步）

对每个需要核对的行（某 correl 的 GIFT 或 GFIX 侧）：

```
Step 1  校准（一次性，鼠标点击）
        用一组 dummy data 生成 dummy page → 渲染 → 人工点击确认
        绿底行(G) / 白底行(W) 各字段的像素范围

Step 2  生成参考图
        把该行 OCR 出来的值填进 mock page，同一组数据渲染两遍：
        第 1 行绿底(BG_G)、第 2 行白底(BG_W) → 裁出两张参考图

Step 3  像素匹配
        拿 BG_G 和 BG_W 两张参考图，直接去 source snap
        (snap\<GIFT|GFIX>_HM\<correl>.png) 上找像素匹配
```

**判定规则**：

| 结果 | 含义 | `検証` 列 |
|---|---|---|
| BG_G 命中 **或** BG_W 命中 | OCR 值与图像一致 | `OCR-OK`（还需同时通过确定性核对） |
| **G 和 W 都命不中** | 无法确认 | `要確認`（人工） |

这条规则天然抓 9→3：如果 OCR 读成 `03:53:07`，渲染出来的 `53` 在快照上找不到
（快照上是 `59`），两个背景都命不中 → 判 `要確認`。保守方向正确：宁可多标一个
让人瞄一眼，也绝不能把错值自动放行。

### 为什么要两个背景色

`mock-page/templates/hm-batch-status.html` 里（照抄真实 ids.css）：

```css
tr.unevenrow td { background-color: #eeffee; }  /* 浅绿，真实页面第 1 个数据行 */
tr.evenrow   td { background-color: #ffffff; }  /* 白 */
```

真实快照里目标 correl 落在奇数行还是偶数行是不确定的，所以同一组数据必须
渲染两遍，得到 `BG_G` / `BG_W` 两张参考图，分别去匹配。

---

## 3. 与现有代码的关系

### 复用（不改）

| 组件 | 用途 |
|---|---|
| `Locate-ByImage.ps1` | LockBits 模板匹配。`-SourcePath -TemplatePath -Tolerance -Quiet`，返回 `X/Y/Width/Height` 或 `$null`。**有 `param()` → 必须用 `&` 调用，禁止 dot-source** |
| `OldSnapVerify.ps1` 的 `Get-OldSnapVerifyVerdict` | 判定层**完全不用改**：它已经接受 `-PixelResult 'ok'/'ng'/''` 和 `-PixelEnabled`。新方案只是换了 `PixelResult` 的产生方式 |
| `mock-page/templates/hm-batch-status.html` | 已按真实页面校准的模板（列宽 150/150/80/80/40/80/130/70/70/100，`{{TITLE}}/{{META}}/{{LISTCLASS}}/{{ROWS}}` 占位符） |
| `Calibrate-HmGeometry.ps1` | 校准 UI 的**样板**：WinForms + PictureBox + 点击收集 + 结果写剪贴板。新校准脚本照抄这个结构 |
| `ScreenRegion.ps1` | 纯裁剪坐标运算 |
| `PixelDigitMatch.ps1` 的 `Get-DigitNcc` | **备用打分器**：如果 `Locate-ByImage` 的容差匹配对抗锯齿太脆，改用 NCC 打分 |

### 封存（不删，标注 superseded）

- `OldSnapPixelVerify.ps1`（GDI+ 逐位模板）—— 在文件头注明被本方案取代，
  `ProcessTime.OldSnapVerify.PixelDiff.Enabled` 保持 `$false`。
  不要同时存在两条互相竞争的 D2 路径。

### 新增

| 文件 | 有无 `param()` | 说明 |
|---|---|---|
| `MockPageBuild.ps1` | **无** → 可 dot-source | **纯逻辑**：HTML 行拼装、3↔9 变体生成、几何推导。进单测 |
| `Render-MockPage.ps1` | 有 → 用 `&` 调用 | Edge headless 渲染 HTML → PNG。静态检查 |
| `Calibrate-MockRowGeometry.ps1` | 有 → 用 `&` 调用 | WinForms 点击标定。静态检查 |
| `OldSnapMockMatch.ps1` | **无** → 可 dot-source | 裁剪 + 调 `Locate-ByImage` + 汇总成 `PixelResult`。COM/GDI 部分静态检查 |
| `Tests\Test-MockPageBuild.ps1` | — | `MockPageBuild.ps1` 的单测 |

---

## 4. 各阶段详细设计

### 4.1 阶段 A：校准（`Calibrate-MockRowGeometry.ps1`）

**目的**：确定在 **mock 渲染出的 PNG** 上，绿底行和白底行各字段的像素矩形。

**流程**：

1. 用一组固定的 dummy data 生成 dummy page（2 行，**同一组数据**，第 1 行
   `unevenrow` 绿底、第 2 行 `evenrow` 白底）。dummy 值用好认的哨兵值，
   例如 `2026/07/23 03:59:07` / `2026/07/23 03:59:52` / `00:00:45`。
2. 渲染成 PNG（§4.3）。
3. WinForms 窗口显示这张 PNG，操作员**鼠标点击**确认像素范围：
   - 点 1、2：绿底行「開始日時」单元格的**左上角**和**右下角**
   - 点 3、4：白底行「開始日時」单元格的**左上角**和**右下角**
   - → 4 次点击即可。「終了日時」「処理時間」由模板已知的列宽比例
     （150 / 150 / 80 CSS px）按实测缩放系数推导出来
     （`scale = 实测開始列宽px / 150`）。
4. 窗口上**叠加画出推导结果的矩形**让操作员目视确认，不对就重点。
5. 结果存盘（§5 的 `MockMatch.Geometry`），并同时写剪贴板
   （照抄 `Calibrate-HmGeometry.ps1` 的做法）。

> **纯逻辑部分**（进单测）：由 4 个点击点 + 列宽表 + 缩放系数推导出全部字段
> 矩形的那段运算，放进 `MockPageBuild.ps1`，例如
> `Resolve-MockFieldRects -ClickTL -ClickBR -ColumnWidths -Scale`。

#### ⚠️ 校准里最关键的一步：确认缩放一致

模板匹配**不是尺度不变的**。mock 渲染出来的字号像素尺寸必须和真实快照
**一致**，否则永远匹配不上。所以校准阶段必须包含这一步验证：

1. 挑一条**已知 OCR 值正确**的真实旧快照。
2. 用它的值走一遍 §4.2 + §4.4，看能不能命中。
3. 命不中 → 调 `RenderScale`（Edge 的 device scale factor）/ 窗口宽度，重试，
   直到命中。命中时的 `RenderScale` 就是要写进配置的值。

**这一步不通过，后面全部无意义。** 建议把它做成校准脚本的第二段
（`-VerifyAgainst <真实快照路径>`），而不是留给操作员手工判断。

### 4.2 阶段 B：无 node 填充 HTML（`MockPageBuild.ps1`，纯）

把 `gen.mjs` 的 `buildRowsHtml` 用 PowerShell 重写：

```powershell
# 纯字符串拼装，无 I/O、无 COM。可 dot-source，可单测。
New-MockRowHtml   -Row <hashtable> -RowIndex <int>   # 按 index 奇偶决定 unevenrow/evenrow
New-MockPageHtml  -Template <string> -Rows <object[]> -Title -Meta -ListClass
```

要点：
- 替换 `{{TITLE}}` / `{{META}}` / `{{LISTCLASS}}` / `{{ROWS}}`，注意 `gen.mjs`
  用的是 `replaceAll`（模板注释里也提到了这些 token，单次 replace 会替换错
  地方）—— PowerShell 用 `-replace` 或 `.Replace()` 都是全替换，没问题。
- HTML 转义 `& < >`。
- `データ作成日` 为空时填全角空格（真实页面就是这样，`gen.mjs` 的 `FW_SPACE`）。
- **源码保持 ASCII**：日文字面量（`正常終了` 等）一律走 `ProjectLabels.ps1`
  或本文件内的 `[char]` 常量，不要在 `.ps1` 里写生日文
  （`Check-Encoding.ps1` 会报警）。

**批量优化（建议实现）**：一次渲染可以带很多行。对 N 个 correl，发 2N 行：
第 `2i` 行放 correl i（落在绿底）、第 `2i+1` 行放同一个 correl i（落在白底）。
因为 `unevenrow`/`evenrow` 本来就是按行序交替的，这样一次渲染就同时拿到了
所有 correl 的 G/W 两个版本 —— 比"每个 correl 渲染一次"快一个数量级。

### 4.3 渲染（`Render-MockPage.ps1`，静态检查）

Edge 是 Chromium 内核，支持无头截图：

```powershell
& $EdgePath --headless=new --disable-gpu `
            --force-device-scale-factor=$RenderScale `
            --window-size=$W,$H `
            --screenshot="$OutPng" `
            "file:///$HtmlPath"
```

- `$EdgePath` 默认 `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`，
  做成配置项。
- **备用方案**（如果公司策略禁掉 headless）：正常打开 Edge 显示该 HTML，
  用现有 `Common.ps1` 的窗口截屏助手抓图 —— 这条路项目里到处都在用，稳。
  实现时把渲染封装成一个函数，两种后端可切换。
- 渲染完必须确认：**字体真的是 MS Gothic**。模板里写的是
  `"ＭＳ ゴシック", "MS Gothic", monospace`，办公室 PC 上应当命中第一个。
  Linux/CI 上会 fallback 到别的字体 —— 所以这一步只能在办公室 PC 验证。

### 4.4 阶段 C：裁剪 + 匹配 + 判定（`OldSnapMockMatch.ps1`）

```powershell
# 1. 从渲染 PNG 上按标定矩形裁出参考图
#    输出到 snap\_mockref\<correl>\<side>_<field>_G.png / _W.png
# 2. 逐字段匹配
foreach ($field in 'Start','End','Duration') {
    foreach ($bg in 'G','W') {
        $box = & $LocateByImage -SourcePath $snapPng -TemplatePath $refPng -Tolerance $T -Quiet
        if ($box) { $hit = $true; break }
    }
}
# 3. 汇总 -> PixelResult
```

**汇总规则**（保守）：

- 全部字段都至少命中一个背景 → `'ok'`
- 有任何字段两个背景都命不中 → `''`（unknown）

> 注意：这里**只产出 `'ok'` 和 `''`，不产出 `'ng'`**。因为"没找到"只能证明
> "无法确认"，不能证明"OCR 一定错"。而 `Get-OldSnapVerifyVerdict` 里
> `'ng'` 和 `''` 在 `PixelEnabled` 时都会落到 `NeedsCheck`，所以行为正确，
> **判定层一行都不用改**。

### 4.5 可选增强（第二期，强烈建议做）：3↔9 反转变体

把"命不中 → 人工"升级成"确定就是 9→3 误读"：

对 OCR 值里每个 `3`/`9`，生成把它翻转后的候选值（`03:53:07` → `03:59:07`），
一并渲染成参考图去匹配：

- 原值命中 → OCR 正确
- **原值命不中、但某个翻转变体命中** → **确定是 3↔9 误读，且知道正确值**
  → 可以直接在 `検証` 列写明"疑似 9→3，正确值应为 xx:xx:xx"，
  甚至自动修正（建议先只提示，不自动改）
- 都命不中 → 人工

变体生成函数是**纯逻辑**（`New-DigitSwapVariants -Text '03:53:07'`），进单测。
这一步几乎是白送的，因为渲染是批量的，多几行成本很低。

---

## 5. 配置项

挂在已有的 `ProcessTime.OldSnapVerify` 下（它已在 `ProcessTime` 组里，
所以 **`Get-ConfigOverlayGroups` 不用加新组**，但 `Get-ConfigOverlayReadmeText`
的说明要更新 —— 见 `CLAUDE.md` 的"Config overlay groups must track
VerifyConfig.psd1"一节）：

```powershell
OldSnapVerify = @{
    # ... 现有的 Enabled / EmitHyperlink / EmitVerifyColumn / SnapDirPattern ...

    PixelDiff = @{ Enabled = $false }   # v2.17.0 的 GDI+ 逐位方案，封存

    MockMatch = @{
        Enabled     = $false            # 标定完成前保持关闭
        EdgePath    = ''                # 空 -> 用默认安装路径
        RenderMode  = 'Headless'        # 'Headless' | 'Window'（备用截屏）
        RenderScale = 1.0               # device scale factor，由校准确定
        Tolerance   = 15                # Locate-ByImage 的容差（ClearType 余量）
        TemplatePath = ''               # 空 -> mock-page\templates\hm-batch-status.html
        RefDir      = 'snap\_mockref'   # 参考图输出目录
        SwapVariants = $false           # §4.5 的 3<->9 反转变体
        Geometry    = @{                # 由 Calibrate-MockRowGeometry.ps1 写入
            Scale = 1.0
            G = @{ Start = @{X=0;Y=0;W=0;H=0}; End = @{}; Duration = @{} }
            W = @{ Start = @{}; End = @{}; Duration = @{} }
        }
    }
}
```

同时把 `Write-ProcessTimeWorkbook` 里现在传给 `Get-OldSnapRowPixelVerdict`
的那一段，改成走新的 `OldSnapMockMatch.ps1`（`PixelEnabled` 改由
`MockMatch.Enabled` 驱动）。

---

## 6. 要遵守的项目约定（`CLAUDE.md`）

- `.ps1` **纯 ASCII 源码**；日文一律 `[char]` 或 `ProjectLabels.ps1`。
  改完跑 `Check-Encoding.ps1`。
- 只有**没有 `param()`** 的文件才能 dot-source。`Locate-ByImage.ps1`、
  `Render-MockPage.ps1`、`Calibrate-MockRowGeometry.ps1` 都有 `param()`，
  必须 `& $path @args`。
- dot-source 前先 `$forceFlag = [bool]$Force.IsPresent`。
- 纯逻辑进 `Tests\`，CI 里跑；COM / GDI / Edge / WinForms 只做静态检查，
  代码里明确标注，交给办公室 PC 验证。
- 目标运行时是 **Windows PowerShell 5.1**。注意已知的 PS 5.1 坑：
  - 不要写 `@($hashtable[$key])` 包 `List[object]`（会抛
    "Argument types do not match"）—— 用
    `ConvertTo-ProcessTimeBucketArray` 那种显式 helper。
  - `[datetime]::TryParseExact` 传格式数组时必须显式 `[string[]]` 强转，
    否则会选错重载（v2.17.0 踩过）。
  - 构造 Windows 绝对路径用 `[System.IO.Path]::Combine`，不要用 `Join-Path`
    （非 Windows CI 上 `Join-Path` 会做盘符解析而抛错）。
- 改完跑 `Tests\Run-Tests.ps1`。注意 `Test-EvidencePlan.ps1` 有 **2 个只在
  Linux 上失败**的路径分隔符用例（`\` vs `/`），在 Windows 上会通过，属既有
  情况，不是回归。

---

## 7. 测试

### CI（Linux，纯逻辑）

- `MockPageBuild.ps1`：行 HTML 拼装（奇偶 → unevenrow/evenrow）、HTML 转义、
  空 datestamp 填全角空格、占位符全替换、`New-DigitSwapVariants`、
  `Resolve-MockFieldRects` 的几何推导。
- 匹配结果 → `PixelResult` 的汇总规则（全命中→`ok`；任一字段全不中→`''`）。

### 办公室 PC（必须）

1. 跑校准，确认 4 次点击 + 叠加矩形正确。
2. **`-VerifyAgainst` 一条已知正确的真实快照 → 必须命中**（缩放对齐验证，
   见 §4.1 的警告）。
3. 一条**已知 9→3** 的行 → `検証` 必须是 `要確認`，**绝不能** `OCR-OK`。
4. 一条**已知正确**的行 → `検証` 应为 `OCR-OK`。
5. （做了 §4.5 的话）已知 9→3 的行应能报出"疑似 9→3，正确值 xx:xx:xx"。
6. 确认渲染出的字体确实是 MS Gothic。

---

## 8. 风险与未决问题

| # | 风险 | 缓解 |
|---|---|---|
| 1 | **缩放/DPI 不一致**（最大风险）—— 模板匹配非尺度不变 | §4.1 的 `-VerifyAgainst` 强制验证；`RenderScale` 可调 |
| 2 | 抗锯齿 / ClearType 子像素差异导致精确匹配失败 | `Tolerance`（默认 15）；仍不行改用 `Get-DigitNcc` 打分 |
| 3 | 真实快照可能有 Ctrl+F 橙色高亮行（第三种背景） | 先只做 G/W，命不中→人工（保守，可接受）；后续可加第三个变体 |
| 4 | 旧快照本身分辨率太低（这本来就是 bug 根源），可能两个背景都命不中 | 落到人工，符合保守原则；不要为了提高命中率放宽 Tolerance |
| 5 | 公司策略可能禁用 Edge headless | 备用：正常窗口 + `Common.ps1` 截屏（§4.3） |
| 6 | 只有嵌在 Excel 里、没有独立 PNG 的快照 | 沿用上一版结论：这类行**只给超链接**，一律 `要確認`，不参与自动确认 |

---

## 9. 建议实现顺序

1. `MockPageBuild.ps1`（纯）+ `Tests\Test-MockPageBuild.ps1` —— CI 可验证，先做。
2. `Render-MockPage.ps1` —— 办公室 PC 上先手工跑通一次，肉眼看渲染像不像真页面。
3. `Calibrate-MockRowGeometry.ps1` + `-VerifyAgainst` 缩放验证 —— **卡点，过不了就停**。
4. `OldSnapMockMatch.ps1` 接进 `Write-ProcessTimeWorkbook`，`MockMatch.Enabled` 打开。
5. （可选）§4.5 的 3↔9 反转变体。

---

## 10. 复核结论（2026-07-27）

复核了 §1–§9 与仓库现状（`OldSnapVerify.ps1` / `OldSnapPixelVerify.ps1` /
`Locate-ByImage.ps1` / `mock-page/templates/hm-batch-status.html`）。

### 10.1 结论：方案成立，可以照原顺序做

三条关键假设都在代码里核对过，成立：

1. **判定层确实不用改。** `Get-OldSnapVerifyVerdict` 的签名是
   `-PixelResult 'ok'/'ng'/''` + `-PixelEnabled`，且 `PixelEnabled` 时
   `$PixelResult -ne 'ok'` 一律 `NeedsCheck` —— 所以新匹配器只产 `'ok'`/`''`
   完全够用，`'ng'` 不需要（§4.4 的判断正确）。
2. **`Locate-ByImage.ps1` 有 `param()`**，必须 `& $path @args`，不能
   dot-source（§3 已写明，别忘）。
3. **模板已按真实页面校准**（列宽 150/150/80/80/40/80/130/70/70/100 +
   `unevenrow`/`evenrow` 两种底色），§4.2 的两遍渲染有依据。

### 10.2 §4.1 的缩放验证要做成硬性卡点

模板匹配不是尺度不变的，这是本方案唯一会"全盘失败"的风险。建议把
`-VerifyAgainst` 做成**退出码不同**的独立步骤（命中 = 0，未命中 = 非 0），
而不是打印一行让人自己看 —— 这样第 3 步没过就不会有人接着写第 4 步。

配套建议：`Calibrate-MockRowGeometry.ps1` 在写入 `Geometry` 的同时，把
**验证用的真实快照路径 + 通过时的 `RenderScale`** 一并写进配置，作为标定
"何时、针对哪张图通过"的凭据。窗口尺寸一变（`WindowWidth/Height` 或 DPI），
这份标定就作废，配置里留着来源才看得出来。

### 10.3 §4.5（3↔9 反转变体）建议提前到第一期

原文把它列为"第二期，强烈建议做"。复核后建议**直接放进第一期**，因为：

- 成本几乎为零：渲染本来就是批量的，变体只是多几行 HTML。
- 收益是质变：从"命不中 → 人工看"变成"确定是 9→3 且知道正确值"。
- **它还顺带解决一个单独的问题**：v2.17.0 的
  `Repair-ProcessTimeStartFromStamp` 只能在有干净 14 位 datestamp 时纠正
  `hh:mm`，秒位无能为力；反转变体匹配对秒位同样有效。

但仍然坚持"只提示、不自动改"：在 `検証` 列写"疑似 9→3，正确值 xx:xx:xx"，
让人一眼确认后手工改。自动改值需要另一轮回归，不要混进来。

### 10.4 与 v2.18.0 的衔接

v2.18.0 给 D1 加了**回退图片**（没有独立 snap PNG 时，超链接指向本阶段从
证据簿导出的图片，`snap\ProcessTime\<correl>\<SIDE>_<correl>_NN.png`），
并让 `検証` 的 `SnapExists` 接受任一图片。这对本方案的影响：

- **像素匹配只能用真实 snap PNG**，不能用回退图片 —— 回退图片是证据簿里
  被重新缩放过的副本，尺寸与 mock 渲染对不上（正是 §10.2 的风险）。
- 因此 `OldSnapMockMatch.ps1` 接进 `Write-ProcessTimeWorkbook` 时，
  **沿用现在 `PixelEnabled` 那一段的 `$snapExists` 门（不是
  `$imageExists`）**。ProcessTime.ps1 里已经这样写了，注释也标了原因。
  结果就是：只有证据簿图片的行 → 无像素结果 → `要確認`，正是 §8 风险 6
  要的行为。

### 10.5 落地顺序（在 §9 基础上微调）

1. `MockPageBuild.ps1`（纯）+ 单测 —— **含 §4.5 的 `New-DigitSwapVariants`**。
2. `Render-MockPage.ps1` —— 办公室 PC 手工跑通，肉眼比对真实页面。
3. `Calibrate-MockRowGeometry.ps1` + `-VerifyAgainst`（**退出码卡点**）。
4. `OldSnapMockMatch.ps1` 接线，`MockMatch.Enabled` 打开，沿用 `$snapExists` 门。
5. 变体命中时在 `検証` 列写出疑似正确值（只提示）。
