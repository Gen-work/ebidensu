# SnapVerify 规划文档 — 截图阶段即时异常检测

状态: **规划中（未实装）**。本文档是后续开发会话的设计依据。
来源: 2026-06-12 规划讨论（OCR 完成后的下一步功能群）。

实装前必读: 本文末尾「未确认事项清单」中所有 `[待确认]` 项需要操作员
提供实际样本/答复后才能写解析代码。已拿到答复的决策记录在「已确认决策」。

---

## 1. 目标（5 个功能 + 2 个附带改进）

| # | 功能 | 阶段 | 一句话 |
|---|------|------|--------|
| F1 | 異常終了即时检测 | GiftHmSnap / GfixHmSnap | 截图后读页面文本，発見「異常終了」且时间在本次窗口内 → NG 待人工复核 |
| F2 | 对象数据不存在检测 | GiftMqSnap | Correl_ID + 时间窗口的组合不在查询结果里 → NG |
| F3 | Jenkins 文件不存在检测 | GiftJenkins / GfixJenkins | 页面文件列表里没有匹配文件 → NG（现状只有淹没在日志里的 warn） |
| F4 | NoGfix 旧记录标注 | GiftJenkinsNoFile | 本应无文件却发现旧记录 → 记录其位置，Mark 阶段画框+加「旧数据不影响本次验证」说明 |
| F5 | 对象数据像素定位 | 所有截图阶段 | 截图时定位目标行的像素矩形，供 Mark 阶段精确画框 |
| A1 | 页面文本随图存档 | 所有截图阶段 | Ctrl+A 文本存为 PNG 旁的 .txt，可离线重跑判定/复核 |
| A2 | 文本轮询代替傻等 | 所有截图阶段 | 轮询页面文本直到 Correl_ID 出现或超时，再截图（消灭"截早了"） |

设计总原则（本次讨论确认）:

- **剪贴板文本优先，图像/OCR 为辅**。Windows OCR 在办公机上 WinRT `.Text`
  属性返回空的问题未解（见 docs/SendVsGift.md），词级 bounding box 可用性
  未验证，定位/检测功能不押在 OCR 上。
- 判定逻辑全部进**纯函数库**（无 COM、无 SendKeys、可单测），Snap 脚本只做接线。
- NG 语义沿用 SendVsGift 惯例，不发明新机制。

---

## 2. 已确认决策（2026-06-12 操作员答复）

### 2.1 NG 落盘方式

复用现有 snap 字段（`GIFT_HM_snap` / `GIFT_MQ_snap` / `GIFT_Jenkins_snap` /
`GFIX_*` / `GIFT_noGfixfile_snap`）的值域，扩展为:

- `0` / 空 = 未做
- `1` = 截图完成且检测通过（或检测被禁用）
- `2` = 截图完成但检测 NG —— **仍算 pending**，下次运行该阶段会重新截图重新
  检测；跑完输出 NG 汇总（学 SendVsGift 的 NG summary）。

不新增判定列。现有 pending 判定（`-ne '1'` / `-eq '1'` skip）天然把 `2`
当 pending，无需改 MappingStore。每个 NG 同时写一条 `status\progress.jsonl`
事件（Action='verify', Status='ng', Message=原因），人工复核的依据在
progress.jsonl + 截图旁的 .txt 存档里。

### 2.2 时间窗口基准（重要，操作员详细说明过业务流程）

业务现实: 同一项目的截图**不一定同一天完成**，存在行级差异。典型一天的
流程: Kase-san 发来标注了范围的 excel snap → 操作员查 WBS 确认范围 →
`Generate-HostOpenMapping -Add` 加入新 JOB → 新行没有任何进度标记，
`Expected_Time` 列也是空的 → 这些空行就是本轮时间要解决的对象。

因此采用 **逐行 `Expected_Time` 列 + 开跑时一次性询问** 的方案:

1. Snap 阶段（HM/MQ/Jenkins）启动时询问一次:
   - `[Enter]` = 用当前本机时间
   - 直接输入 `yyyy/MM/dd HH:mm:ss` = 用输入值
   - `n` = 本轮不做时间检查（no time）
2. Tolerance 默认 **前后 30 分钟**（`ToleranceMinutes = 30`），询问是否修改。
3. 输入完成后，默认应用到**本轮 pending 的所有行**: `Expected_Time` 为空的
   行写入所选时间并持久化到 mapping；已有值的行保留各自的值（与
   Resolve-ExpectedTime 的 keep-current 行为一致）。
4. 选 `n`（无时间）时窗口检查跳过:
   - HM: 页面上发现異常終了 → 无法区分新旧，降级为警告 + 当场询问操作员;
   - MQ: 只做"Correl_ID 匹配行存在与否"检查;
   - Jenkins: 只做"文件名存在与否"检查（Parse-JenkinsList 本来就支持
     ExpectedTime=MinValue 时只回 Found）。

复用件: `ExpectedTime` 配置节（VerifyConfig.psd1 已有 TimeColumn/IdColumn/
TimeFormat）、`Resolve-ExpectedTime.ps1` 的交互范式（但它是逐 correl 询问，
本功能是批量一次询问，需新写批量版函数）。

### 2.3 HM 異常終了判定范围

**窗口内 NG，窗口外仅警告**: 異常終了行的时间落在 `Expected_Time ±
Tolerance` 内 → snap 字段=2 + NG 汇总；窗口外（历史遗留的旧异常）→ 控制台
黄色警告 + progress.jsonl（Status='warn'），不算 NG、不阻塞。

### 2.4 MQ NG 条件（三条都启用）

1. 无 Correl_ID 匹配行，**或页面出现 "No Data!" 字样** → NG。
2. 有匹配行但 RecvDate 不在时间窗口内（可能是旧数据）→ NG。
3. `Rtncd` / `Rsncd` 非零 → 也标记 NG。
   注: 操作员表示**不清楚 Rtncd/Rsncd 的业务含义**，先按"非零即异常"标记，
   语义待向 host 团队确认（NG 汇总里单独注明原因是 rtncd/rsncd，便于复核时
   区分）。

---

## 3. 共通基盘（Milestone 1，所有功能的前置）

### 3.1 新纯函数库 `SnapVerify.ps1`

无 `param()` 块（可安全 dot-source，进 CLAUDE.md 的 dot-source 白名单），
ASCII 源码，日文字符串用 `[char]` 构造（如 異常終了 =
`[char]0x7570+[char]0x7570...` 实际为 0x7570 0x5E38 0x7D42 0x4E86）。
单测进 `Tests\Test-SnapVerify.ps1`，并入 Run-Tests.ps1。

函数清单（草案）:

```
ConvertFrom-HmPageText      -Text            -> 结构化行（待样本确认格式）
Test-HmAbend                -Rows -CorrelId -Expected -ToleranceMin
                            -> @{ Verdict='ok'|'ng'|'warn'; Reason; AbendRows }
ConvertFrom-MqPageText      （吸收 Parse-GiftMq.ps1 的正则，函数化以便单测）
Test-MqRecord               -Parsed -CorrelId -Expected -ToleranceMin
                            -> @{ Verdict; Reason; MatchedRow }
ConvertFrom-JenkinsListText （吸收 Parse-JenkinsList.ps1 的正则）
Test-JenkinsFile            -Files -CorrelId -Expected -ToleranceMin -ExpectExists
                            -> @{ Verdict; Reason; File }   # ExpectExists=$false 即 NoGfix 模式
Resolve-SnapRunTime         批量时间询问（2.2 的交互），返回 @{ Time; ToleranceMin; Disabled }
                            （Read-Host 交互部分放接线侧，纯库只做解析/应用逻辑）
```

既存 `Parse-GiftMq.ps1` / `Parse-JenkinsList.ps1` 保留为薄壳（param() 转调
库函数），不破坏现有调用方（JenkinsDownload）。

### 3.2 文本随图存档（A1）

每次截图同时把 Ctrl+A 文本写到 `snap\<folder>\<Correl_ID_S>.txt`
（UTF-8 无 BOM）。成本≈0，收益: 判定规则改了可离线重跑、Validate 阶段可
复核、NG 争议时有原始依据。受 `SnapVerify.SaveText`（默认 true）控制。

### 3.3 文本轮询代替固定等待（A2）

现状三个 Snap 都是 `Start-Sleep ResultWaitSec` 后直接截图。改为:
搜索动作后轮询（间隔 ~500ms）取页面文本，直到文本中出现 Correl_ID_S 或
超时（`SnapVerify.PollTimeoutSec`，默认 ~10s），然后截图。超时不算失败
（页面可能确实无该数据，那正是 F2/F3 要判定的情况），照常截图进判定。

剪贴板注意: Paste-Replace（粘贴 Correl_ID）和取文本都用剪贴板，顺序必须是
"粘贴搜索 → 等结果 → 取文本 → 截图"，且取文本会清掉剪贴板，下一行循环开头
重新 SetText。frameset 页面取文本前要先 click 对应 frame（Read-PageText.ps1
头部注释）。

### 3.4 配置节（VerifyConfig.psd1 + verify_config.json overlay 均可覆盖）

```powershell
SnapVerify = @{
    Enabled          = $true     # 总开关（出问题可整体关掉回到纯截图）
    ToleranceMinutes = 30
    SaveText         = $true
    PollTimeoutSec   = 10
    PollIntervalMs   = 500
}
```

### 3.5 HmSnap / MqSnap 现代化（顺手但必要）

HmSnap.ps1 / MqSnap.ps1 仍在用裸 Import-Csv/Export-Csv（v2 时代遗留，违反
"ALL scripts use MappingStore"约定），且不写 progress.jsonl。接入 F1/F2 时
一并迁移到 MappingStore（Import-Mapping / Export-MappingAtomic /
Get-PendingRows）+ ProgressLog。这也是 NG=2 语义和原子写的前提。

---

## 4. 各功能设计

### F2 — MQ 数据不存在检测（最先做，纯接线）

流程（MqSnap.ps1 每行循环内，截图前后）:

1. 搜索后轮询页面文本（3.3），存档 .txt（3.2），照常截图。
2. `ConvertFrom-MqPageText` → `Test-MqRecord`（2.4 的三条规则）。
3. ok → snap 字段=1；ng → =2 + 控制台红字 + progress.jsonl；跑完 NG 汇总。

注意: Parse-GiftMq 注释说明 MQ 是 frameset，Ctrl+A 前需 click frame_main —
现有 Click-PageBody 是否点中 frame_main `[待确认 Q4]`。

### F3 — Jenkins 文件不存在（现状升级）

JenkinsSnap GiftRecv/GfixRecv 已取页面文本、跑 Parse-JenkinsList，
`Matched=0` 已写 warn 事件（JenkinsSnap.ps1:348 附近）——只差落盘和可见性:

1. `Matched=0`（或 Found 但时间窗口外）→ snap 字段=2 而非 1。
2. 跑完输出 NG 汇总。
3. 时间窗口直接传给 Parse-JenkinsList 的 `-ExpectedTime -ToleranceMinutes`
   （它已实现 IsInRange 判定）。
4. 不依赖 Ctrl+F hit 计数（程序拿不到），就用页面文本判定。

### F1 — HM 異常終了检测

流程（HmSnap.ps1 每行循环内）:

1. 搜索后轮询文本、存档、截图（同上）。
2. `ConvertFrom-HmPageText` 解析结果表格行 → `Test-HmAbend`:
   - 状态含「異常終了」且行时间在窗口内 → ng（=2）
   - 異常終了但窗口外 → warn（仍=1，黄字+jsonl）
   - 无異常終了 → ok（=1）
3. 无时间模式（2.2 选 n）→ 发现異常終了即当场询问操作员判 ok/ng。

图像容错腿（二期可选）: 剪贴板取不到文本（取文本失败/空）时退回
`Find-Abend.ps1` 模板匹配（green/white 模板 + Row1Top/RowHeight 几何，
`Calibrate-HmGeometry.ps1` 校准）。一期先只做剪贴板腿 + "文本为空则警告"。

前置: **HM 结果页 Ctrl+A 文本样本** `[待确认 Q1]` —— 解析正则没有样本写不了。
另需确认状态列取值集合 `[待确认 Q2]`、HM 是否 frameset `[待确认 Q4]`。

### F4 — NoGfix 旧记录标注（依赖 F5 的定位产物）

业务含义: NoGfix 模式期待"该 Correl_ID **没有** GFIX 文件"。若页面文件列表
里出现了该 ID 的记录，按时间判定是旧数据 → 不是 NG-停止，而是
"需要额外说明"流程:

1. 检测: NoGfix 模式也取页面文本（现状只有 Recv 模式取），
   `Test-JenkinsFile -ExpectExists:$false`；发现匹配文件 → 记录文件行的
   像素位置（F5）+ snap 字段=2（提示后续需额外 mark）。
2. 传递: 位置等信息写 sidecar JSON（`snap\GIFT_noGfixfile\<correl>.note.json`，
   含 PixelRect / FileDateTime / Reason）。ReplaceEvidence 插图时读 sidecar，
   把 payload 编进图片 AltText（扩展现有 `verifyMark|folder|idx|correl`
   元数据通道，新 kind 如 `verifyNote|folder|correl|x,y,w,h`）。
3. 标注: Mark 阶段（MarkGift）读到 verifyNote payload → 画红框 +
   旁边加 TextBox「这是旧数据，不影响本次验证结果」（日文文案待定
   `[待确认 Q3]`，文字从 ProjectLabels.ps1 以 [char] 提供）。
4. **像素→point 换算必须做**: sidecar 存的是截图像素坐标，Excel 画框用
   point 且插入图片可能被缩放。Mark 侧用 `Shape.Width / 图片像素宽` 得
   缩放比再换算，否则画偏。

### F5 — 对象数据像素定位（按页面类型分别实现，放弃通用 OCR 定位）

| 页面 | 定位手段 | 状态 |
|------|----------|------|
| Jenkins（截图时 Ctrl+F 高亮还开着） | `Find-ActiveHighlightRow.ps1`（橙色 FF9632 行扫描）已存在，直接对截好的 PNG 跑 | 零件现成 |
| HM / MQ（表单查询，无 Ctrl+F） | 固定几何: 命中行号从剪贴板解析结果得出，`Row1Top + (n-1)*RowHeight` 算像素矩形（Find-Abend 同款参数，Calibrate-HmGeometry 校准） | 零件现成 |
| OCR word box | 等办公机 WinRT `.Text` bug 解决并验证 BoundingRect 可读后再考虑 | 暂缓 |

产物统一为 sidecar JSON（`<correl>.loc.json`: `{ "x":, "y":, "w":, "h": }`，
截图像素坐标系），ReplaceEvidence → AltText → Mark 的传递与 F4 同一条管道。
定位必须发生在**截图当时**（窗口尺寸/位置可控），不做事后图片分析。

---

## 5. 实施顺序

| 里程碑 | 内容 | 依赖 |
|--------|------|------|
| M1 | SnapVerify.ps1 纯库 + 单测、配置节、批量时间询问、A1/A2 进 MqSnap | Q4 |
| M2 | F2（MQ 判定接线 + MqSnap 迁移 MappingStore/ProgressLog） | M1 |
| M3 | F3（JenkinsSnap NG=2 + 汇总；A1/A2 进 JenkinsSnap） | M1 |
| M4 | F1（HM 解析 + 判定 + HmSnap 迁移 MappingStore；A1/A2 进 HmSnap） | M1、Q1、Q2 |
| M5 | F5 定位（Jenkins 高亮 + HM/MQ 几何，sidecar 产出） | M3/M4 |
| M6 | F4（NoGfix 检测 + AltText 管道 + Mark 画框/TextBox + 像素换算） | M5、Q3 |

每个里程碑独立可交付；NG 判定出问题时 `SnapVerify.Enabled=$false` 整体回退。
所有纯逻辑必须有 Tests\Test-SnapVerify.ps1 用例（用 Q1 拿到的真实样本文本
做 fixture）。COM/SendKeys 接线部分照惯例只能静态检查，需办公机实跑确认。

---

## 6. 未确认事项清单（实装前需操作员答复）

- **[Q1] HM 结果页 Ctrl+A 文本样本**（最高优先，没有它 F1 写不了解析）。
  采集方法: 在 Edge 打开某个 Correl_ID 的 HM 查询结果页（最好找一个有
  異常終了行的），控制台执行:

  ```powershell
  # 3 秒内把 Edge 结果页切到前台（frameset 的话先点一下结果区域）
  Start-Sleep 3; & .\Read-PageText.ps1 | Set-Content hm_sample.txt -Encoding UTF8
  ```

  把 hm_sample.txt 内容（可脱敏）贴回来。MQ 页（含一个 "No Data!" 的样本）、
  NoGfix 的 Jenkins 列表页同样各采一份用作单测 fixture。
- **[Q2] HM 状态列的取值集合**: 正常終了 / 異常終了 之外还有什么（実行中?
  待機? 警告終了?）；異常終了所在行的"前后行时间"具体指哪一列（开始时间/
  结束时间），与 Correl_ID 是否同行。
- **[Q3] F4 的日文说明文案**: 「これは旧データであり、今回の検証結果に
  影響しません」之类，要写进证据 Excel 给 reviewer 看的正式措辞。
- **[Q4] frameset 确认**: HM / MQ 页 Ctrl+A 是否需要先点击 frame_main
  （Read-PageText/Parse-GiftMq 注释暗示 MQ 需要）；现有 Click-PageBody
  点击的位置是否落在正确的 frame 里。
- **[Q5] Rtncd/Rsncd 业务含义**: 向 host 团队确认非零值的含义，决定哪些值
  真正算异常（当前先按非零即 NG 标记）。
- **[Q6] "No Data!" 准确字样**: 大小写/全半角/是否含感叹号，最好在 Q1 的
  MQ 样本里直接体现。

---

## 7. 关联 TODO（本次讨论顺带记录）

- **Generate-HostOpenMapping `-Add` 不能同时按 owner 过滤** —— 操作员日常
  流程是收到 Kase-san 的范围标注后 `-Add` 增量加 JOB，此时无法 owner
  过滤，需修复。（已加入 CLAUDE.md TODOs）
