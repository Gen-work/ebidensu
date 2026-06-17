# SnapVerify 规划文档 — 截图阶段即时异常检测

状态: **M1 + M2 + M3 实装完成；M4–M6 待实装**。本文档是后续开发会话的设计依据。
来源: 2026-06-12 规划讨论（OCR 完成后的下一步功能群）。
2026-06-12 更新: 操作员提供了 HM / MQ 页面真实样本（附录 A/B），Q1–Q4、Q6
已解决，解析规则与重测判定规则已确定。剩余未确认项只有 Q5（Rtncd 语义，
不阻塞实装）。
2026-06-17 更新: **M1**（`SnapVerify.ps1` 纯库 + 单测 + 配置节）与 **M2**
（`MqSnap.ps1` 迁移 MappingStore/ProgressLog + F2 接线: 文本轮询、页面哨兵、
判定 ok=1/ng=2、批量 `Expected_Time` 询问）已实装。新增两个纯函数
`ConvertTo-ExpectedDateTime` / `Set-EmptyRunTimeCells`（已单测）。M3/M4 接线时
照抄 MqSnap 的 `Test-MqSnapDone`（done == '1'）以免 NG='2' 行被当成已完成。
2026-06-17 更新: **M3**（`JenkinsSnap.ps1` 接 F3: GiftRecv/GfixRecv 模式下
轮询页面文本、页面哨兵、存档 .txt，`ConvertFrom-JenkinsListText` +
`Test-JenkinsFile` 判定 ok=1/ng=2、NG 汇总，迁出 `Get-PendingRows` 改用
`Test-JenkinsSnapDone`）已实装。NoGfix（F4）仍走纯截图，留待 M6。F3 纯函数
与单测在 M1 已就绪，本次仅接线。

---

## 1. 目标（5 个功能 + 2 个附带改进）

| # | 功能 | 阶段 | 一句话 |
|---|------|------|--------|
| F1 | 異常終了即时检测 | GiftHmSnap / GfixHmSnap | 截图后读页面文本，発見「異常終了」且时间在本次窗口内 → NG 待人工复核 |
| F2 | 对象数据不存在检测 | GiftMqSnap | Correl_ID + 时间窗口的组合不在查询结果里 → NG |
| F3 | Jenkins 文件不存在检测 | GiftJenkins / GfixJenkins | 页面文件列表里没有匹配文件 → NG（现状只有淹没在日志里的 warn） |
| F4 | NoGfix 旧记录标注 | GiftJenkinsNoFile | 本应无文件却发现旧记录 → 记录其位置，Mark 阶段画框 + 同行 AZ 列写「過去分データー」 |
| F5 | 对象数据像素定位 | 所有截图阶段 | 截图时定位目标行的像素矩形，供 Mark 阶段精确画框 |
| A1 | 页面文本随图存档 | 所有截图阶段 | Ctrl+A 文本存为 PNG 旁的 .txt，可离线重跑判定/复核 |
| A2 | 文本轮询代替傻等 | 所有截图阶段 | 轮询页面文本直到出现预期内容或超时，再截图（消灭"截早了"） |
| A3 | 页面种类哨兵 | HM / MQ / Jenkins | 取到的文本不是该阶段预期的页面形态（焦点错 frame / 完全陌生格式）→ 立即停下询问操作员 |

设计总原则（本次讨论确认）:

- **剪贴板文本优先，图像/OCR 为辅**。Windows OCR 在办公机上 WinRT `.Text`
  属性返回空的问题未解（见 docs/SendVsGift.md），词级 bounding box 可用性
  未验证，定位/检测功能不押在 OCR 上。
- 判定逻辑全部进**纯函数库**（无 COM、无 SendKeys、可单测），Snap 脚本只做接线。
- NG 语义沿用 SendVsGift 惯例，不发明新机制。
- 任何时刻出现非预期页面格式都应停下来请求人工干预（操作员要求）——
  低成本实现方案见 3.6 的"页面种类哨兵"，不做逐字段校验。

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

### 2.3 HM 異常終了判定: 窗口内 NG + 重测时 newest-wins

基本规则: **窗口内 NG，窗口外仅警告**。異常終了行的开始时间落在
`Expected_Time ± Tolerance` 内 → snap 字段=2 + NG 汇总；窗口外（历史遗留的
旧异常）→ 控制台黄色警告 + progress.jsonl（Status='warn'），不算 NG。

**重测规则（操作员补充，关键）**: 失败后重测的间隔可能**小于 30 分钟**，
即同一窗口内会同时出现異常終了行和其后重跑成功的正常終了行。此时
**以 last run 为准**: 取窗口内开始日时最新的一行——

- 最新行是 正常終了 → 判 ok（更早的窗口内異常終了降级为警告，注明
  "retried, last run ok"）
- 最新行是 異常終了 → 判 NG

与 GfixLog.ps1 的 "newest wins" 惯例一致。不依赖页面排序（样本里实际是
降序，与"昇順表示"字样矛盾），解析后自行按開始日時排序。

状态列取值集合（Q2 答复）: **只有 正常終了 / 異常終了 两种**，没有
実行中/警告終了之类。时间列即数据行的第 1、2 字段（開始日時 / 終了日時）。

### 2.4 MQ NG 条件（三条都启用）

1. 无 Correl_ID 匹配行，**或页面只显示 "No Data!"** → NG。
2. 有匹配行但 RecvDate 不在时间窗口内（可能是旧数据）→ NG。
3. `Rtncd` / `Rsncd` 非零 → 也标记 NG。
   注: 操作员表示**不清楚 Rtncd/Rsncd 的业务含义**，先按"非零即异常"标记，
   语义待向 host 团队确认（NG 汇总里单独注明原因是 rtncd/rsncd，便于复核时
   区分）。

### 2.5 F4 标注措辞与位置（Q3 答复）

没有正式文案。做法: 红框框选住文件时间（Jenkins 列表行的日时字段），然后在
**同一行、图片范围之外的 AZ 列**写入 `過去分データー`（[char] 码点:
0x904E 0x53BB 0x5206 0x30C7 0x30FC 0x30BF 0x30FC，经 ProjectLabels.ps1 提供）。
即不用 TextBox，直接写单元格值；列号可配置
（`SnapVerify.NoGfixNoteColumn`，默认 `'AZ'`）。

### 2.6 页面焦点（Q4 答复）

操作员已校准: 现行 `Click-PageBody` 点击位置为距窗口左上角 offset 150px，
**对 HM 和 MQ 都正好落在正确的内容区**。不改点击逻辑；焦点错误的兜底交给
3.6 的页面种类哨兵（取到外层菜单文本时能识别并停下）。

---

## 3. 共通基盘（Milestone 1，所有功能的前置）

### 3.1 新纯函数库 `SnapVerify.ps1`

无 `param()` 块（可安全 dot-source，进 CLAUDE.md 的 dot-source 白名单），
ASCII 源码，日文字符串用 `[char]` 构造（異常終了 = 0x7570 0x5E38 0x7D42
0x4E86，正常終了 = 0x6B63 0x5E38 0x7D42 0x4E86）。
单测进 `Tests\Test-SnapVerify.ps1`（fixture 用附录 A/B 的真实样本），
并入 Run-Tests.ps1。

函数清单（草案）:

```
ConvertFrom-HmPageText      -Text -> 结构化行（解析规则见 4.F1）
Test-HmAbend                -Rows -CorrelId -Expected -ToleranceMin
                            -> @{ Verdict='ok'|'ng'|'warn'|'ask'; Reason; Rows }
ConvertFrom-MqPageText      （吸收 Parse-GiftMq.ps1 的正则，函数化以便单测）
Test-MqRecord               -Parsed -CorrelId -Expected -ToleranceMin
                            -> @{ Verdict; Reason; MatchedRow }
ConvertFrom-JenkinsListText （吸收 Parse-JenkinsList.ps1 的正则）
Test-JenkinsFile            -Files -CorrelId -Expected -ToleranceMin -ExpectExists
                            -> @{ Verdict; Reason; File }   # ExpectExists=$false 即 NoGfix 模式
Get-SnapPageKind            -Phase Hm|Mq|Jenkins -Text -> 页面种类（见 3.6）
Resolve-SnapRunTime         批量时间询问的纯逻辑部分（解析输入/应用到行集合），
                            Read-Host 交互放接线侧
```

既存 `Parse-GiftMq.ps1` / `Parse-JenkinsList.ps1` 保留为薄壳（param() 转调
库函数），不破坏现有调用方（JenkinsDownload）。

### 3.2 文本随图存档（A1）

每次截图同时把 Ctrl+A 文本写到 `snap\<folder>\<Correl_ID_S>.txt`
（UTF-8 无 BOM）。成本≈0，收益: 判定规则改了可离线重跑、Validate 阶段可
复核、NG 争议时有原始依据。受 `SnapVerify.SaveText`（默认 true）控制。

### 3.3 文本轮询代替固定等待（A2）

现状三个 Snap 都是 `Start-Sleep ResultWaitSec` 后直接截图。改为:
搜索动作后轮询（间隔 ~500ms）取页面文本，直到满足"页面种类正确（3.6）且
文本中出现 Correl_ID_S（或 MQ 的 No Data!）"或超时
（`SnapVerify.PollTimeoutSec`，默认 ~10s）。超时后按页面种类决定:
是预期页面但无数据 → 正常进判定（那正是 F1/F2/F3 要判的情况）；
不是预期页面 → 触发哨兵询问（3.6）。

剪贴板注意: Paste-Replace（粘贴 Correl_ID）和取文本都用剪贴板，顺序必须是
"粘贴搜索 → 等结果 → 取文本 → 截图"，且取文本会清掉剪贴板，下一行循环开头
重新 SetText。

### 3.4 配置节（VerifyConfig.psd1 + verify_config.json overlay 均可覆盖）

```powershell
SnapVerify = @{
    Enabled           = $true     # 总开关（出问题可整体关掉回到纯截图）
    ToleranceMinutes  = 30
    SaveText          = $true
    PollTimeoutSec    = 10
    PollIntervalMs    = 500
    NoGfixNoteColumn  = 'AZ'      # F4 的「過去分データー」写入列
}
```

### 3.5 HmSnap / MqSnap 现代化（顺手但必要）

HmSnap.ps1 / MqSnap.ps1 仍在用裸 Import-Csv/Export-Csv（v2 时代遗留，违反
"ALL scripts use MappingStore"约定），且不写 progress.jsonl。接入 F1/F2 时
一并迁移到 MappingStore（Import-Mapping / Export-MappingAtomic /
Get-PendingRows）+ ProgressLog。这也是 NG=2 语义和原子写的前提。

### 3.6 页面种类哨兵（A3，"任何非预期格式都停下来"的低成本实现）

操作员要求"过程中任何时刻出现非预期格式都应停下来请人工干预"，但又担心
实现太繁琐。方案: **不做逐字段校验**，只做一次廉价的页面分类 + 单一询问点。

`Get-SnapPageKind -Phase <Hm|Mq|Jenkins> -Text` 用特征字符串把文本分类:

| 种类 | 特征（任一命中） | 各阶段处置 |
|------|------------------|-----------|
| HmResult | 含 `バッチ処理状況一覧` 或表头 `開始日時<TAB>終了日時` | Hm 预期 → 进判定 |
| MqResult | 含 `Transfer status inquiry results` 或 `Number of records` | Mq 预期 → 进判定 |
| MqNoData | 全文仅 `No Data!`（trim 后） | Mq 合法终态 → 直接判 NG（2.4-1） |
| OuterFrame | 含 `GIFT System`（外层菜单: `<Transfer status>` / `Inquiry` / `<Documents>` / `Download`） | **焦点点错了 frame** → 停下询问 |
| Empty | 文本为空/全空白 | 取文本失败 → 停下询问 |
| Unknown | 以上都不命中 | 陌生格式 → 停下询问 |

"停下询问" = Bring-ShellToFront + 提示
`r=已修正焦点，重试本行 / s=跳过本行 / q=中止`，同时写 progress.jsonl
（Status='warn', Message=PageKind + 文本前 200 字）。`r` 会对同一行重新执行
点击+搜索+轮询。每个 Snap 循环只有这一个询问点，覆盖了"焦点错 frame"
（附录 B-3 的实际案例）和所有未知形态，实现成本是一个分类函数 + 一段
提示代码。

---

## 4. 各功能设计

### F2 — MQ 数据不存在检测（最先做，纯接线）

流程（MqSnap.ps1 每行循环内，截图前后）:

1. 搜索后轮询页面文本（3.3）+ 哨兵分类（3.6），存档 .txt（3.2），照常截图。
2. `ConvertFrom-MqPageText` → `Test-MqRecord`（2.4 的三条规则 + newest-wins
   同 correl 多行时取 RecvDate 最新行判 Rtncd/Rsncd）。
3. ok → snap 字段=1；ng → =2 + 控制台红字 + progress.jsonl；跑完 NG 汇总。

解析格式（附录 B 样本确认）: **每条记录占两行**——第 1 行
`No / Send node / Recv node / Correlid / Send date / Tmode / Recv date /
Rtncd / Rsncd`（TAB 分隔），第 2 行 `Msgid / Reccnt / File size`。
现有 Parse-GiftMq.ps1 的正则只吃第 1 行且 `\s+` 兼容 TAB，**对该样本直接
可用**；第 2 行暂不解析（Reccnt/FileSize 将来有需要再加）。
`Number of records N` 行照旧解析。

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

1. 搜索后轮询文本 + 哨兵分类、存档、截图（同上）。
2. `ConvertFrom-HmPageText` 解析数据行 → `Test-HmAbend`（2.3 的
   窗口内 NG / newest-wins / 窗口外警告规则）。
3. 无时间模式（2.2 选 n）→ 发现異常終了即当场询问操作员判 ok/ng。
4. 窗口内一行数据都没有（含解析出 0 行）→ Verdict='ask'，当场询问操作员
   （HM 无数据可能意味着还没跑，业务上未定义，先交人工）。

解析规则（附录 A 样本确认，**字段 TAB 分隔**）:

- 数据行识别: 行首匹配 `^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}\t`。
- TAB split 后: `[0]`=開始日時, `[1]`=終了日時, `[2]`=処理時間,
  `[3]`=バッチID, `[4]`=ＳＳ（注意全角）, `[5]`=処理状態,
  `[6]`=データ作成日, 之后**不可按固定下标取**——異常終了行比正常行多一个
  空字段（処理結果为空: `17<TAB><TAB>◆<TAB>JIDSK01S`）。
- 因此: 処理状態 = 字段中第一个等于 正常終了/異常終了 的值（取值集合仅此
  两种，Q2 已确认）；相関ID = **最后一个字段**；処理件数含千分位逗号
  （`36,117`），解析时去逗号。
- 解析必须过滤 `相関ID == 目标 Correl_ID_S` 的行（页面虽按 correl 查询，
  防御性保留）。
- 表头行（`開始日時\t終了日時\t...`）与查询表单区文本直接忽略（不匹配
  数据行正则）。

图像容错腿（二期可选）: 剪贴板取不到文本（哨兵 Empty 且重试无效）时退回
`Find-Abend.ps1` 模板匹配（green/white 模板 + Row1Top/RowHeight 几何，
`Calibrate-HmGeometry.ps1` 校准）。一期先只做剪贴板腿 + 哨兵询问。

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
3. 标注（2.5 已确认）: Mark 阶段（MarkGift）读到 verifyNote payload →
   红框框住文件**时间字段**的位置 → 在图片所在行、图片范围之外的
   `NoGfixNoteColumn`（默认 AZ）单元格写入 `過去分データー`
   （ProjectLabels [char] 提供）。
4. **像素→point 换算必须做**: sidecar 存的是截图像素坐标，Excel 画框用
   point 且插入图片可能被缩放。Mark 侧用 `Shape.Width / 图片像素宽` 得
   缩放比再换算，否则画偏。

### F5 — 对象数据像素定位（按页面类型分别实现，放弃通用 OCR 定位）

| 页面 | 定位手段 | 状态 |
|------|----------|------|
| Jenkins（截图时 Ctrl+F 高亮还开着） | `Find-ActiveHighlightRow.ps1`（橙色 FF9632 行扫描）已存在，直接对截好的 PNG 跑 | 零件现成 |
| HM / MQ（表单查询，无 Ctrl+F） | 固定几何: 命中行号从剪贴板解析结果得出（解析行序 = 画面行序），`Row1Top + (n-1)*RowHeight` 算像素矩形（Find-Abend 同款参数，Calibrate-HmGeometry 校准） | 零件现成 |
| OCR word box | 等办公机 WinRT `.Text` bug 解决并验证 BoundingRect 可读后再考虑 | 暂缓 |

产物统一为 sidecar JSON（`<correl>.loc.json`: `{ "x":, "y":, "w":, "h": }`，
截图像素坐标系），ReplaceEvidence → AltText → Mark 的传递与 F4 同一条管道。
定位必须发生在**截图当时**（窗口尺寸/位置可控），不做事后图片分析。

---

## 5. 实施顺序

| 里程碑 | 内容 | 依赖 | 状态 |
|--------|------|------|------|
| M1 | SnapVerify.ps1 纯库 + 单测（附录样本做 fixture）、配置节、批量时间询问、页面哨兵 | — | **done** (v2.9.0) |
| M2 | F2（MQ 判定接线 + MqSnap 迁移 MappingStore/ProgressLog；A1/A2 进 MqSnap） | M1 | **done** (v2.9.4) |
| M3 | F3（JenkinsSnap NG=2 + 汇总；A1/A2 进 JenkinsSnap） | M1 | **done** (v2.9.5) |
| M4 | F1（HM 解析 + 判定 + HmSnap 迁移 MappingStore；A1/A2 进 HmSnap） | M1 | todo |
| M5 | F5 定位（Jenkins 高亮 + HM/MQ 几何，sidecar 产出） | M3/M4 | todo |
| M6 | F4（NoGfix 检测 + AltText 管道 + Mark 画框/AZ 列写入 + 像素换算） | M5 | todo |

M2 实装备注: pending 过滤用本地 `Test-MqSnapDone`（done == 恰好 '1'），**不**用
`Get-PendingRows`（其 `Test-SnapDone` 把任何非 '0' 值都算 done，会把 NG='2'
藏掉）。M3/M4 照抄此模式。`SnapVerify.Enabled=$false` 可整体回退到纯截图。

每个里程碑独立可交付；NG 判定出问题时 `SnapVerify.Enabled=$false` 整体回退。
所有纯逻辑必须有 Tests\Test-SnapVerify.ps1 用例。COM/SendKeys 接线部分照
惯例只能静态检查，需办公机实跑确认。

---

## 6. 未确认事项清单

- **[Q5・未解决] Rtncd/Rsncd 业务含义**: 向 host 团队确认非零值的含义，
  决定哪些值真正算异常（当前先按非零即 NG 标记，不阻塞实装）。
- ~~[Q1] HM/MQ 页面文本样本~~ → 已提供，见附录 A/B。
- ~~[Q2] HM 状态列取值~~ → 只有 正常終了/異常終了 两种（2.3）。
- ~~[Q3] F4 文案~~ → `過去分データー`，同行 AZ 列（2.5）。
- ~~[Q4] 焦点/frame~~ → Click-PageBody 150px offset 已校准对 HM/MQ 均有效
  （2.6）；焦点错误由哨兵兜底（3.6）。
- ~~[Q6] "No Data!" 准确字样~~ → 确认就是 `No Data!`（附录 B-2）。
- 待采集（不阻塞 M1–M4）: NoGfix 的 Jenkins 文件列表页样本一份，给
  Test-JenkinsFile 的 fixture 用（格式与 Parse-JenkinsList 已知格式一致的
  概率高，实装 M5/M6 前确认即可）。

---

## 7. 关联 TODO（本次讨论顺带记录）

- **Generate-HostOpenMapping `-Add` 不能同时按 owner 过滤** —— 操作员日常
  流程是收到 Kase-san 的范围标注后 `-Add` 增量加 JOB，此时无法 owner
  过滤，需修复。（已加入 CLAUDE.md TODOs）

---

## 附录 A — HM 结果页 Ctrl+A 文本样本（2026-06-12 实采，字段 TAB 分隔）

注意点: 数据区第 2 行是異常終了 + 同窗口内 11:05 重跑成功的实例（2.3
newest-wins 规则的现实依据）；異常終了行処理結果为空导致比正常行多一个
空 TAB 字段；ＳＳ 列为全角；件数带千分位逗号；页面虽写"昇順表示"但实际
输出为降序——解析后必须自行排序。

```
 掲示板 	 管理 	 受信 	 特殊業務 

バッチ処理状況一覧			IDSXA041
test
相関ID	
JIDSK01S
バッチID	
ＳＳ	
処理種別	
データ取込
処理状態	
全て
開始日FROM	
20260612
開始日TO	
20260612
昇順表示(開始日時が古いものを上に表示)
日付はYYYYMMDDで指定
開始日時	終了日時	処理時間	バッチID	ＳＳ	処理状態	データ作成日	処理件数	処理結果	相関ID
2026/06/12 11:05:40	2026/06/12 11:06:57	00:01:17	IDSLA013	K	正常終了	20260424040558	36,117	◆	JIDSK01S
2026/06/12 10:35:40	2026/06/12 10:35:43	00:00:03	IDSLA013	K	異常終了	20260424040558	17		◆	JIDSK01S
2026/06/12 07:51:20	2026/06/12 07:51:41	00:00:21	IDSLA013	K	正常終了	20260612075111	5,764	◆	JIDSK01S
```

## 附录 B — MQ 页 Ctrl+A 文本样本（2026-06-12 实采）

### B-1 正常结果（每条记录两行；操作员从未见过 Rtncd/Rsncd 非零的失败样例）

```
Transfer status inquiry results
Number of records 3
No	Send node	Recv node	Correlid	Send date	Tmode	Recv date	Rtncd	Rsncd	
Msgid	Reccnt	File size
1	JSSS004R	JHM102R	JIDSK05S	2026/06/12 07:54:22	TXT	2026/06/12 07:54:23	0	0	
A2009999A000000000000001462026061122542266800001	11002	4986849
2	JSSS004R	JHM102R	JIDSK05S	2026/06/12 10:32:41	TXT	2026/06/12 10:32:42	0	0	
A2009999A000000000000001432026061201324104600001	12334	5752860
3	JSSS004R	JHM102R	JIDSK05S	2026/06/12 11:01:53	TXT	2026/06/12 11:01:57	0	0	
A2009999A000000000000001432026061202015372200001	12334	5752860
```

### B-2 无数据（合法终态 → 判 NG）

```
No Data!
```

### B-3 焦点错误（点到外层菜单 frame，哨兵 OuterFrame 的特征来源 → 停下询问）

```
GIFT System

<Transfer status>

Inquiry

<Documents>

Download
```
