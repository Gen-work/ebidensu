# ebi-dance 重构计划:从「按步骤分组的脚本集」到「按能力正交的模块积木」

> 本文是 `docs/Generalization-Roadmap.md` 的**替代方案**,不是补充。
> 那份路线图的 M2/M3(core/engine/adapters/profiles 分层)只解决「文件放哪」,
> 不解决「能不能被组合」。本计划换一个轴:**统一 step 契约 + 声明式 JSON 工作流**。

---

## 1. 为什么现在做

### 1.1 现状的问题可以量化

`ebidensu` 现有 91 个 `.ps1` / 26,700 行。文件按**业务步骤**切
(`HmSnap.ps1` / `MqSnap.ps1` / `JenkinsSnap.ps1`),同一个能力反复重写:

| 能力 | 被抄了几遍 |
|------|-----------|
| `Invoke-CropPng`(截图裁剪) | **4 份** |
| Edge 激活 / 挪窗口 / 前台切换 | 各 2 份 |
| 截图 | 分散在 5 个文件 |
| `Read-Host` 人机交互 | 27 个文件、77 处,每处格式都不一样 |
| 「解析页面 → 找目标行 → 判规则」 | HM / MQ / Jenkins 各写一遍 |

### 1.2 真正的驱动力:下一份工作马上就来

今天是 **Host → Open** 迁移。下个月可能是 **Host → Host** 或别的:
**页面不同、下载渠道不同、Excel 格式不同**。旧的 GIFT/GFIX 专用流程可预见地
不会再用。所以:

- 保留一个专用旧项目没有价值 → **就地重构,不另起仓库**
- 新工作必须能**快速搭起来**,而不是再写 26,000 行 → 这是 MVP 的真正验收标准
- 所有 `GIFT` / `GFIX` / `HM` / `MQ` / `Jenkins` 这类**具体名字必须中性化**,
  否则下一个工作又要全局改名

### 1.3 目标

一套正交的、自带机器可读文档的模块库 + 一个 JSON 工作流解释器 + 一份
**Agent 访谈剧本**。让「搭一个新的取证工作流」从"写几千行 PowerShell"变成
"和 Agent 聊半小时 + 审一份 JSON"。

**硬约束**:Windows 10/11 + PowerShell 5.1 + .NET Framework 内置。
零外部依赖、零安装权限、办公 PC 可离线运行。

---

## 2. 已定的决策

| # | 决策 | 取值 | 依据 |
|---|------|------|------|
| D1 | 仓库策略 | **就地重构现仓库**。先打 `freeze/pre-ebi-dance` 冻结标签当回滚点,再在同一棵树上长出新骨架(不用 `spec/gift-gfix`——远端已有同名分支,见 `BACKLOG.md` P0-01) | 旧流程不会再用,留着无价值 |
| D2 | MVP 范围 | **垂直切片** + **必须能用来搭下一份工作** | 只做工具箱不跑真流程,契约错误会在第一次真跑时才暴露 |
| D3 | 命名 | **全面中性化**,见 §4 词汇表。所有具体系统名只出现在 profile 数据里 | 下个月换工作不改一行代码 |
| D4 | 敏感信息 | **交互式掩码模块**:导出 trace/bundle 前逐项问「保留还是掩码」,决策持久化 + 一致性替换 | 用户明确要求 |
| D5 | OCR / 图像判定 | **保留为可选降级层,但强制自校准**:任何依赖它的工作流必须先跑 `ebi calibrate ocr` 通过才允许运行 | 用户顾虑:3/9 分不清、每台机器 OCR 版本可能不同、规则会成累赘 |
| D6 | Agent 接入 | **传输可插拔**。当前是「能上网但不能装 AI 工具」,设计要同时兼容完全离线(剪贴板)和办公 PC 直接跑 Agent 三种 | 用户答:现在是 2,未来要兼容 1 和 3 |
| D7 | 判定逻辑 | **文本档不重写**(`SnapVerify`/`GfixLog` 搬过来);**图像档降级**(`TimeDigitVerify`/`PixelDigitMatch` 进 `legacy/`) | 文本判定精确且跨机一致;图像判定是负债 |
| D8 | 不做 DSL 编译器 | JSON 直接解释执行,无中间表示、无代码生成 | 可调试、可 diff、Agent 可直接改 |
| D9 | 不用 PS `class` | 全部 hashtable + `[pscustomobject]` + 命名约定 | PS 5.1 的 class 跨 dot-source 有作用域坑 |
| D10 | 操作界面 | CLI + ASCII 面板;HTML 报告放 P5;GUI 不做 | 零依赖、RDP 友好、Agent 可读可驱动 |

---

## 3. 架构

```
(仓库根,就地重构)
  kernel/            # 无业务知识:运行器、上下文、模板、ledger、trace、gate、schema、mask
  modules/           # 能力原语,按 group 分,每个 step 一个 .ps1 + 内嵌 manifest
    browser/ screen/ file/ excel/ table/ verify/ human/
  workflows/         # JSON 工作流(人/Agent 编写的产物)
  profiles/          # 项目数据:页面绑定、判定规则、清单 schema、工作簿布局
  docs/              # 一部分手写、一部分由 manifest 自动生成
  legacy/            # 旧 phase 脚本 + 图像判定层,迁一条删一条
  Tests/             # 沿用现有 Run-Tests.ps1 模式
  ebi.ps1            # 唯一入口 CLI
```

### 3.1 Step 契约(整个项目的地基)

每个 step 是一个 `.ps1`,导出 `$Manifest` + `Invoke-Step`。
**manifest 就是 Agent 的 API 文档来源** —— 文档全自动生成,永远不会和代码脱节。

```powershell
# modules/browser/browser.find.ps1
$Manifest = @{
  id       = 'browser.find'
  group    = 'browser'
  summary  = '在当前页面用 Ctrl+F 查找指定文本,返回是否命中'
  effects  = 'ui'              # pure | read | write | ui | destructive
  needs    = @('foreground')
  inputs   = @{
    term       = @{ type='string'; required=$true; desc='要查找的完整字符串' }
    closeAfter = @{ type='bool';   default=$true;  desc='查完是否 Esc 关闭查找框' }
  }
  outputs  = @{
    hit  = @{ type='bool' }
    rect = @{ type='rect'; desc='活动高亮行的像素矩形,未命中为 null' }
  }
  failures = @('not_found','no_foreground_window')
  example  = @{ use='browser.find'; with=@{ term='{{item.key}}' } }
}
function Invoke-Step { param($In, $Ctx) ...; return @{ ok=$true; hit=$true; rect=$r } }
```

铁律:
- **step 不知道自己在什么业务里**。`browser.find` 只认识字符串,不认识「相关 ID」。
- 所有业务知识通过 `with:` 从 workflow JSON 注入。
- step 之间**只**通过返回值通信(现有的 `$Global:Timing` / `$Global:Shell` 要消掉)。
- 失败用返回值表达(`@{ ok=$false; failure='not_found' }`),不靠异常。

### 3.2 Runner 的四个横切能力

1. **容错策略**:`onError: { policy: retry|ask|skip|fail, times: 3, backoffMs: 800 }`,
   顶层声明一次、单步可覆盖。现在这套逻辑散在各个 `do { $attempt++ } while` 里。
2. **幂等 + 断点续跑**:每个 item 的每个 step 写 ledger(`run/<runId>/ledger.jsonl`),
   重跑按 `idempotencyKey` 跳过。把现有 bitmask(1/2/4)通用化为 `flow.checkpoint`,
   位定义由 profile 声明。
3. **Trace**:每步写输入、输出、耗时、产物路径、页面文本哈希、判定详情。
   这是 Agent 优化循环的燃料。
4. **人工关卡**:`human.gate` 统一渲染 ASCII 面板(发生了什么 / 下一步会做什么 /
   证据在哪 / 可选动作),替代 77 处各写各的 `Read-Host`。

---

## 4. 中性词汇表(你说"不知道该怎么叫"的那部分)

**这是让下一份工作能直接复用的关键。** 工作流里只出现左列;右列只出现在
profile 数据里。

### 4.1 核心名词

| 中性名 | 含义 | 现在对应 |
|--------|------|----------|
| `worklist` | 工作清单 CSV,一行 = 一件要做的事 | `mapping_<Owner>.csv` |
| `item` | 清单里的一行 | 一个 correl 行 |
| `key` | item 的主键(profile 声明是哪一列) | `Correl_ID_S` |
| `group` | item 的分组属性(用于合并页面访问) | `JOB_NAME` / `TO_code` |
| `deliverable` | 交付物工作簿 | `Excel_NAME` 指向的证据簿 |
| `side` | 对照面。profile 声明有哪几面及其显示名 | `GIFT` / `GFIX` → 建议 `before` / `after` |
| `capture` | 一次截图 + 页面文本归档的产物 | `snap\<folder>\<id>.png` + `.txt` |
| `verdict` | 一次判定的结论(`ok` / `ng` / `unknown`) | `GIFT_MQ_snap` 的 1/2 |

### 4.2 页面角色(page role)—— 按**结构**分,不按用途、不按系统名分

> **v2 修订(操作员反馈)**:初版按「你要从这页拿什么」分类,是错的 ——
> 一页上经常要取好几份证据(既截图又点链接下载,文档页也可能要下载)。
> 用途不唯一,不能当分类轴。**role 描述页面的形状(怎么找到东西),
> 在一页上做几件事是 action,可以有任意多个。**

| role | 结构特征 | 定位方式 | 现在对应 / 你提到的 |
|------|----------|----------|---------------------|
| `entry` | 没有目标数据,只是入口 | 不需要定位 | 各系统首页 |
| `form` | 有输入框,要填要提交 | 焦点序列 | HM / MQ 查询画面 |
| `record` | 单条记录,「标签: 值」结构 | 按标签取值 | HM 处理结果画面 |
| `list` | 多行表格 | 解析全表 → 定位目标行 | MQ 转送状态、Jenkins 文件列表、作业列表、バッチ処理一覧、ETA 页 |
| `document` | 排版好的成品,无可定位结构 | 整页 / 按区域 | 帳票 preview |

**5 个,不是 6 个。** 初版的 `artifact` 删掉了 —— 下载链接是长在某一页上的
**元素**,不是一种页面形状。下载变成 action。

**action(一页可有多个)**:`read` / `locate` / `capture` / `download` /
`input` / `navigate`。典型组合:

```
list 页:  read → locate(找到我那一行) → capture(截图) → download(点那行的链接)
                                      └→ capture(再截一张别的区域)
```

同一页多次 capture,产物用 tag 区分:`<key>__<tag>.png`。

`monitor` 这类词我**故意不用**:MQ 转送状态的结构就是一张多行表格,和 Jenkins
文件列表**完全同型**,应该用同一套 step 处理。

### 4.3 动作(工作流命名)

| 中性动词 | 含义 | 现在对应 |
|---------|------|----------|
| `capture` | 抓证据(截图 + 文本 + 判定) | `*Snap` 系列 |
| `collect` | 下载文件并归档 | `GfixLogDownload` / Jenkins 下载 |
| `compose` | 把证据组装进交付物工作簿 | `Replace*` |
| `annotate` | 在工作簿上画框 / 标注 | `Mark*` |
| `review` | 人工复核 + 记录决定 | `Review*` |
| `deliver` | 交付(文件 / 邮件 / 检查表) | `Deliver*` / `CheckSheet` |
| `sync` | 与基线对比 / 同步 | `Align` |

工作流 id 形如 `before.transferStatus.capture`,取代 `GiftMqSnap`
——**注意用的是 page 名(`transferStatus`),不是 role 名(`list`)**:
同一侧经常有好几个同 role 的页面(当前工作里 MQ 転送状态一览和 Jenkins
文件列表都是 role `list`),按 role 命名会互相撞 id(`spec/PROFILE-SCHEMA.md`
§3、`spec/VOCABULARY.md` §2.6,P0-R1)。
capture 目录形如 `capture/before_transferStatus/<key>.png`,取代
`snap/GIFT_MQ/<id>.png`。

### 4.4 命名怎么落地

- `profiles/<name>/vocabulary.json` 声明本项目的 side 名、role 绑定、列名映射
- `ebi explain` 输出时**同时显示中性名和本项目显示名**:`list(MQ転送状態)`
- 换工作时:复制一份 profile,改 vocabulary + 页面绑定 + 规则,工作流 JSON 大部分能直接抄
- 动词表是**开放的**:加一个新 verb 只是补一行文档,没有代码依赖它

### 4.5 主键:复合 + 规则增量学习

两个来自现场的修正(详见 `docs/ebi-dance/spec/PROFILE-SCHEMA.md` §6):

**(a) 主键可以是复合的。** 当前工作里 `Correl_ID_S` 相同但 `JOB_NAME` 不同,
单列分不开两件事。所以 `key.columns` 是数组,而且工作流可以覆盖成更粗的键。

**(b) 变体规则不是一次声明,是长出来的。** 没人能预先写全:这个系统带时间戳、
那个不带;同名文件一堆、时间戳还对不上(创建时间 ≠ 接收时间);还混着完全没
时间戳的文件,而它们可能在别的表里有自己的时间记录。

所以运行时匹配不唯一就走**歧义面板**:

```
  ⚠ 找不到唯一匹配:key = ABC123 / JOB_A
  找到 4 个候选,没有一个能靠现有规则确定:

  # │ 候选                       │ 创建时间 │ 接收时间 │ 大小
  1 │ ABC123.260824.10515511.dat │ 09:51:02 │ 09:50:03 │ 1.2 MB
  2 │ ABC123.260824.10515533.dat │ 09:53:40 │ 09:53:12 │ 1.2 MB
  3 │ ABC123.dat                 │ 08:12:00 │ (无)     │ 1.2 MB
  4 │ ABC123_old.dat             │ 昨天     │ (无)     │ 0.9 MB

  我的建议:#2 —— 接收时间最新且落在本次运行窗口内
  ⚠ 不确定:#1 和 #2 只差 3 分钟,如果本次是重跑,可能两个都是本次的

  1-4=选  v=看详情  n=都不对(标 unknown)  q=中止
  选完之后:要不要把这次的判断存成规则?(y/n)
```

四条要求:**摊开全部候选**(不许只显示"最佳")、**每个候选带全部证据**、
**给建议+理由+不确定点**、**选完固化成规则**。规则库随使用长大。

对比旧工具:它静默选最新的,只打一行 `[WARN] N 个候选,选了最新的` ——
人根本不会去看那行。

### 4.6 页面解析器:不猜,用 `ebi grammar tune` 调

四种解析器(分隔符 / 标签-值 / 定宽列 / 正则逃生舱)覆盖不了所有页面,而且
**即使覆盖得了,参数也没人能一次猜对**。所以配一个交互式调试器:喂一份真实
页面文本 → 渲染解析结果表格 → 改参数即时重解析 → 满意就存进 profile
**同时存成 fixture**(每次调 grammar 自动积累一个回归测试)。

关键一条:**未识别的行必须显式报出来**。旧工具最恶劣的 bug 就是静默丢行 ——
单位数小时的行被正则漏掉,页面上明明有,判定却说"文件不在列表里"。

可选 `a` 键把文本 + 当前结果交给 AI 提议 grammar,**但 AI 的提议同样要在这个
循环里跑给人看**,不能直接采信。

---

## 5. 模块目录(8 组 ≈ 86 个 step)

`[MVP]` 进第一版(25 个),其余按阶段。「来源」指明可直接复用的现有实现 ——
**多数 step 是包装,不是新写**。

### G1 `browser.*` — 浏览器 / 前台驱动(17)

| step | 作用 | 阶段 | 来源 |
|------|------|------|------|
| `browser.ensure` | 找到并激活浏览器(进程句柄优先,标题回退) | MVP | `Common.ps1 Activate-EdgeWindow` |
| `browser.assert_page` | 断言当前是期望角色的页面,不符则 gate | MVP | `SnapVerify Get-SnapPageKind` |
| `browser.navigate` | Ctrl+L 粘贴 URL 回车;或降级为提示人工打开 | MVP | 新写(薄) |
| `browser.focus_body` | 点击正文区取得键盘焦点 | MVP | `Common.ps1 Click-PageBody` |
| `browser.read_text` | Ctrl+A/Ctrl+C 取全文 | MVP | `Read-PageText.ps1` |
| `browser.wait_for` | 轮询页面文本直到匹配/超时,可同时归档留档 | MVP | `MqSnap Wait-MqPageReady` 通用化 |
| `browser.send_keys` | 发送任意按键序列 | MVP | `Common.ps1 Send-Key` |
| `browser.tab_to` | Tab / Shift+Tab N 次 | MVP | `Common.ps1 Send-Tab` |
| `browser.fill` | 当前焦点输入框 Ctrl+A + 粘贴 | MVP | `Common.ps1 Paste-Replace` |
| `browser.submit` | Enter / 点击指定坐标按钮 | MVP | `Common.ps1 Send-Enter` |
| `browser.find` | Ctrl+F 查找精确串并返回命中与否 | MVP | `SnapVerify Get-JenkinsSearchTerm` 调用侧 |
| `browser.verify_action` | **包装器**:动作前后对比页面文本,无变化即判失败 | P2 | 新写(容错关键件) |
| `browser.tab_to_labeled` | 逐次 Tab + 读焦点元素文本比对定位,替代盲数 Tab | P3 | 新写 |
| `browser.find_active_row` | 定位活动高亮行的像素矩形 | P3 | `Find-ActiveHighlightRow.ps1` |
| `browser.click_at` | 按窗口相对坐标点击 | P3 | `Common.ps1 MouseAPI` |
| `browser.scroll` | 滚动 N 屏 / 到顶 / 到底,返回是否到底 | P3 | 新写 |
| `browser.download_link` | 触发下载并交给 `file.wait_for_download` | P3 | `JenkinsDownload.ps1` |

> `browser.read_html`(读 HTML 源码)**故意不做主路径**:纯 SendKeys 环境下
> Ctrl+U 会开新标签页,状态难恢复,很多内网页面也禁用。主路径是
> `browser.read_text` + `verify.parse_text` 的声明式 grammar,现有
> `ConvertFrom-HmPageText` / `ConvertFrom-GfixJobListText` 已证明足够。

### G2 `screen.*` — 截图 / 图像(12)

| step | 作用 | 阶段 | 来源 |
|------|------|------|------|
| `screen.capture_window` | 指定句柄窗口截图 | MVP | `Common.ps1 Take-WindowScreenshot` |
| `screen.capture_region` | 指定 x/y/w/h,自动 clamp 到屏幕边界 | MVP | `ScreenRegion.ps1 Resolve-ScreenRegion` |
| `screen.fit_window` | 窗口 MoveWindow 到指定尺寸并挪离屏幕边缘 | MVP | `MqSnap Move-EdgeAwayFromBorder` |
| `screen.crop` | 四边裁剪(去窗口阴影),支持 per-role 覆盖 | MVP | `Resolve-DirectionalCrop` + `Invoke-CropPng`(**消掉 4 份重复**) |
| `screen.save` | 按命名模板定位保存 | MVP | 新写(薄) |
| `screen.locate_template` | 模板图匹配返回坐标 | P3 | `Locate-ByImage.ps1` |
| `screen.annotate` | 在图上画框 | P3 | `SnapLocalize.ps1` |
| `screen.ocr` | 区域 OCR **(降级层,需校准,见 §6)** | P4 | `OcrWindows.ps1` |
| `screen.preprocess` | 放大 / 灰度 / 对比度 / 二值化 **(降级层)** | P4 | `ProcessTime ConvertTo-ProcessTimeOcrImage` |
| `screen.stitch_scroll` | **滚动拼接长图**(纵/横):滚动截图 + 重叠带匹配 | P5 | 新写,复用 `Locate-ByImage` 的 LockBits |
| `screen.diff` | 两图像素比对 **(降级层)** | P5 | `PixelDigitMatch.ps1` 通用化 |
| `screen.render_html` | Edge 渲染本地 HTML 并截图(mock page) | P5 | `mock-page/` |

### G3 `file.*` — 文件与下载(11)

| step | 作用 | 阶段 | 来源 |
|------|------|------|------|
| `file.find` | glob/key 查找,**全角回退 + 时间戳变体容忍** | MVP | `WorkbookResolver FullWidthFilenameResolver` + `MappingStore Resolve-CorrelFilePath` |
| `file.assert_exists` | 存在性断言,不存在按策略 gate | MVP | 新写(薄) |
| `file.wait_for_download` | 监视目录直到新文件出现且大小稳定 | P2 | `JenkinsDownload.ps1` |
| `file.rename` | 按模板重命名 | P2 | 新写 |
| `file.move` / `file.copy` | 移动 / 复制 | P2 | 新写 |
| `file.newest` | 候选集里取最新(mtime 或解析出的时间戳) | P2 | `JenkinsDownload Sort-JenkinsFilesNewestFirst` |
| `file.unzip` | 解压,**保留 entry 原名**放进 per-key 子目录 | P2 | `DfSnap Expand-DfZip`(带 v2.21.0 修正) |
| `file.backup` | 时间戳备份到 bk 目录 | P3 | `BackupJ4.ps1` |
| `file.stat` / `file.hash` | 大小 / 时间 / 哈希(外部改动检测) | P3 | 新写 |
| `file.zip` | 打包 | P4 | 新写 |
| `file.cleanup` | 清理临时产物 | P4 | 新写 |

### G4 `excel.*` — Excel COM(16)

| step | 作用 | 阶段 | 来源 |
|------|------|------|------|
| `excel.open` / `close` / `save` | 工作簿生命周期(COM 释放顺序已封装) | P3 | `ExcelHelpers.ps1` |
| `excel.find_workbook` | 按名称解析(前缀 + 全角回退) | P3 | `WorkbookResolver.ps1` |
| `excel.find_sheet` / `list_sheets` | sheet 定位 | P3 | `ExcelHelpers Get-SheetByName` |
| `excel.find_anchor` | 按列文本找行(定位 key 标签行) | P3 | `ExcelHelpers Get-NextAnchorRow` |
| `excel.read_cell` / `read_range` | 读值 | P3 | `ExcelHelpers` |
| `excel.write_cell` / `write_lines` | 写值 / 写多行文本 | P3 | `Write-PlainText` / `Write-LogLines` |
| `excel.clear_below` | 清空锚点以下 | P3 | `Reset-SheetBelowRow` |
| `excel.insert_picture` | 定位插图(前置 / 后置) | P3 | `Insert-Picture*` |
| `excel.export_pictures` | 把内嵌图导出成 PNG(**Excel 后台截图**) | P3 | `EvidenceImageExport.ps1` |
| `excel.draw_rect` | 画框 + 写 AltText 元数据 | P3 | `Add-RedRectangle` + `Set-ShapeMetadata` |
| `excel.remove_shapes` | 按 AltText 前缀清除标记 | P3 | `Remove-MarkShapes` |
| `excel.set_format` | 数字格式 / 字体 / 填充 / 边框 / 列宽 | P3 | `ProcessTime.ps1` 格式化块 |
| `excel.add_hyperlink` | 单元格超链接 | P3 | `ProcessTime.ps1` |
| `excel.set_conditional_format` | 条件格式 | P4 | `TimeDigitVerify` 的公式生成 |
| `excel.replace_sheet` | 用源 sheet 就地替换目标 sheet | P4 | `DeliverFiles.ps1` |
| `excel.probe_format` | 只读格式探针(校准辅助) | P4 | `Probe-SheetFormat.ps1` |

### G5 `table.*` + `progress.*` — 工作清单与进度(11)

| step | 作用 | 阶段 | 来源 |
|------|------|------|------|
| `table.load` | 读 CSV(BOM 策略) | MVP | `MappingStore Import-Mapping` |
| `table.save` | 原子写 | MVP | `MappingStore Export-MappingAtomic` |
| `table.ensure_columns` | 按 profile schema 补列 | MVP | `Ensure-MappingColumns` |
| `table.select` | 筛选(pending / 指定 key / owner) | MVP | `Get-PendingRows` |
| `table.key` | **复合主键 + 变体规则**;歧义时返回全部候选 + 证据(见 §4.5) | MVP | `Get-CorrelIdAliases` / `Test-CorrelIdEquivalent` → 可配置 normalizer |
| `table.set` | 写字段 | MVP | `Update-MappingRows` |
| `flow.checkpoint` | 位掩码 / 值标记完成(位定义来自 profile) | MVP | `Set-MappingBit` |
| `progress.event` | 追加 jsonl 事件 | MVP | `ProgressLog.ps1` |
| `progress.status` | 渲染 ASCII 进度表 | MVP | `VerifyTool Show-Status` |
| `table.derive` | 从 Excel / 另一 CSV 生成清单 | P4 | `Generate-HostOpenMapping.ps1` 通用化 |
| `progress.watch` | 只读监视(不锁 CSV) | P4 | `Watch-MappingProgress.ps1` |

### G6 `verify.*` — 判定(纯函数,零 COM)(8)

**分成两档,这是 §6 的核心。**

| step | 档 | 作用 | 阶段 | 来源 |
|------|-----|------|------|------|
| `verify.parse_text` | **文本档** | 按声明式 grammar 把页面文本解析成 records | MVP | `ConvertFrom-*PageText` + `GfixJobList` 统一 |
| `verify.match_record` | **文本档** | 按 key 找行,支持变体 / newest-wins / 时间窗 | MVP | `Get-MatchedRowIndex` / `Select-JenkinsFileCandidate` |
| `verify.assert` | **文本档** | 对 record 字段跑规则表 → ok/ng/unknown | MVP | `Test-HmAbend` / `Test-MqRecord` / `Test-JenkinsFile` 数据化 |
| `verify.time_window` | **文本档** | 运行时间窗口检查 | P2 | `Resolve-SnapRunTime` |
| `verify.crosscheck` | **文本档** | **多来源同一事实一致性检查**,不一致就停 | P2 | `TimeDigitVerify Resolve-*Conflict` 的**思想**(不含 3/9 猜测) |
| `verify.compare_records` | **文本档** | 两组 records 比对 | P4 | `SendMetadata Compare-SendGiftEvidence` |
| `verify.ocr_read` | **降级档** | OCR 图片得到文本(需校准) | P4 | `OcrWindows` + `ProcessTimeParse` |
| `verify.pixel` | **降级档** | 图像比对(需校准) | P5 | `PixelDigitMatch.ps1` |

### G7 `human.*` — 人机关卡(5)

| step | 作用 | 阶段 |
|------|------|------|
| `human.prepare` | 「请把页面准备好 → Enter」阻塞式 | MVP |
| `human.gate` | **标准确认面板**:发生了什么 / 下一步 / 证据路径 / Enter·n·s·q·m | MVP |
| `human.choose` | 候选列表多选一(如多个下载文件选哪份) | MVP |
| `human.input` | 取值输入(默认值 + 校验) | P2 |
| `human.note` | 自由文本备注写回 worklist | P3 |

### G8 `flow.*` — 控制流(runner 内建)(6)

| step | 作用 | 阶段 |
|------|------|------|
| `flow.foreach` | 遍历 `source.select` 选中的行 | MVP |
| `flow.if` | 条件分支 | MVP |
| `flow.checkpoint` | 见 G5 | MVP |
| `flow.group_by` | 分组遍历(合并同组的页面访问) | P2 |
| `flow.call` | 调用子工作流(可复用片段) | P3 |
| `flow.try` | 局部容错策略覆盖 | P3 |

---

## 6. OCR / 图像判定:降级层 + 强制自校准

你的顾虑完全成立,设计上正面回应:

### 6.1 三条硬规则

1. **文本永远优先。** `browser.wait_for` / `screen.capture_*` 组合**必须**同时归档
   Ctrl+A 文本。有文本就绝不 OCR。这把 OCR 从"主路径"降成"没有文本时的备胎"。
2. **降级层默认关闭。** `verify.ocr_read` / `verify.pixel` / `screen.preprocess`
   在 catalog 里标 `tier: fallback`。工作流引用它们时 `ebi lint` 会报
   `需要校准` 警告。
3. **不校准不许跑。** 引用降级层的工作流,`ebi run` 前会检查
   `.ebi/calibration/ocr.json` 是否存在、是否对应当前机器、是否在有效期内。
   不通过直接拒绝运行,提示 `ebi calibrate ocr`。

### 6.2 `ebi calibrate ocr` 干什么

```
ebi calibrate ocr --samples profiles/<name>/ocr-truth/
```

- 读一个**样本目录**:每个样本是 `<图>.png` + `<图>.expected.txt`(已知正确答案)
- 在**本机**跑一遍 OCR,逐字符比对
- 输出:总正确率、**逐字符混淆矩阵**(直接告诉你这台机器把 9 读成 3 的概率)
- 写 `.ebi/calibration/ocr.json`:机器名、OCR 引擎版本、校准日期、通过的样本集哈希
- 低于阈值(profile 声明,默认 99%)→ 校准失败,该机器上禁用降级层

**样本从哪来**:每次人工复核修正一个 OCR 错误时,`ebi calibrate ocr --add`
把那张图和正确答案存进样本目录。**样本集随使用自动长大**,不需要专门造数据。

### 6.3 旧的 3/9 规则怎么处理

`TimeDigitVerify.ps1` / `PixelDigitMatch.ps1` / `OldSnapPixelVerify.ps1`
**移入 `legacy/`,不进 catalog,不被任何新工作流依赖**。它们只用来清那批历史
老快照;清完即弃。

但保留其中**一个思想**并提升为通用 step —— `verify.crosscheck`:

> 同一个事实如果能从多个来源读到(页面文本、文件名、日志、Excel),就都读。
> 不一致 → **停下来问人**,绝不自己挑一个。

这条不含任何 3/9 特化逻辑,对任何项目都成立,而且正是 v2.20 那个 bug
(把正确的 `00:00:01` 改成错的 `00:00:07`)的根治方案。

---

## 7. 敏感信息掩码模块

### 7.1 交互式决策

```
ebi mask scan run/20260820-1  # 或 ebi bundle 自动触发

  发现 6 类候选敏感项:

  [1] 主机名     mq-prod-01.internal.example      出现 47 次
      k=保留 / m=掩码 / M=永远掩码 / K=永远保留 / v=查看上下文  > _
```

- 规则 + 词典双路检测:员工号 `JP\d{6}`、邮箱域、UNC 路径 `\\host\share`、
  `C:\Users\<id>`、内网 URL、日文人名、公司名 / 项目代号词典
- 决策存 `.ebi/redaction.json`(**该文件 gitignore**,只在本机)
- 之后只对**新出现**的候选项发问,已决策的静默应用

### 7.2 一致性替换(关键设计点)

掩码必须是**稳定的双射**:同一个原文永远映射到同一个占位符。

```
mq-prod-01.internal.example  →  <HOST_1>
jenkins.internal.example     →  <HOST_2>
山田太郎                      →  <PERSON_A>
JP123456                     →  <EMPID_1>
```

否则 Agent 读 trace 时看不出「这两行说的是同一台机器」,推理会崩。
映射表存在本机 `.ebi/redaction.json`,**不随 bundle 导出**。

### 7.3 CI 门禁

`ebi mask check` 扫描全部 tracked 文件,发现未掩码的高危项就失败。
沿用 `docs/Generalization-Roadmap.md` 里 S9 `Check-Sensitive.ps1` 的思路,
挂进 `Tests/Run-Tests.ps1`。

---

## 8. Agent 层:三份文档 + 一个循环

### 8.1 `docs/ebi-dance/INTERVIEW.md` —— Agent 访谈剧本(你要的那份)

**这是整个项目对"解放打键人"贡献最大的东西,而且不需要写任何代码就能用。**
一个人想自动化自己的工作流时,把这份文档交给 Agent,Agent 照着问。

结构:

| 阶段 | Agent 要做什么 |
|------|---------------|
| **0 先看,再问** | 让人先手工做一遍并留下素材:每个页面的 Ctrl+A 文本 + 截图 + 一句话说明。**Agent 从素材反推,而不是凭空发问** —— 这能省掉一半问题 |
| **1 边界** | 一件事的最小单位是什么?key 是什么?清单从哪来?大概多少条? |
| **2 页面盘点** | 每页:角色是 §4.2 里哪个?怎么到达?**怎么确认到对了页**(sentinel)?目标行怎么找? |
| **3 判定** | 什么叫 OK?什么叫 NG?**有没有"看不出来"的情况?** ← 这一问决定容错质量 |
| **4 产物** | 存哪、叫什么、进哪个工作簿的哪个位置 |
| **5 变数** | 时间戳会变吗?key 有变体吗?页面会慢吗?会有重复行吗?会翻页吗?会有空结果吗? |
| **6 关卡** | 哪些地方必须人工确认?哪些**绝不能**自动? |
| **7 出草案** | 生成 workflow JSON + profile → `ebi explain` 人工复核 → `ebi dryrun` → 小批量真跑 |

**写进文档的三条铁律**(直接来自你踩过的坑):

1. **"一般都是"就是红旗。** 人说「一般都是 XXX」「应该是」「基本上」时,
   Agent **必须**追问例外,并且默认生成 `policy: ask` 而不是 `auto`。
2. **能从多处读到的事实,就都读。** 不一致就停下问人。绝不自己挑一个"看起来对的"。
   (v2.20 的教训:一个 OCR 误读的秒数,把正确的 `00:00:01` 改成了错的 `00:00:07`。)
3. **宁可多问一次,不可少记一笔。** 判定不确定时输出 `unknown` 触发人工关卡,
   **绝不**降级成 `ok`。

配套 `docs/ebi-dance/INTERVIEW-CHECKLIST.md`:一页纸的问题清单,人也可以自己对着填。

### 8.2 `docs/AGENTS.md` —— Agent 操作手册

- 怎么读 `catalog.json` 组装合法工作流
- workflow JSON 的 schema 规则和模板语法
- 常见组合套路(recipe)
- 怎么读 trace 定位失败
- **权限边界:Agent 只能改 workflow JSON 和 profile JSON,不能改 step 代码。**
  需要新 step 就写成 issue,不许硬塞

### 8.3 `docs/CATALOG.md` + `catalog.json` —— 全自动生成

`ebi docs build` 扫描 `modules/**` 的 `$Manifest` 产出。人读 md,Agent 读 json。

### 8.4 优化循环(传输可插拔,兼容三种环境)

```
办公 PC                                    家里 / Agent
──────────────────────────────────────────────────────
ebi run     → trace.jsonl + 产物
ebi mask    → 交互式脱敏决策            ┌────────────────┐
ebi bundle  → bundle/<runId>.zip        │ Agent 读:      │
   ├ trace.jsonl(每步输入输出耗时)──▶│  catalog.json  │
   ├ pagetext/*.txt(已脱敏)           │  trace.jsonl   │
   ├ capture/*.png(采样)              │  失败样本      │
   └ workflow.json + profile.json      └───────┬────────┘
                                               │ 产出 patch.json
ebi lint  patch.json                           │ (只含 workflow/profile 改动)
ebi apply patch.json  ◀────────────────────────┘
   → 备份 → 应用 → ebi explain 人工复核 → 再跑
```

**传输层三选一,由 `.ebi/config.json` 的 `transport` 字段决定:**

| transport | 适用 | 实现 |
|-----------|------|------|
| `clipboard` | 完全离线 | 沿用 `Pack-LlmContext` / `Apply-LlmPatch` |
| `git` | **你现在这种(能上网,不能装 AI 工具)** | bundle 提交到分支,Agent 在别处读;patch 走 PR |
| `local` | 办公 PC 能跑 Agent | bundle 和 patch 都在本地文件系统,Agent 直接读写 |

**核心设计不变,只有搬运方式不同** —— 所以现在按 `git` 做,将来换 `local`
只是少一步。

### 8.5 `ebi run --guided` —— 人工跟随引导

每个 item 之前用 `human.gate` 面板显示「现在要做什么 / 上一步结果 / 这一步会
改什么」。新手不需要懂工作流名和 CSV 列义就能跟着走完。

---

## 9. CLI 表面

```
ebi help                       # 分组列出全部 step,一行一句话
ebi help browser.find          # 单个 step 的完整 manifest
ebi explain  workflows/x.json  # 渲染成人类可读执行计划(ASCII)
ebi lint     workflows/x.json  # 静态校验:step 存在?参数齐?模板引用得到?降级层校准了?
ebi dryrun   workflows/x.json  # 干跑,不碰真实系统
ebi run      workflows/x.json  # 真跑    (--guided 引导模式)
ebi doctor                     # 环境自检:PS 版本、Excel、Edge、编码策略
ebi grammar tune <text.txt>    # 交互式调页面解析器,存 profile + 存 fixture
ebi calibrate ocr              # OCR 自校准(见 §6.2)
ebi trace    <runId>           # 回放一次运行,渲染 ASCII 时间线
ebi mask     scan|check        # 敏感信息扫描 / CI 门禁
ebi bundle   <runId>           # 打包给 Agent(自动套用掩码决策)
ebi apply    patch.json        # 应用 Agent 补丁(备份 + lint + explain 三道闸)
```

`ebi explain` 输出形态(人和 Agent 审阅同一份):

```
  before.transferStatus.capture  ── 転送状態ページの証跡取得
  ┌ 数据源: worklist.csv  →  before_transferStatus != ok   (待处理 37 行)
  │
  ├ setup
  │   [人工] 请打开 list(MQ転送状態) 页面
  │   [UI  ] 激活浏览器 → 调整窗口 1050x761
  │
  ├ each  (× 37)
  │   [UI  ] 点击正文 → Tab×1 → Enter → Tab×4 → 粘贴 {{item.Correl_ID_S}} → Enter
  │   [读  ] 轮询页面文本 (≤12s)     → 留档 capture/before_transferStatus/<keySafe>.txt
  │   [写  ] 截图 + 裁剪             → capture/before_transferStatus/<keySafe>.png
  │   [纯  ] 解析 → 找行 → 判定      → ok / ng / unknown
  │   [写  ] 标记 before_transferStatus  ← worklist 原子写
  │
  └ 容错: 默认 ask   人工关卡: 1 处   破坏性操作: 0 处   降级层: 未使用 ✓
```

---

## 10. 规模与工期

| 部件 | 新代码量 |
|------|----------|
| kernel(runner / context / 模板 / ledger / trace / gate / schema / mask) | ~1,800 行 |
| MVP 25 个 step | ~2,200 行(多数是包装现有函数) |
| 文档生成器 + CLI | ~700 行 |
| Tests | ~1,500 行 |
| 手写文档(INTERVIEW / AGENTS / PROFILE / 词汇表) | ~1,800 行 md |
| **MVP 合计** | **≈ 7,000–8,000 行** |
| 全量 86 step + 迁完所有流程 | ≈ 22,000–26,000 行(与现有同量级,业务逻辑总量不变) |

### 10.1 排期方式:任务卡,不是周块

**推进方式是碎片时间(工位上的空档),不是整块的开发日。**
所以不排「P1 = 3 周」这种块 —— 那种粒度根本没法开工。

改成**任务卡**,每张满足三个条件:

1. **一次坐下能做完**(30–90 分钟)
2. **做完就能提交**(不留半成品在树上)
3. **有明确的完成判据**(不靠感觉)

阶段只是卡的分组,没有截止日期。**做完多少算多少,随时可停。**

### 10.2 卡的形态

```
[P1-07] browser.wait_for
  做:  轮询页面文本直到匹配/超时,可选归档到文件
  抄:  MqSnap.ps1 Wait-MqPageReady(去掉 MQ 相关的硬编码)
  完成:manifest 通过 lint;dryrun 打印正确;Tests 里的纯函数部分绿
  估:  60 分钟
```

大部分 step 卡是这个形状 —— **抄现有函数 + 去掉硬编码 + 加 manifest**,
所以估时准、风险低,适合碎片时间。

真正需要连续思考的只有少数几张(`kernel/Runner.ps1`、`kernel/Context.ps1`),
这些标 `[需要整块时间]`,攒到有空档再做。

### 10.3 卡的数量

**完整卡片列表见 `docs/ebi-dance/BACKLOG.md`** —— 那是开工时唯一要看的文档。

| 阶段 | 卡数 | 其中 `[整块]` |
|------|------|--------------|
| P0 骨架(含 9 张规格修订卡 P0-R1…R9) | 17 | 2(会话资源通道 / 最小 runner spike) |
| P1 kernel + 25 个 step + 文档生成 | 34 | 3(Context / Runner 主体 / Runner 容错+ledger)+ 1(table.key) |
| P2 对拍验证(含 human.input/run.window、mask-lite 前移两张) | 8 | 1(办公 PC 首跑) |
| P3 新工作实战 | 8 | 0(主要是访谈 + 填 profile) |
| **合计到 P3 可接新工作** | **67 张** | **7 张** |
| P4 excel/file 组 | 22 | 1(办公 PC 冒烟) |
| P5 掩码 + Agent 循环 + 校准 | 14 | 1(掩码一致性替换) |
| **全部** | **103 张** | **9 张** |

> 2026-08-24 契约审查追加了 P0-R1…R6(规格修订)和 P2-07/08;
> 2026-08-25 又追加了 P0-R7…R9(冻结标签改名、Plan/README 同步、
> P0-00 状态修正)。完整卡片列表和当前状态**始终以
> `docs/ebi-dance/BACKLOG.md` 为准**——这张表只做数量级参考。

按每次坐下做 1 张算,**到 P3 大约 67 次空档**。这个数字比"8 周"有用得多 ——
它不依赖你每周能挤出多少小时。

`[整块]` 的 8 张是**设计而非包装**,需要连续思考,也不建议交给较小的模型。
其余 84 张是「抄现有函数 + 去掉硬编码 + 加 manifest」,估时准、风险低。

### 10.4 关于模型

设计阶段(契约、词汇、访谈剧本)需要判断力,用 Opus 5。
实现阶段的 step 卡是机械的包装工作,换更快/更便宜的模型不影响质量。

**不建议换 Claude Fable 5**:它定位是最难的推理 + 长时程代理任务,价格是
Opus 5 的两倍($10/$50 vs $5/$25),而且单次回合可能跑好几分钟 —— 碎片时间
下这是负作用。它的官方指引还明确说,**写得太细的 prompt 反而会降低它的输出
质量**,所以"更容易跟随 prompt"恰好不是它的卖点。

---

## 11. 实施步骤

### P0 — 冻结与骨架(1.5 周)

1. **打冻结标签** `freeze/pre-ebi-dance`(当前 tip),作为回滚点。旧脚本原地
   不动,继续可运行。(不用 `spec/gift-gfix`——远端已有同名分支,见
   `BACKLOG.md` P0-01)
2. ~~**写三份规格 + 一份词汇表**~~ —— **已完成(2026-08-24)**:
   - `docs/ebi-dance/spec/STEP-CONTRACT.md` — manifest 字段、返回值约定、失败表达、副作用等级
   - `docs/ebi-dance/spec/WORKFLOW-SCHEMA.md` — workflow JSON 全字段 + 模板语法
   - `docs/ebi-dance/spec/PROFILE-SCHEMA.md` — profile 结构(页面绑定、规则、清单 schema)
   - `docs/ebi-dance/spec/VOCABULARY.md` — §4 的完整版,含旧名 → 新名对照表
   - `docs/ebi-dance/INTERVIEW.md` — Agent 访谈剧本(**不依赖代码,现在就能用**)
3. **建骨架目录**,把基础设施搬进去(改造点已标注):
   - `ProgressLog.ps1` → `kernel/Trace.ps1`(事件字段泛化:去掉 `correl_id_s`/
     `job_name` 硬编码,改成 `key` + `tags{}`)
   - `MappingStore.ps1` → `modules/table/*`(`Correl_ID_S` 硬编码换成 profile
     声明的主键列)
   - `ScreenRegion.ps1` / `SnapVerify.ps1` / `GfixLog.ps1` → 搬入 `modules/verify/`
   - `TimeDigitVerify.ps1` / `PixelDigitMatch.ps1` / `OldSnapPixelVerify.ps1` → `legacy/`
   - 沿用 `Tests/Run-Tests.ps1` + `Check-Encoding.ps1` 的 ASCII 源码铁律
4. **打通 3 个 step 端到端**:`human.prepare` → `browser.ensure` → `screen.capture_window`

   **验收:一条 5 行的 workflow JSON 能真的存下一张 PNG。**

### P1 — 内核 + MVP step 集(3 周)

按依赖顺序:
1. `kernel/Context.ps1` — vars/item/steps 作用域 + `{{}}` 模板求值(**纯函数,先写单测**)
2. `kernel/Registry.ps1` — 扫描 modules、加载 manifest、参数 schema 校验(纯函数,单测)
3. `kernel/Runner.ps1` — setup/each/teardown、`flow.foreach`/`flow.if`、onError 策略、ledger 断点续跑
4. `kernel/Trace.ps1` + `kernel/Gate.ps1`(ASCII 面板 widget)
5. 25 个 MVP step
6. `kernel/Docs.ps1` — manifest → `catalog.json` / `CATALOG.md`
7. `ebi.ps1` — help / lint / explain / dryrun / run / doctor / trace

**验收**:`ebi lint` 能挑出错参数;`ebi dryrun` 能完整打印执行计划;单测全绿。

### P2 — 对拍验证(1.5 周)

把 `MqSnap.ps1`(673 行)重写成 `workflows/before.transferStatus.capture.json`(约 30 行)
+ `profiles/host-open/`。

**验收(必须在办公 PC 上做)**:
- 同一批 key,新旧两条路各跑一遍:PNG 尺寸/裁剪一致,CSV 标记一致
- 故意造一个 NG 页面,两边都判 NG
- 中途 Ctrl+C,重跑能从断点续上,不重复截图

> 如果 GIFT/GFIX 工作已结束、没有真环境可对拍,就改用任意一条还能跑的旧流程。
> **对拍的目的是验证引擎,不是验证业务。**

这一步会暴露契约的所有错误 —— **P1 的设计不必追求完美,P2 之后重构一次是计划内的**。

### P3 — 新工作实战(2 周,真正的目标)

用 `docs/ebi-dance/INTERVIEW.md` 给下一份工作(Host → Host 或其他)搭第一条流程。

**验收:全程不写 PowerShell,只写 profile JSON + workflow JSON。**
如果做不到,说明抽象漏了 —— 补的 step 记入 catalog,这是正常的成长方式。

### P4 — Excel / 文件组(3 周)

补 `excel.*` / `file.*`,迁 `compose`(组装证据)+ `annotate`(画框)两条流程。
这两条同时用到 excel + screen + profile 的坐标数据,能验证第二类组合。

### P5 — 掩码 + Agent 循环 + OCR 校准(2 周)

`ebi mask`(§7)、`ebi bundle` / `ebi apply`(§8.4)、`ebi calibrate ocr`(§6.2)。

### P6 — 收尾

滚动拼接、HTML 报告、其余流程迁移。每迁完一条,删掉 `legacy/` 里对应的旧脚本。

---

## 12. 需要改造的关键文件(不是照搬)

| 现有文件 | 问题 | 改造 |
|----------|------|------|
| `ProgressLog.ps1` | 事件字段硬编码 `correl_id_s` / `job_name` | 泛化为 `key` + `tags{}`;新增 step 级 trace |
| `MappingStore.ps1` | `Correl_ID_S` / `JOB_NAME` / bitmask 位义写死 | 主键列、变体规则、位定义全部来自 profile |
| `Common.ps1` | `$Global:Timing` / `$Global:Shell` 全局态 | 拆成 `browser.*` / `screen.*` step;时序参数走 step 输入 |
| `SnapVerify.ps1` | 三套 `ConvertFrom-*PageText` 结构类似各写各的 | 抽出 grammar 驱动的 `verify.parse_text`;规则表数据化进 profile。**判定语义一行不改**,靠现有 `Test-SnapVerify.ps1` 单测护住 |
| `Invoke-CropPng` ×4 | 四份重复 | 合并进 `screen.crop` |
| `VerifyConfig.psd1` (863 行) | 工具默认值和项目知识混在一起 | 工具默认留 psd1;项目知识全进 `profiles/<name>/*.json`。保留三层优先级:CLI > 工作目录 overlay > profile > 工具默认 |
| `TimeDigitVerify.ps1` 等 | 3/9 猜测逻辑是负债 | 移入 `legacy/`;思想提炼为 `verify.crosscheck` |

---

## 13. 验收方式

**纯逻辑(可在当前 Linux 环境跑)**
```
powershell -File Tests/Run-Tests.ps1     # 解析检查全部 .ps1 + 纯函数单测
powershell -File Check-Encoding.ps1      # ASCII 源码 + 编码策略
ebi lint  workflows/*.json               # 全部工作流静态校验
ebi dryrun workflows/<x>.json            # 干跑执行计划
ebi mask  check                          # 敏感信息门禁
```

**必须在办公 PC 做(COM / Edge 路径)**
1. `ebi doctor` 全绿
2. P2 对拍:新旧两条路同批 key,PNG 与 CSV 逐项一致
3. 断点续跑:中断后重跑不重复劳动
4. NG 路径:造一个异常页面,`human.gate` 面板正确停下并给出足够上下文
5. `ebi bundle` 产物经 `ebi mask` 后可安全带出办公网

---

## 14. 风险与对策

| 风险 | 对策 |
|------|------|
| **契约一次做不对** | P2 对拍就是为了尽早撞墙;**计划内允许 P2 后重构一次 kernel**,P1 不追求完美 |
| **就地重构动了生产工具** | P0 先打 `freeze/pre-ebi-dance` 冻结标签;旧脚本迁完一条才删一条,任何时刻都能回滚 |
| **JSON 表达力不够,滑向自制编程语言** | 硬规则:workflow JSON **不引入表达式、不引入函数定义**。判定逻辑一律回到 `verify.*` 的 PowerShell 纯函数 + profile 规则表。JSON 只做"连线" |
| **中性词汇表设计不当,新工作套不进去** | P3 就是它的考试。套不进去就改词汇表,**在只有一个 profile 时改是廉价的** |
| **OCR 降级层重新变成负债** | 三条硬规则(§6.1)+ 强制校准闸门;样本集随使用自动增长,不需要专门维护 |
| **PS 5.1 的坑** | 沿用现有铁律:ASCII 源码、无 `param()` 才 dot-source、`Check-Encoding` 入 CI、已知坑写进 `docs/ebi-dance/spec/PS51-PITFALLS.md` |
| **Agent 改坏工作流** | `ebi apply` 强制备份 + `ebi lint` + `ebi explain` 人工复核三道闸;Agent 无权改 step 代码 |
| **办公 PC 无法验证的部分越积越多** | 每阶段末尾定义 ≤10 项的办公 PC 冒烟清单,跑完再开下一阶段 |

---

## 15. 第一步

P0 的第 2 项 —— **写四份规格文档**:
`STEP-CONTRACT.md` / `WORKFLOW-SCHEMA.md` / `PROFILE-SCHEMA.md` / `VOCABULARY.md`,
外加 `INTERVIEW.md`(Agent 访谈剧本)。

全是纯文本,在当前环境就能完成,能和 Agent 反复迭代。一旦定稿,后面 7,000 行
基本是机械劳动。

**在写码之前把契约和词汇吵清楚,是这个项目唯一真正需要"想"的地方。**
而 `INTERVIEW.md` 甚至不用等代码 —— 写完当天就能拿去给下一份工作做访谈。
