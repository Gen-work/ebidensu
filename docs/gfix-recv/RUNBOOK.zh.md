# GFIX 受信 —— 第一次实战手册（ebi-dance 首测）

> 对象：在办公 PC（Windows + PowerShell 5.1 + Excel）上第一次跑 `gfixRecv.*` 四条工作流（workflow）。
> 云端只做过静态检查（static check）、纯逻辑单测和 DryRun（空跑）；**COM / SendKeys / 截图这一层是第一次真跑**，
> 所以每一步都留了人工确认的关卡（gate）。出问题就按 `r`（重试）/ `s`（跳过这条）/ `q`（中止）。

---

## 0. 开工前（约 10 分钟，只做一次）

1. 取代码：

   ```powershell
   git fetch origin claude/gift-test-tracking-ncqisu
   git checkout claude/gift-test-tracking-ncqisu
   ```

2. 跑测试和环境检查（doctor）：

   ```powershell
   powershell -ExecutionPolicy Bypass -File Tests\Run-Tests.ps1
   powershell -ExecutionPolicy Bypass -File ebi.ps1 doctor
   ```

   云端（Linux）上已知会失败的 2 条 `EvidencePlan: ... X\DF\A.png` 是路径分隔符（`\` 和 `/`）的差异，在 Windows 上应该会通过。

3. 在工作目录（WorkDir，例 `C:\Users\<you>\work-BIX`）放一个 **`ebi.local.json`**：本机覆盖（overlay），
   **不进仓库**。仓库里的 `profiles/gfix-recv/paths.json` 只有占位值。模板：

   ```json
   {
     "paths": {
       "bix": "<共享盘>\\...\\REQ-000xxxxx_GIFT廃止対応\\40.J4\\08.BIX",
       "wbs": "<共享盘>\\...\\REQ-000xxxxx_GIFT廃止対応\\00.管理\\【GIFT廃止対応】WBS.xlsx",
       "mapping": "mapping.xlsx",
       "dfExe": "C:\\tools\\DF\\DF.exe",
       "goAnywhereUrl": "https://<GoAnywhere>/goanywhere/jobs/CompletedJobs.xhtml?logs=true",
       "jenkinsReportUrl": "http://<jenkins>/jenkins_jpc/userContent/logs/files/JPC/Ver1/data/report/",
       "receiveLogUrl": "http://<jenkins>/jenkins_jpc/userContent/logs/files/JPC/Ver1/batchlogs/GFIXReceive.log",
       "receiveUnzipLogUrl": "http://<jenkins>/jenkins_jpc/userContent/logs/files/JPC/Ver1/batchlogs/GFIXReceiveUnzip.log"
     }
   }
   ```

   JSON 里的 `\` 要写成 `\\`。填好真实值的版本，聊天里会直接给你。

4. 确认 `mapping.xlsx` 在 WorkDir 里，并且 O 列（GFIX実行日）和 P 列（GFIX TIME）是今天的值。

5. profile 自检：

   ```powershell
   powershell -ExecutionPolicy Bypass -File ebi.ps1 profile check gfix-recv -WorkDir C:\Users\<you>\work-BIX
   ```

下面的命令都省略了前缀 `powershell -ExecutionPolicy Bypass -File`，`-WorkDir` 也写成 `$W`
（可以先执行 `$W = 'C:\Users\<you>\work-BIX'`）。

---

## 1. 早上：`gfixRecv.plan`（今天做哪些）

```powershell
ebi.ps1 run workflows\gfixRecv.plan.json -WorkDir $W
```

- 只读打开 WBS，筛出：`B=受信`、`G=修正後実施`、`P` 是「厳」/「厳←○○」/「○○→厳」、`AF`（最新計画 開始）= 今天。
  箭头规则和旧的 `OwnerFilter.ps1` 一样：「ニンヌ←厳」**不算**你的。
- **作业清单只靠 WBS 就能建**：W 名 = J 名的第 5 个字母 J → W（SJDSJM40 → SJDSWM40）。
  不合这个规则的名字会照抄并给出警告（warning）。
- mapping.xlsx 只是补充：O 列（GFIX実行日）= 今天的行，会补上 GFIX 时刻和文件数；
  只在 mapping 里、不在 WBS 里的 job（例：WBS 写「ニンヌ」的 MJDSJM40）也从这里进清单。
  没补 mapping 也没关系：文件数显示 `-1`（不知道），时刻在 track 时从 Teams 消息里取。
- WBS 和 mapping 的差别只**显示**（`setup/cmp` 那一行），不再停下来问。
- 结果：`$W\gfixrecv.csv`（作业清单，用 Excel 打开也行）。用你的 WBS 和 mapping 模拟过：
  F/K/L/M/Q/RJDSWM40，6 个 job、时刻和 p1 完全一致。

---

## 2. 白天：`gfixRecv.track`（每个 HOST job 一轮）

**窗口准备**（一次）：

- GoAnywhere：InPrivate 窗口，「完了したジョブ」，日付範囲 Today、プロジェクト名 JDS、開始時間 ▼。
  **最大化**，显示器 1920×1080、缩放 100%。截给 Teams 的那块图是按这个尺寸量的坐标。
- Jenkins data/report：另开一个**普通窗口**（不要和 GoAnywhere 放同一个窗口的不同标签页）。工具会把它调成 1410×807。

```powershell
ebi.ps1 run workflows\gfixRecv.track.json -WorkDir $W
# 只跑一个：  -Only RJDSWM40
```

每个 job（按作业清单的顺序；顺序和领导的预定表不一样时，用 `-Only <W名>` 一个一个跑最省事）：

**作业清单 `gfixrecv.csv` 就是「表」**：`GFIX_DATE`（日期）和 `GFIX_TIME`（开始时刻）两列决定每个 job 用哪一天、哪个时刻去找。
- 已经知道时刻（今天全部跑完了 / 重跑 / 补以前的某一天）：用记事本打开 `gfixrecv.csv`，把 `GFIX_TIME` 填上
  （`10:30:25` 或 `10:30` 都行；以前的日期就把 `GFIX_DATE` 也改掉，`track` 列清空），然后 PASTE 面板直接按 Enter。
  GoAnywhere 页面的日期范围要你自己改成那一天。
- 实时跑的时候，用 Teams 消息决定的开始时刻会**自动写回** `GFIX_TIME`，所以出错后重跑直接 Enter 就是同一个时刻。
- GoAnywhere 的「プロジェクト名」：JDS 和 JMR 的 job 混在一起跑时，把它清空（显示全部），或者每个 job 前切换。
- 开始前把**以前打开的 DF 窗口关掉**（工具也不会再截旧窗口，但关掉最省事）。
- 工具在等页面时（`browser.wait_for`），点一下控制台按 **q 或 Esc** 就能停下来，不用 Ctrl+C（那会把整个运行杀掉）。
  页面跳到别的地方时会立刻停下并显示 `wrong_page`：把页面手动回到「完了したジョブ」再按 `r`。
- GIFT 侧文件夹：先找 `DATA\GIFT\<J名>`（例 JJMRJE6F），没有就找 `<W名>`（JJMRWE6F）。
  两个都没有：工具会**建好 `DATA\GIFT\<J名>` 文件夹并停下来**，你把 GIFT 的文件放进去再按 `r`。文件夹在但是空的也一样。
- PASTE 面板的 `k` = 不用消息继续：表里有时刻就用表里的，没有才用现在的时刻。只有「现在」或你打的时刻才会写回表里。
- 一个 job 显示 ok 的意思是：GIFT 文件和下载的 GFIX 文件**逐行比较完全一致**。DF 截图只是证据图片，
  不参与判断；截图错了就把那一行的 `track` 清空重跑。

1. **PASTE 面板**：Teams 里出现「ジョブ:XXXを実施します。(送信予定:n件)」时，**复制这条消息**（Ctrl+C），
   回到控制台按 Enter。**按 Enter 的时刻 = 开始时刻**（GoAnywhere 的行按「开始 −2 ～ +13 分」来找）。
   - 这一行已经有 `GFIX_TIME`：Enter = 用表里的时刻（不需要消息）；`n` = 用现在的时刻（要先复制消息）。
   - 复制晚了：不要按 Enter，打领导预定表上的时刻，比如 `10:30`。
   - 消息是别的 job 的：面板会写出两个名字，等你复制对的那条；`s` = 先跳过这个 job（留到下次），`k` = 不用消息、时刻 = 现在。
2. **PREPARE 面板**：「Teams の『ジョブ:XXXが正常終了しました。』が来たら Enter」。
   reaction 你自己在 Teams 点。
3. 自动：在 GoAnywhere 上执行 Ctrl+F「フィルタリングする」→ Esc → Tab×5 → Enter（**更新**，不按 F5），
   直到出现开始时刻 −2 分 ～ +13 分之间任意一分钟的行 → 读文本 → 挑出这个 job 的行 → Receive 的 job 号。
   读完文本后会**点一下页面空白处**（`pages.goAnywhere.deselectAt`，默认 1700,900），取消 Ctrl+A 留下的全选
   （Edge 里按 Esc 取消不了），不然截图会是一片蓝。
4. 自动：截取 GoAnywhere「上面那部分」（面板标题到这个 job 的行，就是 p4 的范围）。
5. 自动：Jenkins 按 F5 → Ctrl+End → 读文本 → 找出在 Receive 时间段里放进 report 的 `F…csv`
   → 用 HTTP 直接下载到 `DATA\GFIX\<W名>\`（文件夹不存在会自动建）。
6. 自动：列出 `DATA\GIFT\<J名>\*.csv` → 按「行数相同 + 到达顺序」配对 → 逐行比较内容。
   配不上时会停下来问你（**不猜**）。
7. 自动：每一对文件启动一次 DF.exe（窗口 1133×429，一屏正好 16 行）→ Ctrl+Home 截图；
   超过 16 行的再 Ctrl+End 截一张 → 关掉 DF。
8. **GATE**：内容一致就直接过；不一致或不确定会问你（`Enter` = ok / `n` = ng / `s` = 先跳过）。
9. **SHARE 面板**：文字「OPEN側で確認できました。前後一致です。」+ GoAnywhere 图 + DF 图已经放进剪贴板
   → 点 Teams 的输入框 → **Ctrl+V** → 看一眼 → 发送 → 回到控制台按 Enter。
   多个文件时，DF 图会按顺序全部带上。
10. 记录：`capture\gfix\<W名>\track.json`（后面 logs / evidence 用），作业清单 `track = ok`。

**第一次跑请特别看这几处**（都在 `capture\gfix\<W名>\` 里）：

| 看什么 | 文件 | 不对的话 |
|---|---|---|
| GoAnywhere 截图范围是不是 p4 那样 | `goanywhere.png` | `ebi.local.json` 里覆盖 `pages.goAnywhere.crop`（left/right/top/firstRowTop/rowHeight） |
| 「更新」按钮配方有没有点到 | 控制台：超时（timeout）= 没刷新出来 | `pages.goAnywhere.refreshRecipe`（Tab 的次数） |
| DF 是不是一屏 16 行 | `df\*__first.png` | `layout.df.width/height` |
| Teams 粘贴时图片在不在 | Teams | 把 `layout.share.mode` 改成 `"sequence"`（文字、图片分几次贴） |
| 截图里有没有残留的蓝色全选 | `goanywhere.png`、证据里的 Jenkins 图 | `deselectAt` 换一个空白位置（窗口内像素） |
| DF 有没有被拉起来并截到 | 控制台 `no_window` | DF 是启动器型程序的话会按标题 `DF - ` 找窗口（`layout.df.windowTitle`） |

`ebi.local.json` 覆盖写法示例（深度合并，只写要改的键）：

```json
{ "layout": { "share": { "mode": "sequence" } },
  "pages": { "goAnywhere": { "crop": { "top": 140 } } } }
```

---

## 3. 傍晚（当天！）：`gfixRecv.logs`

GFIXReceive.log 是固定链接，**第二天旧文件就被移到 bk 文件夹**，所以要当天跑。

```powershell
ebi.ps1 run workflows\gfixRecv.logs.json -WorkDir $W
```

- `log\GFIXReceive\mmdd.utf8.log`（原样）→ `mmdd.log`（SJIS）。Unzip 日志不存在（404）也没关系。
  - 日志里那个 SJIS 的全角冒号「：」（`update count：1`），以前贴进 Excel 会变成 `_xDC81_F`；
    现在会正确解码成「：」。这和你以前的证据不完全一样，需要你判断要不要保留原来的样子。
- `log\GoAnywhere\mmdd.txt`、`log\JenkinsOverview\mmdd.txt`：两边一览的全文。
- 每个 Receive job 号：Ctrl+F 号码 → Enter → Ctrl+F「ジョブログをダウンロードする」→ Tab×2 → Enter
  → 等 Downloads 里出现文件 → 移到 `log\GFIX受信ログ\<号码>.log` → Alt+← 返回。
  **这一段的按键配方是按你的描述写的，第一次请盯着看**（`pages.goAnywhere.jobLogRecipe`）。
- 从 job log 里读出转送文件名（`JJPCRS12…`），写进 `capture\gfix\<W名>\logs.json`。

---

## 4. 证据：`gfixRecv.evidence`

```powershell
ebi.ps1 run workflows\gfixRecv.evidence.json -WorkDir $W -Only LJDSWM40
```

**先拿一本复制件试**（把 `paths.bix` 临时指向一个本地复制文件夹）。插图这一步不可重复执行（not idempotent），
跑两次会插两遍。

GFIX受信結果（按样本 QJDSWM39 的版式）：

- B3：GIFT受信結果 的第 1 张图（Excel 截图）的复制。
- 下一行：`GFIX受信ログ` + job log 全文（多个 job log 之间不空行），黄色标记 `successfully uploaded to …` 和 `Command: …`。
  宽度 = 文字半角宽度 ÷ 3 列（用样本的 4 处黄色标记校准过，结果一致）。
- 空一行：`GFIXReceive.log` + 这次转送的 START…END 块（含 END 后面的日期行），黄色标记 `FileGIFTGetBLBean.execute() fileName` 和 `file stored :`。
- 空一行：Jenkins 截图（F5 → Ctrl+End）+ 红框框住收到的文件行。行的位置是按蓝色链接文字的行，对照一览文本**从最后一行往上数**找到的；
  用样本截图验证过（样本红框 481px，算出来 482px）。

GIFTデータvsGFIXデータ：

- B3 = 送信データ!A3（JJPCRS12）。
- B4 起：DF 截图。超过 16 行：开头截图 → Z 列放波浪线（p5）→ 结尾截图；
  红框框住**最后一行的行号**（在截图里找行号那一列最后一行墨迹）和右下角「同一内容」。
  （和样本对比：行号框 213px vs 212px。）
- 多个文件：按配对顺序往下排，中间空一行（**这是推断**，样本里没有多文件的 GIFTvsGFIX）。

最后：每张 sheet 选中 A1 并滚到顶部（等同 Ctrl+Home）、取消复制模式（Esc）、回到第 1 张 sheet、保存；
检查 GFIX受信結果 B 列的字体是不是 ＭＳ ゴシック 10（只报告，不改）。
然后 **GATE**：你在 Excel 里看一遍，`Enter` = ok，作业清单 `evidence = ok`。

---

## 5. 出问题时

- 中途断了：同样的命令加 `-Resume`（续跑上一次的 run，已经做完的步骤不会重做）。
- 想重做某个 job：在 `gfixrecv.csv` 里把对应的 `track` / `logs` / `evidence` 清空。
  同一个 job 以后换了日期再次出现在计划里时，plan 会自动把这三列清空（`resetOnChange`）。
- 所有中间产物：`capture\gfix\<W名>\`（截图、页面文本、track.json、logs.json、receive.txt）。
- 每次运行的记录：`run\<runId>\`（trace、ledger）。出了问题把这个文件夹的内容发给我就能看。

---

## 6. 我做的假设（猜错就要改的）

| 假设 | 依据 | 改哪里 |
|---|---|---|
| GIFT 侧文件夹 = J 名（RJDSJM40），GFIX 侧 = W 名（RJDSWM40） | 你的描述 | workflow 里的 `DATA/GIFT/{{item.JOB}}` |
| 一个 job 的文件数 = mapping.xlsx 里这个 job 的行数 | mapping 里 JJDSJM51 有 2 行 | 不一致只是警告（warning） |
| 「今天的」= WBS 的 AF 列（最新計画 開始） | 修正後実施 那行 AF=AG | `gfixRecv.plan.json` 的 columns |
| GoAnywhere 的行按开始时间从新到旧排 | p3 的 ▼ | 不是的话截图会多带别的行（会出 warning） |
| Jenkins 的文件时间 = 落盘时间，落在 Receive 开始 −2 秒 ～ 结束 +10 秒之间 | 1008 的实际数据（13:45:24–27 → 13:45:27） | `padBeforeSec` / `padAfterSec` |
| 「前後一致」= 两个文件逐行相同（忽略 CRLF/LF 差异） | DF 的「同一内容」 | `file.compare` |

---

## 7. 这次新加的东西（给以后的自己）

- step：从 37 个增加到 69 个。excel 组：ensure_app / open / close / quit_app / read_rows / write_cell / write_lines / highlight /
  insert_pictures / copy_picture / stack_plan / tidy；file 组：download / wait_for_download / move / convert_encoding /
  read_text / list / extract_blocks / compare；screen 组：find_window / row_region / find_ink_rows / list_rects / launch_capture；
  verify 组：filter_records / time_window / pair_files / compare_sets；human.share；browser.download_each；table.upsert。
  另外 browser.wait_for 加了 `refreshRecipe`，screen.capture_window 加了 `bounds`。
- kernel：LogText（混合编码、日志块、高亮宽度、时间窗、配对）、Layout（几何）、RichClip（剪贴板 HTML/RTF）、Excel（COM 小工具）；
  `run.date / dateSlash / mmdd / toolDir`；profile 多了 `paths.json`。
- 测试：`Tests\Test-LogText.ps1`、`Tests\Test-GfixRecvSteps.ps1`（用你的真实样本把整条数据链跑一遍）。
