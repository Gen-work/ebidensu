# JJODWDBx 稼働確認 —— 实机手册（kadou.collect / kadou.evidence）

> 对象：10/26（月）検証環境里自动稼働（かどう，operation）的 JJODWDBF / JJODWDBH / JJODWDBM，
> 3 job × 3 I/F = 9 条传送（HOST JOD → GFIX → OPEN IDS / IGP2）。
> 只做 OPEN 侧两段：**B = GFIX（GoAnywhere 的 job log）**，**C = OPEN（Jenkins 受信フォルダ + バッチ処理状況一覧）**。
> A 段（HOST 的 JES log、送信データ）是 HOST 团队的事，不做。
>
> **只读（read-only）**：不放置数据、不提交 job、不删不移任何文件。
> 工具对页面做的只有：刷新（F5 / GoAnywhere 的「更新」）、Ctrl+F、打开 job log 并下载、返回、Ctrl+A/Ctrl+C 读文字、截图。
> バッチ処理状況一覧的检索条件**由人输入**，工具不碰表单。
>
> 云端只做过静态检查、单测和 DryRun（空跑）。**这是第一次在办公 PC 上真跑**，每一步都有关卡（gate）。
> 出问题按 `r`（重试）/ `s`（跳过这一条）/ `q`（中止）。

---

## 0. 开工前（只做一次）

1. 取代码（分支名以 PR 上的为准）：

   ```powershell
   git fetch origin claude/jod-kadou-003
   git checkout claude/jod-kadou-003
   powershell -ExecutionPolicy Bypass -File Tests\Run-Tests.ps1
   ```

2. 工作文件夹（WorkDir）用一个新的，例 `C:\Users\<you>\work-JOD-kadou`，下面写成 `$W`：

   ```powershell
   $W = 'C:\Users\<you>\work-JOD-kadou'
   ```

3. 在 `$W` 放 **`ebi.local.json`**（本机覆盖，不进仓库；UTF-8；`\` 写成 `\\`）。URL 用你浏览器地址栏里的真值：

   ```json
   {
     "paths": {
       "goAnywhereUrl":    "https://<GoAnywhere>/goanywhere/...CompletedJobs...",
       "jenkinsRecvIds":   "http://<ci-host>:8011/jenkins_ids/userContent/logs/batch/<batch-host>_appl/IDSVer1/gfix/recv",
       "jenkinsRecvIgp2":  "http://<ci-host>:8013/jenkins_igp2/userContent/logs/batch/<batch-host>_appl/IGP2Ver1/gfix/recv",
       "batchStatusIds":   "http://<bizver-host>/ids9a21/x/xa/IDSXA041.do",
       "batchStatusIgp2":  "http://<bizver-host>/igp29a21/x/xa/IGPXA041.do"
     }
   }
   ```

   - `jenkinsRecv*` 末尾**不要**带 `/`：工具在后面接 `/<文件名>` 去下载受信文件（GFIXデータ 用）。
   - バッチ処理状況一覧 的窗口标题工具没见过，先猜的是「バッチ処理状況」。第一次跑如果 `setup/bs` 报 `not_found`，
     把真标题的一段加进 `ebi.local.json`：`"pages": { "batchStatus": { "windowTitle": "<标题的一段>" } }`，然后 `r`。

4. 自检（只读）：

   ```powershell
   powershell -ExecutionPolicy Bypass -File ebi.ps1 profile check jod-kadou -WorkDir $W
   powershell -ExecutionPolicy Bypass -File ebi.ps1 dryrun workflows\kadou.collect.json -WorkDir $W
   ```

下面省略前缀 `powershell -ExecutionPolicy Bypass -File`。

---

## 1. 收集：`kadou.collect`（18:00 那一轮之后，或 10/27）

B 和 C 都保留历史，所以**一次跑在最后一轮之后**就够了：B 拿今天**全部轮次**，证迹用「B 和 Jenkins 都齐的最早一轮」。

**先开三个 Edge 窗口（各自独立的窗口，不是标签页，K-01）：**

| 窗口 | 打开到哪里 | 备注 |
|---|---|---|
| GoAnywhere | 監査ログ > ジョブログ > 完了したジョブ | 最大化；日付範囲 Today；表示 100 件 |
| Jenkins | 上面 `jenkinsRecvIds` 或 `jenkinsRecvIgp2` 任一个 | 工具会自己切换 IDS / IGP2 |
| バッチ処理状況一覧 | IDS 或 IGP2 任一个 | 工具会切换；**检索你来做** |

```powershell
ebi.ps1 run workflows\kadou.collect.json -WorkDir $W
```

第一次跑会先把 9 条写进 `$W\kadou.csv`（作业清单，worklist）。之后按 **I/F（GoAnywhere project）分 3 组**，每组 3 条：

1. **每组一次 PREPARE**：「プロジェクト名 = JOD_VER1_IF5050_001_SSH ...」
   → 在 GoAnywhere 的プロジェクト名里填这个名字、日付範囲 Today、表示 100 件，列表出来后回车。
   工具会刷新、存列表文字（`log\GoAnywhere\<mmdd>-IF5050.txt`），然后**逐个打开今天这个 project 的 job，下载 job log**
   （`log\GoAnywhere\jobs\<mmdd>\<ジョブ番号>.log`；已下载的不再下）。**这段时间别碰键盘鼠标。**
2. **每条传送**：
   - Jenkins 窗口 F5，找这个相関ID 的文件（20 秒内找不到也继续，C = 未取得）；
   - B 判定（`verify.job_logs`）：今天每一轮的 log 都看 4 行——`successfully uploaded to '/appl/<SYS>/<SYS>Ver1/gfix/recv/<file>'`、
     `return code '0'`、`... was deleted.`、`End Date and Time`，另外核对 Command 行的 shell 和最后的 U / S；
   - 用 Ctrl+F 把证迹那一轮的文件名在 Jenkins 上标成橙色 → 截图 `capture\kadou\<相関ID>\jenkins.png`；
   - 从 Jenkins 下载这个受信文件 → `DATA\GFIX\<相関ID>\<file>`（GFIXデータ 用，只读 GET）；
   - **PREPARE「バッチ処理状況一覧(IDS/IGP2)」**：你输入 相関ID / 処理種別 = データ取込 / 開始日 FROM・TO = 今天（yyyyMMdd），检索，
     结果行（正常終了・処理件数・バッチID・SS）出来后回车 → 工具按 Esc 把窗口拉到前面、截图 `batch.png`、读文字判定
     （画面上有 相関ID、正常終了、バッチID，没有 異常終了）；
   - **GATE**：B 每一轮的结果（一轮一行）、红字 = 自动判定有疑问的地方、三个证据文件。看一眼 Jenkins / バッチ 的截图，
     Enter = ok，`n` = ng，`m 备注` 可以留话。
3. 某条今天一轮都没有成功（没有 job log，或者全部 ng）→ 出一个 ng 关卡说明原因，记 `collect = ng` 然后跳到下一条。

看总表（9 行 × B / C）：

```powershell
ebi.ps1 status -Profile jod-kadou -WorkDir $W
```

`B` = GoAnywhere（ok / ng / unknown＝未取得），`C_JK` = Jenkins 上找到文件（True / False），`C_BS` = バッチ処理状況一覧。
`RUNS` 列是今天的轮次「ok 数 / 轮数」，`FILE` 是证迹用的那一轮的文件名。件数（附件 N 列）只作参考，不判定。

中途停了：`ebi.ps1 run workflows\kadou.collect.json -WorkDir $W -Resume`。只重做一条：`-Only JIGPS87S`。

---

## 2. 证迹簿：`kadou.evidence`

```powershell
ebi.ps1 run workflows\kadou.evidence.json -WorkDir $W
```

每个 job 一本：`$W\evidence\J4検証資料(REQ-000xxxxx_GIFT廃止対応)_<JOB>_稼働確認.xlsx`。
第一次从空白模板（`profiles\jod-kadou\assets\kadou-template.xlsx`）复制，以后在同一本里改写。
**只有两个 sheet**：GFIX受信結果、GFIXデータ。送信側的 sheet 是 HOST 团队的；合进哪本最终文件由你决定。

布局照竹村さん 10/22 版的参考资料（DBH）：

| 行 | GFIX受信結果 |
|---|---|
| B10 / B105 / B200 | 相関ID（IDSU / IGPS / IGPU），下一行是 バッチ処理状況一覧 的截图 |
| B54 / B149 / B244 | **▼GFIXログ**（粗体），下一行起是 job log 全文；`Command:` 那一行涂黄 |
| B296 | Jenkinsの受信フォルダ |
| B297 / B341 / B385 | 相関ID，下一行是 Jenkins 受信フォルダ的截图（目标文件橙色） |

GFIXデータ：B3 / B63 / B123 = 相関ID，下面是受信文件的内容（**文字**，从 Jenkins 取的那个文件）。
今天没有 GIFT 数据，所以不做对比。

每本最后一个 GATE：在 Excel 里看一遍（清单在面板上），回车 = 保存并关闭。

**手工要做的：** GFIX受信結果 **B3 的图**（附件 9 行表中这个 job 的 3 行）自己贴。

---

## 3. 已知的不确定点（第一次跑会知道答案）

- **バッチ処理状況一覧 的窗口标题**是猜的（见 0.3）。检索是人做的，以后可以把表单的 Tab 顺序记下来再自动化。
- **Jenkins 受信フォルダ的页面文字**：工具只要求页面上出现文件名；目录很长时 Ctrl+F 会滚到那一行。
- **截图大小**：窗口按 1100 × 680 截（`profile window.json`），这样一张图放得进参考资料给它的 43 行。显示不全就在
  `ebi.local.json` 里改 `"window": { "width": ..., "height": ... }`（高度别超过约 680，否则会盖住下面的 ▼GFIXログ）。
- **GFIXデータ 用文字，不是截图**：主线默认写的是「Jenkins『参照』的画面」。文字内容相同、不受窗口大小影响；
  如果要截图，告诉雨会。文件行数超过 58 行会碰到下一条的位置，关卡会提示。
- **GoAnywhere 一个 project 今天的行**：3 job × 6 轮 = 18 行，表示 100 件够。job log 保留天数没有实测记录
  （gfixRecv 取到过几天前的），所以最好当天或第二天跑。
