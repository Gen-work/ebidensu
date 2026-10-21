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
     "operator": { "owner": "厳", "reviewer": "<确认者的姓>" },
     "paths": {
       "bix": "<共享盘>\\...\\REQ-000xxxxx_GIFT廃止対応\\40.J4\\08.BIX",
       "wbs": "<共享盘>\\...\\REQ-000xxxxx_GIFT廃止対応\\00.管理\\【GIFT廃止対応】WBS.xlsx",
       "mapping": "mapping.xlsx",
       "reviewSheet": "<共享盘>\\...\\REQ-000xxxxx_GIFT廃止対応\\00.管理\\<レビューチェックシート>_OPEN.xlsx",
       "mailTemplate": "review-mail.json",
       "dfExe": "C:\\tools\\DF\\DF.exe",
       "goAnywhereUrl": "https://<GoAnywhere>/goanywhere/jobs/CompletedJobs.xhtml?logs=true",
       "jenkinsReportUrl": "http://<jenkins>/jenkins_jpc/userContent/logs/files/JPC/Ver1/data/report/",
       "receiveLogUrl": "http://<jenkins>/jenkins_jpc/userContent/logs/files/JPC/Ver1/batchlogs/GFIXReceive.log",
       "receiveUnzipLogUrl": "http://<jenkins>/jenkins_jpc/userContent/logs/files/JPC/Ver1/batchlogs/GFIXReceiveUnzip.log"
     }
   }
   ```

   JSON 里的 `\` 要写成 `\\`。填好真实值的版本，聊天里会直接给你。
   **`"operator"` 是顶层的一个键**，不要把 `{ "operator": … }` 整个塞进别的对象里（10/15 这样写过，plan 读不到 owner，
   全员的 job 都进了清单）。存成 **UTF-8**（记事本「名前を付けて保存」→ 文字コード UTF-8；有没有 BOM 都能读）。
   `operator.owner` 是 WBS 担当列里你的名字（规则同 OwnerFilter：「厳」「厳←○○」「○○→厳」算你的）；
   **留空 = 不按担当过滤，当天所有人的 job 都进清单**（plan 开头会提示一句）。仓库里的 `profiles/gfix-recv/operator.json` 是空的，名字只写在本机。

4. 确认 `mapping.xlsx` 在 WorkDir 里，并且 O 列（GFIX実行日）和 P 列（GFIX TIME）是今天的值。

5. profile 自检：

   ```powershell
   powershell -ExecutionPolicy Bypass -File ebi.ps1 profile check gfix-recv -WorkDir C:\Users\<you>\work-BIX
   ```

下面的命令都省略了前缀 `powershell -ExecutionPolicy Bypass -File`，`-WorkDir` 也写成 `$W`
（可以先执行 `$W = 'C:\Users\<you>\work-BIX'`）。

---

## 0.5 入口：`ebi.cmd` / `ebi.ps1 menu` / `ebi.ps1 status`（v2.25.0 起）

下面第 1～4 节的四条 `run` 命令仍然是正本，但平时不用手打：

- **看盘**：`ebi.ps1 status -WorkDir $W`。一屏：**今天的行**（plan 给了今天的 `WORKDAY`）每本一行，JOB / GFIX_DATE /
  GFIX_TIME / FileCount，track / logs / evidence / deliver 四列判定（`ok` / `NG` / `-` 未做），每列合计；
  表下面另外几行说不在表里的行（v2.29.0 起）：`on hold:`（保留中，带理由）、`earlier days, not finished:`
  （往日没做完的，带做到哪一步）、`planned for a later day:`、`earlier days, all ok: N row(s)`（只计数）。
  然后是没跑完的 run（`-Resume` 用的 id）、NEXT（先续跑；**今天还没有一行时先 plan**；再清 NG、再是最早有行在等的阶段）
  和「下一条该打的命令」。只读，不碰 CSV。
  每段（track / logs / evidence / deliver）**只拿今天的行**，保留中的行永远不拿；往日的行要么在 plan 里选进今天，
  要么开跑前的面板按 `t`（见下），要么 `-AllDays`。所以菜单里直接回车就是「今天的这几本」，不会再给旧行做草稿。
- **入口**：仓库目录里双击 `ebi.cmd`。第一次（还没记住工作文件夹时）它会问「Which work folder?」，
  把工作文件夹**拖进窗口**再回车即可；也可以把文件夹直接拖到 `ebi.cmd` 图标上，或 `ebi.cmd C:\Users\<you>\work-BIX`。
  之后记住在 `.ebi\menu.json`，或 `ebi.ps1 menu -WorkDir $W`。
  `ebi.local.json` 写坏了（少逗号、路径里单个 `\`）时会直接说是哪个文件、哪里坏了，不会再报「no profile has its worklist」。先显示上面的盘，再选：
  **Enter = 推荐的下一步**（菜单第一行 `Enter  run track for today's rows` 写着是什么）。
  数字 = plan / track / logs / evidence / deliver（plan 之后的几段会再问一次 Excel_NAME：
  回车 = 今天的行，`A,B` = 几本，**名字的一部分也行**：`DBC` = 所有带 DBC 的；`a` = 所有日子的待做行 `-AllDays`），
  `r` = 续跑没跑完的 run，`k` = 自检（表上 ok 但文件没了、GIFT 文件夹空了、已交付但前面的列空着；`w` 再对一遍 WBS 的实绩，见 KNOWLEDGE K-26），
  `m`（mark）= **手动标 ok**：在工具外面做完了的（手工补的证据、别人代做的），选行（名字或一部分）→ 选列（`1` track `2` logs …）→ Enter；
  `c`（clear）= 清掉判定让那段重做；
  `h`（hold，保留）= 把行**保留**（等 HOST 修复之类，理由默认 `waiting for HOST`）；对已经保留的行再按 `h` = **解除**；
  `s` = 重画，`d` = doctor，`q` = 退出。
  **开跑前的对账**（reconcile／突き合わせ，v2.29.0 起）：选了某一段之后、真正跑之前，菜单先看两件事——
  ① **看起来已经做完的行**（`LOOKS DONE ALREADY`）：这一段的判定是空的，但这一段自己的记录文件已经在
  （`capture\gfix\<W名>\track.json` / `logs.json` / `gate.json`，而且是这一行的日子当天或以后写的）。
  多半是上次跑到一半断了、判定没写上。`Enter` = 标 ok、这次跳过；`n` = 再跑一遍。
  ② **往日没做完的行**（`EARLIER DAYS, NOT FINISHED`）：`t` = 也算今天（`WORKDAY` 改成今天，GFIX 跑的日期不变，logs / evidence 接着做），
  `h` = 保留，`Enter` = 不管（这次只跑今天的）。
  选完它会**把那条 `ebi.ps1 run ...` 原样打印出来再执行**（子进程，输出和 GATE 的提问都实时显示在同一个窗口），
  跑完印一行结果（done / stopped with a failure / cancelled）再回到菜单。
  （v2.26.1 以前子进程的输出被整段吞掉、最后挤成一行，GATE 的提问看不见：track / deliver 请用 v2.26.2 以后的版本从菜单跑。）
  选了前置没做完的阶段（比如 logs 还没 ok 就选 evidence）会先 `[WARN]` 提示，
  workflow 本身也会拒绝（`file_not_found`）。
- 只有一个 profile 的作业清单在 WorkDir 里时不用说 `-Profile`；有几个时加 `-Profile gfix-recv`。
- README 里有脱敏数据录的截图和 GIF（「A receive day in one screen」）。

---

## 1. 早上：`gfixRecv.plan`（今天做哪些）

```powershell
ebi.ps1 run workflows\gfixRecv.plan.json -WorkDir $W
```

- 只读打开 WBS，读出所有 `B=受信`、`G=修正後実施` 的行：`P` 担当、`AF`／`AG`（最新計画 開始／終了）、
  `AJ`（実績 開始：有日期 = 已经有人做完）。owner 来自 `ebi.local.json` 的 `operator.owner`；
  箭头规则和旧的 `OwnerFilter.ps1` 一样：「ニンヌ←厳」**不算**你的。owner 留空 = 所有人的行都算你的。
- **分成几个候选集合**（candidate set，v2.29.0 起，`table.day_sets`）：
  - `today`：计划日包含今天（AF ≤ 今天 ≤ AG）、还没做；还有今天早些时候已经选进来的行
  - `late`：计划在今天以前、还没做（往回看 21 天，`-Var backDays=…`）
  - `ahead MM/dd`：接下来 2 个工作日的（提前做，`-Var aheadDays=…`）
  - `others, today`：别人今天的（设了 owner 时）
  - `mapping.xlsx only`：mapping.xlsx 里今天有、上面都没有的
  - `on hold`：作业清单里保留中的（不能选；菜单 `h` 解除）
  **已经做完的不进集合**：以**作业清单**为准——四列全 ok（或最后的 deliver ok）的只在面板上面列一行「done here」。
  WBS 的 AJ（実績 開始）**只算参考**：AJ 是人手填的，可能早上开工前就填了（10/20 的 M/SJODJSPQ）、也可能填错（10/16 的 K/LJODJDBC）。
  所以 AJ 有日期、但作业清单里没做完（或 AJ 就是今天）的**照样列出来**，后面注一句 `WBS AJ 10/16`，面板上面也列一行「WBS AJ filled, not done here」。
  只有作业清单里从来没有这本、AJ 又是今天以前的，才当作「工具以外做完了」不列（「done by the WBS actual start」那一行）。
- **选**（`human.pick`，两层）：
  第一层是**推荐的组合**：`1` = 今天的（推荐，**直接回车**）、`2` = 今天 + late、`3` = 今天 + 下一个工作日的 ahead。
  也可以**直接打名字或名字的一部分**：`DBC` = 所有带 DBC 的、`2 -SJODWDBC` = 2 去掉一本、`1 FJODWWBV` = 今天的再加一本；全角也行。
  `m` = **第二层**：所有集合编号列出（每本后面有说明：since 10/15、track ok、mapping 10:30 …），数字 = 那个集合，`b` = 回第一层。
  `h 名字` = 把它们**保留**（理由默认 `HOST 待ち`），保留之后再选今天的。`0` = 今天什么都不做，`q` = 停。
  打完会**复述一遍**（`-> 4 jobs today: …`），`Enter` = 对，打别的 = 换一个答案。
  **已经知道要做哪几本时**：`ebi.ps1 run workflows\gfixRecv.plan.json -WorkDir $W -Only MJODWSPQ,QJODWSPP,SJODWSPQ`
  —— `-Only` 就是这个面板的答案（不管它们在哪个集合：今天、遅れ、先取り、别人的），只复述一遍等你按 Enter。
  （v2.28.0 以前 plan 不认 `-Only`：只拿 AF = 今天、担当是你的行，所以 10/21 的、或者 WBS 担当写着别人的那几本，plan 不会放进清单，track 的 `-Only` 就说 not in the worklist。）
- 选中的行写进作业清单：`WORKDAY` = 今天；`GFIX_DATE` = 今天，**track 已经 ok 的行保持原来的日子**（logs / evidence 接着那天的做）。
  `GFIX_DATE` 变了的行 track / logs / evidence / **deliver** 清空（重跑）。保留的行写 `hold` 列。
- 跑完表格下面直接印 **next:** 下一条命令（track / logs / evidence 也一样），不用翻手册。
- 开头出现 UNFINISHED RUN 面板（以前某天没跑完的 run）：`d` = 把它们标成 abandoned、以后不再问（`run\<id>\` 文件夹还在），然后开新 run；`n` 只是这次跳过、下次还会问。
- **作业清单只靠 WBS 就能建**：W 名 = J 名的第 5 个字母 J → W（SJDSJM40 → SJDSWM40）。
  不合这个规则的名字会照抄并给出警告（warning）。
- mapping.xlsx 只是补充：O 列（GFIX実行日）= 今天的行，给选中的行补上 GFIX 时刻和文件数；
  只在 mapping 里、不在 WBS 里的 job 在 `mapping.xlsx only` 集合里（owner 留空时它在推荐里）。
  没补 mapping 也没关系：文件数显示 `-1`（不知道），时刻在 track 时从 Teams 消息里取。
- 选中的行还会从 mapping.xlsx（所有行，不只今天的）带上 `GIFT_DATE`（M 列 GIFT実行日）和 `GIFT_TIME`（N 列 GIFT TIME），下一段 gift 用。
- 结果：`$W\gfixrecv.csv`（作业清单，用 Excel 打开也行）。跑完的 next 是 gift。

---

## 1b. plan 之后：`gfixRecv.gift`（GIFT 文件在不在，v2.29.0 起）

```powershell
ebi.ps1 run workflows\gfixRecv.gift.json -WorkDir $W
```

菜单里 plan 跑完直接回车就是这一段（`2  gift`）。track 要用的 GIFT 文件（`DATA\GIFT\<J名>`，没有就 `<W名>`）在这里先查好：

- **文件夹里有文件** → `gift = ok`，什么都不问。
- **没有** → 读 Jenkins data/report 一览（就是 track 用的那个窗口，要先开着；没开会停下，开了按 `r`），
  用 plan 从 mapping.xlsx 带过来的 **GIFT実行日（M 列）/ GIFT TIME（N 列）**，找那个时刻 −2～+13 分钟里放进 report 的文件：
  - 窗口里**正好是该有的个数**（mapping 的文件数，不知道就当 1 个）→ 不问，直接下载到 `DATA\GIFT\<J名>\`，`gift = ok`。
  - 个数不对，或者 GIFT TIME 是空的 / `-` → 出选择面板（和 plan 的一样）：`1` = 时刻窗口里的那几个（回车），
    后面是那一天其余的文件**按小时分组**（`2` = 09 点那一组 …），每个文件后面写着时刻；也可以打文件名的一部分（`0026`）。
    回车前会复述一遍要下载哪几个。`0` = 都不要（这一本保持 pending，自己放文件）。
  - GIFT実行日 是空的 → 停下来说「no GIFT run day」：在 mapping.xlsx 补上日期按 `r`，或者自己把文件放进去按 `s`。
- 下载完 track 照常从 `DATA\GIFT\<J名>` 取文件、逐行比较；放错了文件，比较会是 ng，照旧会问你。

---

## 2. 白天：`gfixRecv.track`（每个 HOST job 一轮）

**窗口准备**（一次）：

- GoAnywhere：InPrivate 窗口，「完了したジョブ」，日付範囲 Today、プロジェクト名 JDS、開始時間 ▼。
  **最大化**，显示器 1920×1080、缩放 100%。截给 Teams 的那块图是按这个尺寸量的坐标。
  **列里加上「解決済み」**（詳細 ▼ 的列选择）：ステータス 那个红叉图标复制不进文字，但 解決済み 列在异常终了的行上是 `No`、其他行空白，
  工具靠它判断 job 有没有在 OPEN 侧失败。没加这一列的页面会被当成「不是那个页面」停下（wrong_page），加上再 `r`。
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
- DF：每对文件启动前，工具会**先关掉所有 DF 窗口**，屏幕上只剩新开的那个再截图。
- 工具在等页面时（`browser.wait_for`），点一下控制台按 **q 或 Esc** 就能停下来，不用 Ctrl+C（那会把整个运行杀掉）。
  页面跳到别的地方时会立刻停下并显示 `wrong_page`：把页面手动回到「完了したジョブ」再按 `r`。
- GIFT 侧文件夹：先找 `DATA\GIFT\<J名>`（例 JJMRJE6F），没有就找 `<W名>`（JJMRWE6F）。
  两个都没有：工具会**建好 `DATA\GIFT\<J名>` 文件夹并停下来**，你把 GIFT 的文件放进去再按 `r`。文件夹在但是空的也一样。
- PASTE 面板的 `k` = 不用消息继续：表里有时刻就用表里的，没有才用现在的时刻。只有「现在」或你打的时刻才会写回表里。
- 一个 job 显示 ok 的意思是：GIFT 文件和下载的 GFIX 文件**逐行比较完全一致**。DF 截图只是证据图片，
  不参与判断；截图错了就把那一行的 `track` 清空重跑。

1. **PASTE 面板**：Teams 里出现「ジョブ:XXXを実施します。(送信予定:n件)」时，**复制这条消息**（Ctrl+C），
   回到控制台按 Enter。**按 Enter 的时刻 = 开始时刻**（GoAnywhere 的行按「开始 −2 ～ +13 分」来找）。
   面板上会写这一行的日期：`YJDSWM40 (2026-10-09) 開始 10:45:20` 这种是**以前某天**留下的 pending 行，不是今天的，`s` 跳过（或 `-Only` 只跑今天的）。
   消息里的「送信予定:815件、46件」会被读成**两个文件**（件数按文件逐个列），用来核对 Receive 行数、下载数和 GIFT 文件夹里的文件数；
   和 mapping 的 `FileCount` 不同时报 `expect_disagree`。
   **打的时刻偏了**（比如打成 10:45 而 job 实际 10:15 跑的）：以前会一直刷到超时再问 r/s/q。现在等到超时时如果页面上有**附近**（窗口外再 ±5 分）的行，
   会弹一个 GATE：「10:45:20 没有行；页面上最近的行 10:15:52（−29 分），Enter = 用这个时刻继续 / s = 跳过这行、打对时刻再跑」。
   Enter 之后采用的时刻写回 `GFIX_TIME`。连附近都没有（job 真的还没出来）照旧 r 再等。
   - 这一行已经有 `GFIX_TIME`：Enter = 用表里的时刻（不需要消息）；`n` = 用现在的时刻（要先复制消息）。
   - 复制晚了：不要按 Enter，打领导预定表上的时刻，比如 `10:30`。
   - 消息是别的 job 的：面板会写出两个名字，等你复制对的那条；`s` = 先跳过这个 job（留到下次），`k` = 不用消息、时刻 = 现在。
2. **PREPARE 面板**：「Teams の『ジョブ:XXXが正常終了しました。』が来たら Enter」。
   reaction 你自己在 Teams 点。
3. 自动：在 GoAnywhere 上执行 Ctrl+F「詳細」→ Esc → Tab → Enter（**更新**，不按 F5），
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
   **OPEN 侧异常终了**（10/13 的 JJDSWM41：GoAnywhere 的 Send 行红叉、解決済み = No）：读完这个 job 的行、截完 GoAnywhere 的图就**停在这里**——
   不下载、不比较、不开 DF；SHARE 面板的文字变成「OPEN側で異常終了でした。」+ 哪一行 No，只带 GoAnywhere 一张图，贴到 Teams 让他们知道；
   作业清单 `track = ng`。之后 logs / evidence 碰到这行会自己跳过（「track judged this job ng: nothing to collect」），HOST 侧重跑后把 `track` 清空再跑一次就行。
9. **SHARE 面板**：文字「OPEN側で確認できました。前後一致です。 (受信データ:815件、46件)」+ GoAnywhere 图 + DF 图已经放进剪贴板。
   括号里是 Teams 消息里的件数，逐文件和下载文件的行数核过（允许多 1 行表头、0 件 = 空文件或 1 个空行；**行数比件数少**也算正常——HOST 侧多算 1 条、去掉空行时会差 2 行以上，
   10/15 起不报 warning，`-Var countsShortOk=false` 可以改回严格模式）；对不上的文件写成 `行数件(消息件数件)`，
   例 `0件(2件)、1件(4件)`；行数比「件数＋表头」还**多**时，才会在末尾 warnings 里报 `count_disagree`——这不是判定，判定还是逐行比较。没有消息（k 跳过）就不带括号。
   → 点 Teams 的输入框 → **Ctrl+V** → 看一眼 → 发送 → 回到控制台按 Enter。
   多个文件时，DF 图会按顺序全部带上。
10. 记录：`capture\gfix\<W名>\track.json`（后面 logs / evidence 用），作业清单 `track = ok`。

**第一次跑请特别看这几处**（都在 `capture\gfix\<W名>\` 里）：

| 看什么 | 文件 | 不对的话 |
|---|---|---|
| GoAnywhere 截图范围是不是 p4 那样 | `goanywhere.png` | `ebi.local.json` 里覆盖 `pages.goAnywhere.crop`（left/right/top/firstRowTop/rowHeight） |
| 「更新」按钮配方有没有点到 | 控制台：超时（timeout）= 没刷新出来 | `pages.goAnywhere.refreshRecipe`（从「詳細」数 Tab 的次数） |
| DF 是不是一屏 16 行 | `df\*__first.png` | `layout.df.width/height` |
| Teams 粘贴时图片在不在 | Teams | 把 `layout.share.mode` 改成 `"sequence"`（文字、图片分几次贴） |
| 截图里有没有残留的蓝色全选 | `goanywhere.png`、证据里的 Jenkins 图 | `deselectAt` 换一个空白位置（窗口内像素） |
| DF 有没有被拉起来并截到 | 控制台 `no_window` | DF 是启动器型程序的话会按标题 `DF - ` 找窗口（`layout.df.windowTitle`） |

`ebi.local.json` 覆盖写法示例（深度合并，只写要改的键）：

```json
{ "layout": { "share": { "mode": "sequence" } },
  "pages": { "goAnywhere": { "crop": { "top": 140 } } } }
```

### 2b. 一个 job 两个以上文件（FileCount = 2；10/13 第一次：JJDSJM10 ×2、JJDSJM51 ×2）

mapping.xlsx 里同一个 job 有几行，作业清单的 `FileCount` 就是几（plan 合并成一行）。整条链都是按「一个 job、n 个文件」写的，
但 10/13 之前没有真跑过 n ≥ 2，这次要盯的地方：

- **开始**：Teams 消息里的「送信予定:n件」也会读出来。它和 `FileCount` 不一致时出警告 `expect_disagree`（以 `FileCount` 为准）；
  `FileCount` 是 `-1`（mapping 里没这一行）时就用消息里的 n 来核对。
- **GoAnywhere**：要等到这个 job 的 **2 行 Receive** 都出现（少一行是警告 `count_mismatch`，不是错误；第二行还没出来就 `r` 再等）。
  截图从这个 job 的第一行截到最后一行（中间夹着别的 job 的行也一起带上），不再只数「匹配到的行数」。
- **Jenkins / 下载**：Receive 时间段里的 `F…csv` 应该也是 2 个，全部下载到 `DATA\GFIX\<W名>\`。
- **GIFT 侧**：`DATA\GIFT\<J名>\` 里要放 **2 个** 文件。少了会先警告「expected 2, found 1」，接着配对那一步会停下来问（1 左 / 2 右配不上）：
  把缺的文件放进去，`r`。
- **配对**：行数相同 + 到达顺序。GIFT 侧的「到达顺序」其实是**文件名顺序**（ALOC9002 → ALOC9004），GFIX 侧是 `F+时刻` 的顺序。
  两个文件行数一样、而 HOST 发送的顺序和文件名顺序不同时，两对都会比出「不一致」——这不是数据问题，是配错了对。
  现在 `file.compare` 会在这种情况下**按内容重配**：每个 GIFT 文件恰好能找到一个内容完全相同的 GFIX 文件、而且不重复，才重配并重新比较；
  结果 ok，但会在运行末尾的 warnings 里列出「re-paired … by content」，`track.json` 里 `repaired` = 重配的对数、
  `pairsAsListed` = 原来的配法。找不到唯一匹配（比如两个 GIFT 文件内容一样、或者真有一个文件不一致）就**不动**，照旧 ng 停下来问你。
- **DF**：每一对各开一次 DF、各截一张（>16 行再一张），窗口标题要同时含两边的文件名才截。Teams 的粘贴里 DF 图按配对顺序全部带上。
- **logs**：2 个 Receive job 号各下载一份 job log → `log\GFIX受信ログ\<号码>.log` ×2；两份里各读出一个转送文件名（带微秒时刻，不会重名）。
- **evidence**：
  - `GFIX受信ログ` 下面两份 job log **连着贴、中间不空行**（样本里就是这样）；黄色标记每份各一组。
  - `GFIXReceive.log` 下面两个转送文件的 START…END 块按日志顺序连着贴。两个转送**同时**跑时日志是交错的
    （START A、START B、END A、END B——10/13 的 JJDSWM10 就是）：每块从自己的 START 切到自己的 END，重叠部分只写一次，
    末尾 warnings 里是 `block_interleaved`（信息），不再是 `block_incomplete`。
  - Jenkins 截图：两个文件各一个红框；两行紧挨着时合成一个框（和你手工画的一样）。
  - `GIFTデータvsGFIXデータ`：DF 图按配对顺序往下排，**每对之间空一行**，每对的最后一张图上画「最后一行行号」和「同一内容」两个框。
    这一条是推断（样本里没有多文件的这张表），第一本做完请对一眼。
  - B3 = 送信データ!A3 是 IF 名（JJPCRS12）。但 10/13 加瀬さん的レビュー（JJDSWM10，检查表 328 行）要求每一段都在 G 列写上
    送信データ 的 STEP 名（ニンヌ 改好的 JJDSWM51：B3 + G3 `STEP:MQFXRSX1`，B26 + G26 `STEP:MQFXRSX2`）。
    v2.27.0 起自动写**第一段**的 G3；第二对文件那一段的相关ID + STEP 名**还要手写**（TODO）。

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
  现在是：找到「ジョブログをダウンロードする」后直接 Enter 下载，然后 Tab×2 → Enter 用页面自己的按钮回到一览（不用浏览器的「返回」，那会出「フォームを再送信しますか?」）。没下载下来的话，你手动下载，Downloads 里有带 job 号的文件，按 `r` 就会直接拿走它。
- 从 job log 里读出转送文件名（`JJPCRS12…`），写进 `capture\gfix\<W名>\logs.json`。

---

## 4. 证据：`gfixRecv.evidence`

```powershell
ebi.ps1 run workflows\gfixRecv.evidence.json -WorkDir $W -Only LJDSWM40
```

**先拿一本复制件试**（把 `paths.bix` 临时指向一个本地复制文件夹）。插图这一步不可重复执行（not idempotent），
跑两次会插两遍。

工作簿被别人打开着时，`open` 会失败并写出**是谁**（读 Excel 放在旁边的 `~$J4検証資料…xlsx` 锁文件里的用户名）：
「J4検証資料(...)_JJDSWM10.xlsx opened read-only -- JP0xxxxx has it open. Ask them to close it, then r」。锁文件读不到名字时照旧只说 read-only。

GFIX受信結果（按样本 QJDSWM39 的版式）：

- B3：**从真正的 GFIX一覧 现截的 Excel 截图**（v2.27.0 起，旧 ExcelSnap.ps1 的做法：`送信 / ジョブ` 列筛选成这个 job 的 J 名，
  从 `from / システム種別` 截到 `ジョブ`）。同一个 job 有多行、而且 from～送信ジョブ **每一列都一样**时，一直截到
  `元ファイル(SNDFILE)`，并把 **GIFT受信結果 的第 1 张图也换成这张**（加瀬さん 10/13 レビュー 326/327 行；IF番号 不同就不算重复）。
  要先在 `ebi.local.json` 的 `paths` 里加 `"gfixList": "<GFIX一覧.xlsx 的完整路径>"`（只读打开，筛选后不保存）。
  工作表名和表头的字在 `profiles/gfix-recv/layout.json` 的 `evidence.list`。
- 下一行：`GFIX受信ログ` + job log 全文（多个 job log 之间不空行），黄色标记 `successfully uploaded to …` 和 `Command: …`。
  宽度 = 文字半角宽度 ÷ 3 列（用样本的 4 处黄色标记校准过，结果一致）。
- 空一行：`GFIXReceive.log` + 这次转送的 START…END 块（含 END 后面的日期行），黄色标记 `FileGIFTGetBLBean.execute() fileName` 和 `file stored :`。
- 空一行：Jenkins 截图（F5 → Ctrl+End）+ 红框框住收到的文件行。行的位置是按蓝色链接文字的行，对照一览文本**从最后一行往上数**找到的；
  用样本截图验证过（样本红框 481px，算出来 482px）。

GIFTデータvsGFIXデータ：

- B3 = 送信データ!A3（JJPCRS12），G3 = 送信データ!B3（STEP 名，前后空格去掉）。
- B4 起：DF 截图。超过 16 行：开头截图 → Z 列放波浪线（p5）→ 结尾截图；
  红框框住**最后一行的行号**（在截图里找行号那一列最后一行墨迹）和右下角「同一内容」。
  （和样本对比：行号框 213px vs 212px。）
- 多个文件：按配对顺序往下排，中间空一行（**这是推断**，样本里没有多文件的 GIFTvsGFIX）。

最后：每张 sheet 选中 A1 并滚到顶部（等同 Ctrl+Home）、取消复制模式（Esc）、回到第 1 张 sheet、保存；
检查 GFIX受信結果 B 列的字体是不是 ＭＳ ゴシック 10（只报告，不改）。
然后 **GATE**：你在 Excel 里看一遍，`Enter` = ok，作业清单 `evidence = ok`。

GATE 的面板上：
- `check:` 行 = 要看的点（`profiles/gfix-recv/layout.json` 的 `evidence.checklist`，看漏过什么就往里加一条）。
- 第一行如果是 **红字 `!! …`** = 自动检查和预期不一样：B3 那张图 OCR 出来的「ジョブ」不是这个 job 的 J 名
  （另一个 job 的图？），或者 GIFT受信結果 里有 Ctrl+G 编组的图（复制的可能不是第 1 张）。**不会停**，你看了再决定 o / n。
- 自动检查和你的回答并排存在 `capture\gfix\<W名>\gate.json`（以后据此调整检查，见 `docs/ebi-dance/KNOWLEDGE.md` 第 9 节）。

跑的时候 Excel 会被拉到前面，并滚到正在写的位置：**可以看，别点**。点了 Excel，下一次写入会被拒
（`0x800AC472`）；现在会自动重试几次，还不行才停下来问你。想中途停：点一下控制台窗口，按 `q`（或 Esc），
下一步开始前会问 `STOP?`，`y` = 停（已经做完的行保持 ok）。

---

## 4b. 交付：`gfixRecv.deliver`（v2.26.0 起；10/19 起「先全部写、一次确认；再全部草稿、一次确认」；10/20 起最后填 WBS 実績）

```powershell
ebi.ps1 run workflows\gfixRecv.deliver.json -WorkDir $W -Only QJODWSPN,RJODWSPN,SJODWSPN
```

evidence = ok 的行：

1. **每本**（each）：先**只看**レビューチェックシート（Check Sheet_J4）的 F 列（レビュー対象 = 证据工作簿的文件名）里有没有这本。
   **已经有了**（手工写过，比如 10/14 的 R/S）→ 当作已经委托过：不写行、不建草稿，`deliver = ok`。
   没有 → 往检查表末尾**写一行**（B 记入日 = 今天、C JAVA、E J4内部ﾚﾋﾞｭｰ、F 文件名、G owner、H reviewer），**先不保存**。Excel 跟着显示写到的那一行。
2. **全部写完后一次确认**（after / rowsOk）：Excel 被调到前面，看这几行。
   `Enter` = 全部正确 → 保存并关闭检查表 → 下一步。`s` / `n` = 先不交付：什么都不保存，这几本保持 pending。
3. **一次建好全部草稿**（after / mails）：用 `mailTemplate`（`sendj4mail.bat` 本身、`review-mail.json` 或 `.msg`）每本建一封 **Outlook 草稿**并打开。**工具不会发送。**
4. **一次确认全部送出**（after / sent）：所有草稿窗口被调到前面，逐封检查（收件人、主题里的 job 名、正文里的文件名）并按「送信」，全部送完回控制台按 `Enter` → 这几本一起 `deliver = ok`。
   `s` = 稍后再送：检查表的行已经保存，deliver 仍是 pending；从「下書き」里送完后再跑一次 deliver，它会看到检查表里已经有这几行，直接记 ok，不会再建草稿。

5. **WBS 的実績**（v2.29.0 起；v2.34.0 起「薄紫确认、确认后才真正写」）：全部送出（第 4 步 `Enter`）之后：
   - 问实绩日（Enter = 今天，也可以打 `10/21`）和工数（Enter = 0.1）。
   - 等 WBS 没人开着，打开（**不是**只读；只读就停下，说是谁开着——可能就是你自己的 Excel）。
   - 写哪几行：A = job、B = 受信、G = 修正後実施 / 結果まとめ、P = 你（`owner` / `owner←○○` / `○○→owner`）。写 AJ = AK = 实绩日、AQ = 工数。
   - **只写空格子**；**有公式的格子不写**（看起来空也不写，面板列出「has a formula」）；日期写进「常规」格式的格子会设成 `yyyy/m/d`（和手打日期一样）。
   - 写到的格子涂成**薄紫色**（`#E6D6FF`；表里已经有相近的颜色就换下一个候选），Excel 调到前面，面板列出每一行。被筛选藏起来的行标 `(hidden by a filter)`。
   - `Enter` = 对：**不保存**关掉这份（薄紫色不会进文件），重新打开，写同样的值（**不涂色**），保存。
   - `s` / `q` / 其他 = 不对：**不保存**关掉，WBS 保持原样（自己手填）。

检查表被别人开着时，`open` 会说是谁，等对方关掉再 `r`。需要 Classic Outlook（新版 Outlook 不能自动化）。

---

## 5. 出问题时

- FAILED 面板上的 `hint:` 行 = 这种错误常见的原因和下一步。临时性的错误（窗口还没出来、Excel 在忙）`Enter` = 重试，
  其余的 `Enter` = 跳过这一行。
- 出现 UNFINISHED RUN 面板时，`Enter` = 新开一次 run。**红字的是以前某一天的 run**：一般选 `d`（忘掉它）或 `n`，
  不要 `r`（旧 run 按 job 名重放，同一个 job 换了日期再来时会被当成已经做完）。

- 中途断了：同样的命令加 `-Resume`（续跑上一次的 run，已经做完的步骤不会重做）。
- 想重做某个 job：在 `gfixrecv.csv` 里把对应的 `track` / `logs` / `evidence` 清空。
  同一个 job 以后换了日期再次出现在计划里时，plan 会自动把这三列清空（`resetOnChange`）。
- 所有中间产物：`capture\gfix\<W名>\`（截图、页面文本、track.json、logs.json、receive.txt）。
- 每次运行的记录：`run\<runId>\`（trace、ledger）。出了问题把这个文件夹的内容发给我就能看。

---

## 5b. 可以按你的习惯改的默认值

| 默认 | 现在 | 怎么改 |
|------|------|--------|
| FAILED 面板的 Enter | 临时性错误 = 重试，其余 = 跳过这一行 | step 的 `transient`（告诉我改哪个） |
| UNFINISHED RUN 的 Enter | 新开一次 run（以前某天的 run 画红） | `ebi.ps1` |
| 第 1 张图的 job 名对不上 | GATE 画红字，**不停** | `gfixRecv.evidence.json` 的 `look` |
| Excel 忙的时候写入失败 | 自动重试 4 次再问 | `gfixRecv.evidence.json` 的 `onError.byFailure` |
| 文件行数比 Teams 件数少 | 正常，只在回复里标 `行数(件数)` | `-Var countsShortOk=false`（只这一次）或 `gfixRecv.track.json` 的 `vars` |
| GATE 的 `check:` 行 | 5 条 | `profiles/gfix-recv/layout.json` 的 `evidence.checklist` |
| 邮件 | 只做草稿，你来发送 | 不提供自动发送 |
| 窗口跳到前面 | 每步把用到的窗口（Edge / DF / Excel）拉到前面，步与步之间回到控制台 | **静默模式**（v2.33.0）：`ebi.local.json` 写 `{ "operator": { "silent": true } }`，或命令行加 `-Silent`。Excel 在后面写、不再停顿给你看；Edge / DF 这类必须看见或打字的步骤仍会占前台一小会儿（这期间别打字），之后**还给你原来在用的窗口**；只有要你回答的面板才把控制台拉到前面；跑完只闪任务栏按钮 |

觉得哪个不顺手就说一声；改动会记进 `docs/ebi-dance/KNOWLEDGE.md`。

---

## 6. 我做的假设（猜错就要改的）

| 假设 | 依据 | 改哪里 |
|---|---|---|
| GIFT 侧文件夹 = J 名（RJDSJM40），GFIX 侧 = W 名（RJDSWM40） | 你的描述 | workflow 里的 `DATA/GIFT/{{item.JOB}}` |
| 一个 job 的文件数 = mapping.xlsx 里这个 job 的行数（Teams 消息的「送信予定:n件」做交叉核对） | mapping 里 JJDSJM10、JJDSJM51 各 2 行 | 不一致只是警告（warning）；两边都知道且不同时报 `expect_disagree` |
| 两个文件行数相同、配错对时，按内容重配不算猜 | 「前後一致」= 内容逐行相同；DF 标题里有两边文件名，证据自描述 | 不想要就把 track 里 `cmp` 的 `repairByContent` 改成 false |
| 多文件的 GIFTデータvsGFIXデータ：每对之间空一行，B3 只写一个 IF 名 | 10/13 JJDSWM10 实物：4 张 DF 图 + 2 条波浪线、B3 = JJPCRS12，通过 | `excel.stack_plan` 的 `gapRows` |
| 打错的开始时刻：页面上最近的行就是这个 job | 10/13 打 10:45 实为 10:15；gate 让你确认，不自动采用 | `pages.goAnywhere.nearBefore/nearAfter`（±5 分之外的提议会标「farther」） |
| Teams 消息的件数 = 每文件一个数，文件行数 = 件数或件数+1（表头），0 件 = 空文件或一个空行 | 10/13「815件、46件」对两份文件 | `verify.record_counts` 的 `headerLines` |
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
