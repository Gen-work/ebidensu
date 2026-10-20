# ebi-dance 现场知识库(KNOWLEDGE)

办公 PC 上实跑时撞到的边界情况(edge case),一条一条记下来。**写新 workflow
之前先读这一页**——人也好、agent 也好,这些坑不用再踩一遍。

每一条都有同样的五栏:

- **症状**:现场看到了什么(原文日志、面板上的字)
- **原因**:查下来是为什么(没查清就写「未确认」,不要编)
- **现在的对策**:落在哪个 step / profile / workflow 里(给文件名,能点过去)
- **写新 workflow 时**:照这条该怎么写
- **来源**:哪天、哪个 job、哪个分支或 PR

---

## 0. 写新 workflow 的检查清单(从下面各条归纳)

1. 找窗口的 step:每个页面**单独一个浏览器窗口**,而且这个页面是该窗口的当前标签页(K-01)。
2. 写 Excel 的 step:`onError.byFailure` 给 `write_failed` / `save_failed` 加 `retry`(K-02)。
   不幂等的 step(`excel.insert_pictures`、`excel.copy_picture`)**不要**自动重试:重试会多贴一张图。
3. 从别人的成果物里**复制**来的东西,不只核对位置,还要核对**内容**:能读文字就读文字,只能看图就 OCR(K-03、K-04)。
4. 人工关卡(`human.gate`):要看的点放进 profile 的 `checklist`;自动检查有异议时用 `alert` 画成红字;**不要**因为自动检查有异议就停下(K-03)。
5. 写共享文件(检查表等):**先只看**(`checkOnly`),等人确认之后再写;按 key 判断是否已存在,保证幂等,手工补过的行不会写成两份(K-14)。
6. 对外发送的东西(邮件、Teams):工具只做到**草稿 / 剪贴板**,发送永远由人来按(K-15)。
7. 关卡的自动结论和人的回答**并排存下来**(`capture/<job>/gate.json`),这是第 9 节「从人的回答里长出规则」的原料。
8. **默认行为先问操作员的偏好**(Enter 的含义、对不上时停不停、重试几次、旧 run 怎么处理、数字容差),并告诉他在哪里改。问法见 `INTERVIEW.md` 的 6.5。不问,操作员就以为「程序只能这样」(K-05、K-06、K-10 都是跑过之后才说出来的偏好)。
9. 从对方系统下载东西:以**唯一的编号**(job 番号)为 identity,哪个 job 的东西**看内容**判断,不看列表上的名字;要换筛选的页面,先下载完当页再换(K-18)。

---

## 1. 窗口与浏览器

### K-01 窗口标题 = 当前激活标签页的标题
- **症状**:`setup/jk screen.find_window not_found: no window title contains "data/report [Jenkins]"`,按 `r` 又找到了(10/15)。
- **原因**:Edge 窗口的标题就是当前激活标签页的标题。那一刻激活的是别的标签页,或者页面还在加载,标题还是 URL。
- **现在的对策**:`kernel/Gate.ps1` 的 `Get-EbiFailureHints` 在 FAILED 面板上给出提示;`not_found` 是临时性(transient)失败,所以 Enter = `r`。标题相同的窗口有好几个时,`screen.find_window` 用最前面那个,并发 `several_windows` 警告。
- **写新 workflow 时**:PREPARE 面板里写明「另开一个普通窗口」;找窗口用的标题片段要选页面加载完以后才会出现的那一段。
- **来源**:10/15 R/S track。

### K-02 用户在 Excel 里点击 → `0x800AC472`
- **症状**:`excel.write_lines write_failed: HRESULT からの例外:0x800AC472`;按 `r` 又成功了(10/15)。
- **原因**:`0x800AC472`(VBA_E_IGNORE)是 Excel 在说「我正忙,这次调用不接」。工具正通过 COM 写入的时候,有人切到 Excel、点了单元格(或者正处于编辑状态、开着对话框),下一次调用就会被拒。
- **现在的对策**:evidence / deliver 两段对 `write_failed` 自动退避重试(1.5 s 起,翻倍,4 次);写入前 `Show-EbiExcelAt` 把 Excel 拉到前台,并滚动到正在写的那一行(`kernel/Excel.ps1`),**看可以,别点**;两步之间按 `q` / Esc 会弹出 STOP? 面板。
- **写新 workflow 时**:同上面清单第 2 条。不要假设操作员会盯着控制台看。
- **来源**:10/15 RJODWDB6 evidence。

### K-03 证迹里复制来的「第 1 张图」是别的 job 的
- **症状**:GFIX受信結果 B3 的 Excel 快照是另一个 job 那一行的(操作员 10/15 报告)。
- **原因**:GIFT 侧贴图的人贴错了行;工具只按「第 1 张」取图,不看内容。
- **现在的对策**:`excel.copy_picture savePng` 把图导出成 PNG → `screen.ocr` → `verify.find_token`,确认最后一列「ジョブ」= 这一行的 J 名(`item.JOB`)。读到别的名字时 ng,差一个字符时 unknown,两种都在 GATE 上**画红**,但不停下来。
- **写新 workflow 时**:复制别人的图之前先问一句「这张图上该有什么字」,有就加一道 OCR 检查。OCR 会把 O/0、I/1、S/5、B/8 认混,`verify.find_token` 已经做了折叠处理。
- **来源**:10/15 反馈第 9 条。

### K-04 Ctrl+G 编了组的图不算「图片」
- **症状**:还没有实际发生;是推测出来的风险。
- **原因**:`Get-EbiPictureShapes` 只数 msoPicture(13 / 11),编组(msoGroup 6)不算在里面,于是「第 1 张」变成了组后面那张。
- **现在的对策**:`excel.copy_picture` 发现源工作表里取图位置及其上方有编组时,出 `group_shape` 警告,GATE 上也有红字。
- **写新 workflow 时**:按序号取图的地方,一律带上这个检查。
- **来源**:10/15 讨论。

---

## 2. 运行、恢复、交互

### K-05 恢复前几天的 run,会按 job 名重放
- **症状**:10/15 运行 evidence 时,面板问要不要恢复 10/14 的 run;选了 `r` 之后,新 job 照常跑完了。
- **原因**:ledger 的 key 是「job 名 | step id」,**不包含日期**。如果同一个 job 换了日期又出现,而旧 run 里它已经 ok,恢复旧 run 时它的 step 会被当成「已完成」直接重放,证迹并不会生成。
- **现在的对策**:UNFINISHED RUN 面板默认 `n`(新开一次 run),别的日期开始的 run 画红,并标注 `<- another day`。
- **待办**:ledger key 加上行的日期,见 `docs/TODO.md`。操作员 10/15 表示现在这样(提示 + 默认 `n`)就可以,但要讲清楚:这是一个可以选的默认值。
- **来源**:10/15 反馈第 5 条。

### K-06 FAILED 面板上 Enter 的含义
- **症状**:`not an option: ""`;操作员以为 Enter 就是 `r`。
- **现在的对策**:临时性失败(窗口还没出来、Excel 忙)Enter = `r`;其余失败(没有 track.json 之类)Enter = `s`。面板上的 `hint:` 行写着常见原因和下一步该怎么做(`Get-EbiFailureHints`)。
- **写新 workflow 时**:step 的 manifest 里认真标好 `transient`,它同时决定了重试策略和 Enter 的含义。
- **来源**:10/14、10/15。

### K-07 ebi.local.json 写坏了,或者存成了 ANSI
- **症状**:`no operator.owner set`,当天所有人的 job 都进了清单(10/15 plan 多出 7 行)。
- **原因**:`{ "operator": { "owner": "厳" } }` 被当成一个**值**塞进了顶层对象里,应该是 `"operator": { "owner": "厳" },` 这样作为顶层的一个键。另外,记事本按「ANSI」保存时文件是 CP932。
- **现在的对策**:CP932 照常能读,但会警告;owner 为空时,plan 会说明「不过滤」。
- **写新 workflow 时**:按人过滤的值放在 `operator.*`,workflow 开头要把「没有设置」的情况**说出来**。
- **来源**:10/15 反馈第 2 条。

### K-08 当天计划临时改了,只跑其中几个
- **症状**:预定表从 8 个 job 改成 4 个,你只负责其中 2 个(10/15)。
- **现在的对策**:`-Only k1,k2`(逗号分隔,W名)。plan 照常跑,多进来的行就留着;以后换了日期再出现时,`resetOnChange` 会自动清空。
- **来源**:10/15。

---

## 3. 数据与对方系统

### K-09 0 字节的传输 → `box_not_placed`
- **症状**:`box_not_placed no ink found ... expected for a 0-byte transfer`。
- **原因**:清单没有行,自然没有「最后一行」可以画框。
- **现在的对策**:只出警告,不判失败。
- **来源**:10/14 CJODWDB6 / FJODWDB6,10/15 RJODWDB6。

### K-10 Teams 的件数比文件行数多(这是正常的)
- **症状**:10/15 RJODWDB6「送信予定:1件」→ 文件 0 行;SJODWDB6「127件」→ 126 行。两个都报了 `count_disagree`,但前后比较完全一致。
- **原因**:HOST 侧数件数时,会多算一条文件里没有的记录,所以大多差 1;如果连空行也被去掉了,就会差 2 行或更多(操作员 10/15 确认)。另一方面,文件里带表头时又会多出 1 行(10/13)。
- **现在的对策**:`verify.record_counts shortOk`(gfixRecv.track 的 `vars.countsShortOk = true`):行数**少于**件数算正常,不再报 warning,但回复文本里照样写成 `126件(127件)`,把差异标出来。行数比「件数 + 表头」**还多**的,仍然报 `count_disagree`。`-Var countsShortOk=false` 可以恢复成严格模式。
- **写新 workflow 时**:件数和行数「差几行算正常」是这份工作自己的规则,要先问操作员(清单第 8 条)。
- **来源**:10/15 track 日志 + 操作员说明。

### K-11 对方系统(OPEN 侧)异常终了
- **症状**:10/13 JJDSWM41,当天是人工通知的。
- **现在的对策**:GoAnywhere 的「解決済み」列(只有异常终了的行是 `No`)→ `verify.assert_each` → 后续步骤用 `when` 跳过 → `shareFail` 往 Teams 发「OPEN側で異常終了でした。」;后面几段 workflow 开头的 `flow.require` 把这一行挡掉。
- **来源**:PR #167。

### K-12 保存时弹出「文書分類/機密区分の登録」
- **现在的对策**:`dismissDialogs`(`excel.tidy`、`excel.append_row`)自动用 Esc 回答。
- **来源**:10/12。

### K-13 共享工作簿被别人打开着
- **现在的对策**:`excel.open` 的 `read_only` 会读 `~$` 锁文件,说出是谁在用;问一声,对方关了再 `r`。`excel.append_row` 遇到只读的工作簿,不会去写。
- **来源**:10/13。

### K-16 WBS 的「计划开始」不等于「今天要做」
- **症状**:10/16 提前做了 10/20 的 CJODWSPQ、K/R/SJODWSPP,WBS 的 AF(最新計画 開始)还是 10/20 → 只看 AF 的 plan 在 10/20 会把做完的又拉进清单;DBC 几本 AF 是 10/15～10/19 却一直没做(等 HOST 修复),只看「AF = 今天」永远看不到它们;往日没做完的行又一直算 pending。
- **原因**:WBS 里计划(AF/AG)和实绩(AJ/AK)是两组列,实绩靠人手填、会晚填、会填错(10/16 把还在等的 K/LJODJDBC 填了 AJ)。
- **现在的对策**(v2.29.0):plan 读 AF/AG/AJ,分成 today / late / ahead / others / mapping / hold 候选集合(`table.day_sets`),AJ 有日期或作业清单全 ok 的只列在「done already」一行;操作员用 `human.pick` 选(回车 = 今天的)。作业清单加 `WORKDAY`(这一行是哪天的活)和 `hold`(保留),各段只拿今天的、不拿保留的;deliver 送完后把 AJ/AK 自动写回 WBS(只写空格子)。
- **10/20 补**:AJ 不能单独当「做完」:操作员早上开工前就给 M/SJODJSPQ 填了 AJ 10/20。现在「做完」以作业清单为准(最后一列 deliver ok);AJ 只在作业清单从没有这本、且日期早于今天时才当作做完,否则照样列出并注明 `WBS AJ MM/dd`。`plan -Only A,B` = 直接选这几本(以前 plan 不认 -Only,10/21 的和 WBS 担当写别人的那几本进不了清单)。
- **写新 workflow 时**:「今天做哪些」要分清计划和实绩两个来源,把「做完了」的信号写回原表,而不是只在工具里记;人手填的实绩只当参考。
- **来源**:10/19～10/20 WBS 实数据(DETOUR-GFIX-001.10)+ 操作员说明(DBC 全部在等;AJ 靠人手填)。

### K-17 GIFT 文件不在 DATA\GIFT 里
- **症状**:10/20 M/SJODJSPQ、QJODJSPP 跑 track 时 `DATA\GIFT\<J名>` 没有文件,track 建好空文件夹停下来等人放。
- **原因**:GIFT 侧文件要人手先从 Jenkins data/report 下载;它们和 GFIX 侧同一个页面、同样是 `F<yyyyMMdd><nnnn>` 的名字,名字里看不出是哪个 job,只能按 GIFT 运行的时刻找(操作员 10/16)。
- **现在的对策**(v2.29.0):plan 从 mapping.xlsx 带上 GIFT実行日(M)/ GIFT TIME(N);新的一段 `gfixRecv.gift` 在 track 之前检查文件夹,没有就读 Jenkins 一览,时刻 −2～+13 分钟里正好是该有的个数就直接下载,否则让人选(时刻窗口、那天按小时分组)。
- **还没做**:处理時間(BIX).xlsx 的开始/结束时刻当来源(更准),Teams「実施します」消息当来源。
- **来源**:10/20 操作员(余晖的 10/20 受信日指示)+ docs/TODO.md 10/16 的卡片。

### K-18 GoAnywhere 接收 log:IF 号和 project 名都不唯一,只有 job 番号唯一
- **症状**(10/20,一次性的 12 件 log 收集):同一个 IF 的接收 job 在列表里有两三行(两个 correl 共用一个 IF、发送方重发);project 名 VER1 / DEV1 混着出现;不同发送方(LPS / ALC / CRY)用同一个 IF 号,又落在同一个接收文件夹;correl 末位不固定(M / S)。列表一次只能按一个 project 名筛,Ctrl+F 只找得到当前页上的 job 番号。完成ジョブ页面现在没有 User 列、末尾多一列「解決済み」。
- **原因**:job 番号唯一;IF 号 / project 名不唯一。哪个 correl 的 log,只有 log 里的 Command 行(`/gfix/recv/<correl>`)写得清楚。
- **现在的对策**:`GoAnywhereLogFetch.ps1`(独立脚本,不是 phase;配置 JSON 留在本地)。候选 = IF 号(可多个)+ 接收文件夹;归属 = log 内容(correl 末位任意);每筛一次就在当页下载;重跑默认取最新的,旧的写进 report;文件名保持下载时的原名;共享盘上已有的文件不覆盖(EXISTS)。列表按「形状」找列(第一个 `/` 开头的是文件夹,日期形的是时刻),不按位置。实跑结果:12 件里 OK 11、EXISTS 1、NG 0。
- **还没定**:重跑时取哪一次(实跑中有一件两次 log 都写着同一个 correl,默认复制了较新的那次,要人看实施日再判断)。`modules/verify/GfixJobList.ps1` 按固定位置取列,在没有 User 列的新页面上 StartTime 之后的列会错位;GfixLogDownload 只用 ProjectName / Folder,是否受影响**未确认**。
- **写新 workflow 时**:以 job 番号为 identity,不靠 project 名定位;归属一律看内容;筛选和下载放在同一轮(换筛选之前下载完);重跑的取舍交给人,报告里把几次并排列出;共享盘只复制、不移动、不覆盖。做成 step 时大致是 `browser.read_text` → 解析列表 → 每个 job 番号 `browser.find` + 下载 → `verify.match_record`(内容)→ 复制。
- **来源**:10/20 一次性作业,分支 `claude/youthful-franklin-m983ft`(v2.29.1),docs/GoAnywhereLogFetch.md。

---

## 4. 交付(gfixRecv.deliver)

### K-14 检查表里已经有手工写的行
- **症状**:10/14 已经手工把 R/S 写进了 Check Sheet_J4(第 355、356 行)。
- **现在的对策**:`excel.append_row checkOnly` 先只看 F 列(レビュー対象)。已经有了 → 视为依赖已发:不建草稿、不提问,直接 deliver = ok。没有 → 建草稿 → 人按发送 → Enter → 这时才写入检查表。中途按 `s` 什么都不写。
- **来源**:10/15 反馈第 12-3 条。

### K-15 邮件模板里 job 名前面多了一个空格
- **症状**:正文里是 `J4検証資料(...)_ MJDSWM39.xlsx`,`_` 后面有一个空格。
- **原因**:模板 HTML 里 token 单独占一个 `<span>`,前面带了空格。
- **现在的对策**:`human.draft_mail trimBefore`(默认开)会删掉 `_` 或 `>` 和 token 之间的空格;操作员手里的 bat 也修好了。
- **来源**:10/15 反馈第 11 条。

---

## 9. 从人的回答里长出规则(反馈闭环)

每次 GATE,自动检查的结论(`auto`)和人的回答(`person`:ok / ng / 备注)都会并排存进
`capture/<job>/gate.json`。攒上几天之后:

1. **自动是 ng、人是 ok** 的情况反复出现 → 自动检查太严了(比如 OCR 的折叠规则还不够)→ 改那个 step,把这次的样本加进测试。
2. **人用 `m <备注>` 反复写同一件事** → 把它加进 profile 的 `checklist`(只是多问一句,不用改代码)。
3. **checklist 的某一条每次都被人确认过** → 想办法把它变成自动检查,在 GATE 上画红。
4. 这些变化都在这一页里补一条 K-xx,写明「来源」。

**怎么贡献**:先发现症状,再追原因;对策写进 step 或 profile 之后,在这一页加上 K-xx(五栏都要有)。
没查清的写「未确认」,和代码改动放在同一个 PR 里。
