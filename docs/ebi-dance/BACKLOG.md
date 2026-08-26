# BACKLOG — ebi-dance 任务卡

> **这份文档的用途**:把重构拆成一次坐下能做完的卡片。
> 每张卡自带足够上下文,**一个全新的会话拿到卡号就能开工**,不需要读完全部规格。
>
> 推进方式是碎片时间,不是整块开发日。**做完多少算多少,随时可停。**

## 怎么用

1. 挑一张**依赖已满足**的卡
2. 读卡里的「读什么」列出的文档段落(通常 1-2 节,不是整份)
3. 做,直到「完成判据」全绿
4. 提交,一张卡一个 commit,commit message 首行带卡号:`feat(P1-13): browser.read_text`
5. 回来把这张卡的状态改成 `[x]`

**不要一次做多张。** 卡的价值就在于边界清晰。

## 全局铁律(每张写码的卡都适用)

违反任何一条,`Tests/Run-Tests.ps1` 会失败:

| # | 铁律 | 为什么 |
|---|------|--------|
| R1 | step 文件**不许有 `param()` 块** | step 是被 dot-source 的;有 `param()` 会覆盖调用方的变量 |
| R2 | **源码纯 ASCII**,日文/中文用 `[char]` 或从 profile 读 | 无 BOM 的 `.ps1` 在 JP locale 主机上会乱码,曾静默破坏 owner 匹配 |
| R3 | 编码 UTF-8 **无 BOM** | 同上 |
| R4 | **禁止 `@($hashtable[$key])` 这个形状** | PS 5.1 binder 在 `List[object]` 上抛「参数类型不匹配」,本仓库栽过两次 |
| R5 | 不用 PowerShell `class` | PS 5.1 的 class 跨 dot-source 有作用域坑 |
| R6 | 失败用**返回值**表达,不抛异常 | 见 `spec/STEP-CONTRACT.md` §3.1 |
| R7 | `effects` 为 `write`/`destructive`/`ui` 的 step **必须**处理 `$Ctx.DryRun` | 见 `spec/STEP-CONTRACT.md` §3.3 |

## 状态

- 阶段 P0–P3 共 **68 张** = 「能接下一份工作」的最小集
- 阶段 P4–P5 共 **36 张** = 补齐 Excel/文件组 + Agent 循环
- `[整块]` 标记 = 需要连续思考,不适合碎片时间,也**不建议交给较小的模型**
- **计数口径**:「P0–P3 共 68 张」**不算 `P0-00`**(它已经是历史状态说明,
  不是一张待执行的卡);「P4–P5 共 36 张」不涉及 `P0-00`,和直接
  `grep -c '^### \['` 数出来的一致。只在数 P0–P3 时,直接 grep 会多数出
  1 张(69)——那 1 张就是 `P0-00`,不是漏卡也不是多算。

> **2026-08-24 评审修订**:开工前审了一遍契约,发现 6 个「现在改是文本、
> 写完 7000 行再改是重构」的洞,追加为 P0-R1…R6(规格修订卡,全部先于写码);
> P2 追加 2 张(P2-07/08);受影响的实现卡已就地改写。详见各卡的「问题」段。
>
> **2026-08-25 追加执行**:P0-R1…R6 全部执行完毕(纯文档改动,四份 spec +
> `Plan.md` + `docs/README.md` + `INTERVIEW.md` 同步到位),另追加三张卡
> P0-R7(冻结标签改名,避开远端已有的同名分支)、P0-R8(`Plan.md`/
> `docs/README.md` 与 BACKLOG 同步)、P0-R9(修正 P0-00 的虚假完成状态)。
> 九张 R 卡全部 `[x]`。
>
> **2026-08-26 第三轮评审追加**:P0-R2 只定义了 Session 资源的注册侧,
> 释放侧(谁在什么时候关掉 Excel/浏览器窗口、Ctrl+C 算不算数)从没定义,
> 追加第十张 R 卡 P0-R10(`[整块]`,Session 资源的生命周期·释放侧),
> 三条决定(显式释放 + 穷举异常退出路径 + 未注册资源必须自释放)已在
> 同一批提交里落进 `STEP-CONTRACT.md`/`WORKFLOW-SCHEMA.md`,`[x]`。**十张
> R 卡全部 `[x]`**。同批还修了一处悬空章节引用(`WORKFLOW-SCHEMA.md` §1
> 指向 `PROFILE-SCHEMA.md` 一个已被删掉的子标题)、`needs:session:<kind>`
> 与 `sessionKind` 的重复声明(改成只由 `sessionKind` 一处声明,`needs`
> 不再出现 `session:<kind>`)、`flow.call` 命名空间规则的两处遗漏(子
> 工作流内部引用改写、id 含 `.` 后 `.out.` 边界解析)、P1-12「抄」列表
> 漏掉的 `Send-ShiftTab`。

---

# P0 — 骨架(18 张,3 张整块)

目标:**一条 5 行的 workflow JSON 能真的存下一张 PNG。**

### [x] P0-00 契约与词汇定稿
已完成(2026-08-24):`spec/` 四份 + `INTERVIEW.md`。**状态修正记录
(P0-R9)**:2026-08-24 首次标记 `[x]` 时评审随即发现 P0-R1…R6 六个契约洞——
这份"定稿"当时其实并不算数,`[x]` 是假的;2026-08-25 P0-R1…R9 全部执行
并复核完毕后,才实至名归地恢复 `[x]`。

## P0-R — 规格修订(评审发现的契约洞,全部是改文本,先于任何写码卡)

### [x] P0-R1 [规格修订] page 身份与 role 解耦
- **估** 60min | **依赖** — | **改** `spec/PROFILE-SCHEMA.md` §2,3,4,5;`spec/VOCABULARY.md` §2,3;`spec/WORKFLOW-SCHEMA.md` §8
- **问题**:pages.json / grammar.json / rules.json / 工作流 id / capture 目录全都拿
  **role 当唯一键**,但一个项目同一侧可以有多个同型页面 —— 当前工作 before 侧
  就同时有 MQ転送状態(`GIFT_MQ`)和 Jenkins 文件列表(`GIFT_Jenkins`)两个
  list 页,after 侧还有 GoAnywhere 作业一览。按现规格,`GiftMqSnap` 和
  `GiftJenkins` 都叫 `before.list.capture`,截图都落 `capture/before_list/<key>.png`,
  **id 冲突 + 文件互相覆盖**。PROFILE-SCHEMA 的示例其实已经露馅:grammar.json
  里混着 `list`(role)和 `fileList`(页面名)两种键;vocabulary.json 示例还在用
  已被删掉的 `artifact` role 和不存在的 `query` role。这是 P2-01 一开工就会撞死的墙。
- **做**:引入 **page(命名页面实例)** 作为 profile 的第一等公民:pages.json /
  grammar.json / rules.json 一律按 page 名为键,每个 page 声明一个 `role` 属性
  (role 只决定「用哪套定位/解析机制」);工作流 id 改为 `<side>.<page>.<verb>`;
  capture 目录改为 `capture/<side>_<page>/`;清理两份 spec 示例里的 role/page 混用。
- **完成**:三份 spec 相互一致;把当前工作的全部页面(HM/MQ/Jenkins/GoAnywhere/帳票)
  逐个写出 page 名 + role,无一冲突

### [x] P0-R2 [整块][规格修订] 会话资源通道($Ctx.Session)
- **估** 90min | **依赖** — | **改** `spec/STEP-CONTRACT.md` §3,§6;`spec/WORKFLOW-SCHEMA.md` §4
- **问题**:契约规定 step 之间**只**通过 `{{steps.X.out.Y}}` 传值,且引用**只限同段**。
  但 `browser.ensure` 在 `setup` 里拿到的窗口句柄,`each` 里的
  `screen.capture_window` 根本引用不到(跨段);Excel COM 对象(P4 的 16 个
  step 全靠它)更不可能塞进 JSON 模板或 trace。规格里 §8 示例的
  `screen.capture_window` 没有任何窗口输入 —— 它隐式依赖「当前前台窗口」,
  这正是要消灭的 `$Global:Shell` 换了个马甲。不定这条,25 个 step 的实现者
  只能各自偷偷用全局变量,P4 时已积重难返。
- **做**:定义 `$Ctx.Session`:运行期命名资源注册表(browser 窗口句柄、Excel app、
  打开的工作簿)。规则:(1) 句柄/COM 对象**只进 Session,永不进 outputs**;
  (2) outputs **必须 JSON-可序列化**(P0-06 检查器强制 —— 这也是 P0-R3 输出重放
  的前提);(3) manifest 用 `provides` / `needs` 声明谁注册、谁消费哪个资源,
  `ebi lint` 静态检查「用了 browser 资源但工作流里没人 ensure 过」;
  (4) `excel.open` 类 step 用 `with.as = <名>` 注册命名资源,后续 step 用名字引用。
- **完成**:STEP-CONTRACT 新增 Session 一节;P0-08 的三步链(ensure→capture)能够
  不靠全局变量、不靠跨段 steps 引用写出来

### [x] P0-R3 [规格修订] 断点续跑 = ledger 输出重放 + setup 重跑
- **估** 60min | **依赖** P0-R2 | **改** `spec/STEP-CONTRACT.md` §6;`spec/WORKFLOW-SCHEMA.md` §7
- **问题**:「重跑按 ledger 跳过已完成的 (item, step)」—— 但跳过 `shot` 之后,
  下一步 `screen.crop` 引用的 `{{steps.shot.out.path}}` 从哪来?规格没说。
  不定义,P1-04 的实现者只能自己发明一套,P2-06 的断点续跑验收(中途 Ctrl+C)
  必然返工。`once: group` 更是双重没定义:它的 ledger 键是什么?组内后续 item
  引用它的输出算不算「在前面」?
- **做**:定死三条:(1) ledger 每条记录**连同 outputs 一起持久化**,resume 时被
  跳过的 step 其 outputs 从 ledger **重放**进模板作用域;(2) `setup` / `teardown`
  在每次 resume 都**重跑**(负责重建 Session 资源),因此 setup step 必须幂等;
  (3) `once: group` 的 ledger 键是 **(group, step)**,outputs 对组内所有 item 重放。
  在 spec 里加一个「中断发生在 shot 与 crop 之间」的完整推演例子。
- **完成**:P1-04 可照抄此节实现;推演例子覆盖同段引用、跨 item、once:group 三种情况

### [x] P0-R4 [规格修订] key 单一事实源 + 文件名安全形 + 学习规则落盘
- **估** 60min | **依赖** — | **改** `spec/PROFILE-SCHEMA.md` §2,§6;`spec/VOCABULARY.md`;`spec/WORKFLOW-SCHEMA.md` §8
- **问题**:四个会互相放大的小洞:(a) key 被声明了**两次** ——
  vocabulary.json `columns.key`(单列)和 worklist.json `key.columns`(复合数组),
  必然漂移;(b) 复合键下 `{{item.key}}` 求值成什么(拼接符?)没定义,而它直接
  进文件名(`capture/.../{{item.key}}.png`)—— key 含路径非法字符或全角时就出事;
  (c) 歧义面板确认后追加的 confirmedRules 写到哪没说(git 里的 profile?办公 PC
  上的部署副本怎么同步回来?);(d) 两份 spec 已经漂移:WORKFLOW-SCHEMA 示例用
  `profile.key.aliases`,PROFILE-SCHEMA 里叫 `key.confirmedRules`。
- **做**:worklist.json 是 key 的**唯一事实源**(删除 vocabulary 的 `columns.key`);
  定义 `{{item.key}}`(显示形,定死拼接符)与 `{{item.keySafe}}`(文件名安全形,
  定死编码规则,文件/目录名一律用它);学习规则:运行时先落
  `<WorkDir>/ebi.local.json`,面板提示「已学到 1 条规则,记得回填 profile 并提交」,
  `ebi profile check` 检测未回填的本地规则;统一命名为 `confirmedRules`。
  同时定义**候选列表的标准形状**(candidate + evidence{} + suggestion + doubts),
  `table.key` / `file.find` / `file.newest` / `verify.match_record` 返回它,
  `human.choose` 渲染它 —— 一个形状,不许四家各造。
- **完成**:两份 spec + 两个示例一致;P1-21 / P1-22 / P1-27 / P1-34 可直接引用

### [x] P0-R5 [规格修订] 按失败种类的容错策略 + warnings 通道
- **估** 60min | **依赖** — | **改** `spec/STEP-CONTRACT.md` §2,§3;`spec/WORKFLOW-SCHEMA.md` §6
- **问题**:(a) onError 策略是每 step 一刀切:`browser.wait_for` 的 `timeout`
  该 retry,`not_found` retry 毫无意义(还会对着错误页面连打三轮键盘);失败种类
  明明已在 `manifest.failures` 里穷举,策略却完全用不上它 —— P2 对拍一定会先
  撞上这堵墙,然后被迫在 runner 里打补丁。(b) step 返回值只有 ok/failure 两态,
  **非致命异常没有标准通道** —— 「3 行未识别」这类必须上报的警告(本仓库最痛的
  静默丢行事故)没有落点,P1-30 只能发明私有输出字段,runner/trace/汇总都看不见它。
- **做**:(a) `manifest.failures` 从字符串数组升级为
  `@( @{ id='timeout'; transient=$true }, ... )`;`retry` 只重试 `transient` 的失败,
  非 transient 直接降到 `ask`;`onError` 增加可选 `byFailure: { <id>: <policy> }`。
  (b) 返回值约定增加可选 `warnings = @( @{ code=; message=; data= } )`;runner
  **必须**把 warnings 写进 trace、计入 run 末尾汇总、在 `--guided` 面板即时显示。
  `internal_error` 声明为保留失败 id,manifest 不必列出。
- **完成**:两份 spec 更新;P1-30 的「未识别行」改用 warnings 表达

### [x] P0-R6 [规格修订] 模板 page 绑定 + profile 内模板的求值规则
- **估** 60min | **依赖** P0-R1 | **改** `spec/WORKFLOW-SCHEMA.md` §1,§4;`spec/PROFILE-SCHEMA.md` §5
- **问题**:(a) 模板禁止嵌套(这条是对的),但代价是 workflow 里只能写死
  `{{profile.pages.list.url}}` 这样的完整路径 —— 换一个 page 就要全文替换路径段,
  「换工作时 workflow 原样抄」的核心卖点直接破产,还会催生一批只差一个路径段的
  复制粘贴工作流,恰好复刻旧工具的病。(b) rules.json 示例里出现了
  `{{run.window}}`:profile 数据里的模板到底求不求值、何时求值,规格没说;
  而且 `run.window` 根本不在 run 作用域的枚举里(runId/startedAt/operator/workDir),
  是个未定义引用。
- **做**:(a) workflow 顶层新增 `"page": "<page名>"` 绑定,模板新增 `{{page.X}}`
  作用域 = `profile.pages[<当前 page>].X`,并让 `{{page.grammar}}` / `{{page.rules}}`
  解析到 grammar.json / rules.json 的同名条目;嵌套模板依旧禁止。
  (b) 定死:经 `{{profile...}}` / `{{page...}}` 取出的子树在传给 step 前
  **递归求值一次**;时间窗正式加入 run 作用域(由 human.input 或 CLI
  `--time-window` 写入,接线见 P2-07)——**执行时改名为 `run.timeWindow`**,
  不叫 `run.window`(和 `profile.window`、Session 窗口句柄名撞概念,
  复核时发现,见 PR #141 的审查记录)。
- **完成**:WORKFLOW-SCHEMA §8 示例改写后,换 page 只改一行;lint 检查项同步(P1-08)

### [x] P0-R7 [规格修订] 冻结标签改名
- **估** 10min | **依赖** — | **改** `BACKLOG.md`(P0-01 卡本身)、`Plan.md`
- **问题**:P0-01 要 `git tag spec/gift-gfix <tip>`,但远端**已经存在一个同名分支**
  `refs/heads/spec/gift-gfix`(旧 `docs/Generalization-Roadmap.md` 计划留下的
  冻结/热修分支,指向 `0f5343e`,PR #103,`git ls-remote` 验证过仍然存在)。
  git 允许同名 branch + tag 共存,但之后 `git checkout spec/gift-gfix` 会变成
  歧义引用,`git show spec/gift-gfix` 也会警告。
- **做**:P0-01 的标签名改成 `freeze/pre-ebi-dance`,卡里加一行说明为什么不用
  `spec/gift-gfix`(避免后人再踩同一个坑);同步 `Plan.md` 里全部提到
  `spec/gift-gfix` 的地方(D1 决策表、§11 P0 步骤、§14 风险表)。
  `docs/Generalization-Roadmap.md` 自己对 `spec/gift-gfix` 分支的引用不动——
  那是另一份仍然有效的计划的产物,不是这次要改的东西。
- **完成**:BACKLOG / Plan.md / `docs/README.md` 里不再出现 `spec/gift-gfix` 这个
  **标签**名(分支名本身当然还在,不受影响)

### [x] P0-R8 [规格修订] Plan.md 与 README.md 同步
- **估** 90min | **依赖** P0-R1…R7 | **改** `Plan.md` §4.3/§9/§10.3/§11/§14、
  `docs/README.md`、`INTERVIEW.md`
- **问题**:PR #140 只改了 `BACKLOG.md`,`Plan.md` 和 `docs/README.md` 没跟上,
  当时互相矛盾:`Plan.md` §4.3/§9/§11 还写着 `<side>.<role>.<verb>` 和
  `before.list.capture`(R1 已推翻的口径);§10.3 的卡数表是「56/6、92/8」而
  BACKLOG 已经是「64/更多」;§9 的 `ebi explain` 样例用旧 workflow id 和
  `pagetext/before_list/`(连 `pagetext/` 这个目录名本身都是过时的,
  VOCABULARY.md 从来只有 `capture/`);`docs/README.md` 里 `P0-R` 和 `64`
  出现次数为 0。`INTERVIEW.md` 也有三处同款漂移(`key.aliases`、pages.json
  按 role 描述、页面盘点提示词没提醒 page≠role)。
- **做**:R1…R7 每改一处规格,同步检查并修正 `Plan.md`/`docs/README.md`/
  `INTERVIEW.md` 里因此过期的内容;卡数表按 BACKLOG 实际卡数重新数一遍
  (发现 P2 也已经从 6 张长到 8 张,表格之前没跟上——这一条原始审查没提到)。
- **完成**:四份 spec + `Plan.md` + `docs/README.md` + `INTERVIEW.md` +
  `BACKLOG.md` 之间没有再发现矛盾的 id 命名 / 卡数 / 标签名

### [x] P0-R9 P0-00 状态修正
- **估** 5min | **依赖** P0-R1…R8 | **改** `BACKLOG.md`(P0-00 卡本身)
- **问题**:`P0-00 契约与词汇定稿` 一直标着 `[x]` 已完成,但 P0-R1…R6 六个
  契约洞恰恰是评审在 P0-00 标完成**当天**就发现的——那个 `[x]` 从落笔起
  就是假的,一直没人回去改。
- **做**:R1…R8 全部执行、复核完毕后,把 P0-00 的状态说明补上这段历史
  (曾经是假 `[x]`,现在是真 `[x]`),而不是让下一个会话以为它从一开始
  就经得起审查。
- **完成**:P0-00 卡文本里能看到这段状态修正记录;本文件顶部「状态」段落
  和 P0 分组标题的卡数统计已更新(17 张,2 张整块)

### [x] P0-R10 [整块][规格修订] Session 资源的生命周期(释放侧)
- **估** 75min | **依赖** P0-R2 | **改** `spec/STEP-CONTRACT.md` §2.1,§3.4,§4,
  §6.2,§6.3,§7;`spec/WORKFLOW-SCHEMA.md` §1(新增 §1.1),§7.2,§7.4,§7.5,§9
- **问题**:P0-R2 只定义了 Session 资源的**注册侧**(`provides` 声明种类、
  `with.as` 注册实例名、`type='session'` 输入消费),**释放侧从没定义**,
  留下三个洞:(a) P4-01 卡标题写着 `excel.close`,但契约里没有任何一条
  规则要求它必须出现在 `teardown` 里——写不写、由谁触发,全凭实现者自觉。
  (b) `WORKFLOW-SCHEMA.md` §1 说 teardown「整个 run 结束后跑一次(含失败
  退出)」,但没说 Ctrl+C 算不算「失败退出」——而 §7.5 的断点续跑推演
  例子字面意思就是一次 Ctrl+C 中断,例子里完整过了一遍 setup 重跑、
  ledger 重放,却从头到尾没提 teardown,也没提这次中断对 Session 里已经
  注册过的资源意味着什么——如果 Ctrl+C 不保证跑 teardown,重复
  Ctrl+C/resume 循环会不会在系统里堆出一串没人关掉的 Excel COM 进程,
  规格没有答案。(c) P0-R2 定的「`provides` 非空但不写 `as` 是合法的」
  这条规则,副作用是造出了一批**注册不了、因而也释放不了**的资源:
  `excel.close` 要靠名字在 Session 里找到要关的工作簿,没写 `as` 的资源
  没有名字——这条规则对窗口句柄(泄漏无害)没问题,对 COM 对象(泄漏会
  累积、需要人工杀进程)是个漏洞。(d,第四轮追加)**(a)(b)(c) 三条决定
  堵上了"资源怎么释放",但没堵上"同一个名字被重复注册,旧的那个悄悄
  没人管了"**:`with.as` 的值是字符串字面量、不走 `{{}}` 模板(§3.4 第 3
  点),所以注册名不能随 item/group 变化;而"名字的作用域是一次 run"
  (同一节)又明说 `each` 里可以注册。`groupBy` 分组的 compose 类工作流
  典型写法是「每个交付物开一个工作簿」:`once:"group"` 在每组开头调
  `excel.open` 注册同一个名字(比如 `wb`)。第 2 组的 `open` 一跑,
  `$Ctx.Session['wb']` 被同名覆盖,第 1 组那个 Workbook/Application COM
  对象当场变成孤儿——没有名字能传给 `excel.close`。而 `WORKFLOW-SCHEMA.md`
  §7.2 现在只有 `once:"group"` 这个"组开头"钩子,没有"组结束"钩子,连
  想在组尾释放都做不到,只能拖到 `teardown`——但 `teardown` 只跑一次,只
  关得掉最后一组,前面全部泄漏。§9 现在那条 lint(检查 `teardown` 里有
  释放调用)是**满足**的,照样泄漏。不是假想:`ReplaceEvidence.ps1:148`
  就是 `Group-Object Excel_NAME`,`P4-20 workflows/*.compose.json` 是它的
  替代品。"不注册、自己释放"这条路走不通——同一个 item 内部要跨好几个
  step 用同一个工作簿(open → find_sheet → insert_picture → save),不注册
  后面的 step 根本拿不到它。
- **做**:三条决定,不留 TBD——
  1. **谁负责释放:显式,不是 runner 自动**。manifest 新增可选字段
     `releases = @(<kind>, ...)`(与 `provides` 对称,默认 `@()`),声明
     这个 step 会释放 `$Ctx.Session` 里哪个种类的资源(它照样通过自己
     `inputs` 里 `type='session'` 的参数拿到具体实例名)。工作流作者
     必须在 `teardown` 里显式调这类 step(比如 `excel.close`)。不选
     runner 自动释放的理由:不同种类资源的释放顺序/方式完全不同(工作簿
     要先 `Close` 再让 App `Quit`、句柄类资源什么都不用做)——让 runner
     替 Session 里每个种类内置一套释放逻辑,等于让 runner 替工作流做
     「计算」,违反 `WORKFLOW-SCHEMA.md` §0 的设计铁律;而且这本来就是
     本仓库 `ExcelHelpers.ps1` 的既有写法(`Close-Workbook`/
     `Close-ExcelApp` 从来是调用方显式调用,从没有框架自动挡在中间)。
     对应 `ebi lint` 检查(§9 新增项):某个种类如果在 catalog 里**存在**
     带 `releases` 覆盖它的 step,那么任何一次 `with.as` 注册过这个种类
     的调用,`teardown` 里就必须有一次对**同名**实例的释放调用;种类在
     catalog 里根本没有 `releases` 覆盖(比如 `window`)→ 不要求。
  2. **异常退出路径,穷举,不用"含失败退出"这种含糊说法**:正常跑完、
     `onError.policy=fail` 中止、step 抛出未预期异常(`internal_error`)
     ——这三种都发生在同一个 PowerShell 进程的正常控制流里,runner 用
     `try { ... } finally { 跑 teardown }` 包住整条执行路径就能保证,
     **这三种 teardown 保证跑**。**Ctrl+C(以及被杀进程/终端被关/系统
     重启这类硬中断)——不保证**:PowerShell 5.1 的 Ctrl+C 默认直接
     终止进程,不触发 `finally`;就算 runner 注册 `CancelKeyPress` 尽力
     兜底,中断到达时线程可能正卡在一次还没返回的 COM 调用里,teardown
     想开始跑都进不去。**泄漏怎么办:接受泄漏,不做孤儿检测,但要有
     文档警告**——不新增任何"下次启动扫描孤儿进程"的基础设施,如实
     写清楚这就是本项目 Excel COM 场景一直以来的真实运维方式(操作员
     手动在任务管理器里杀多余的 `EXCEL.EXE`)。
  3. **`STEP-CONTRACT.md` §3.4 第 3 点里"没有默认名这回事……"这句后面
     加一条约束**:没注册(没写 `as`)的资源,必须在产生它的这次 step
     调用**返回之前**由 step 自己释放完——它没有名字,后面没有任何
     step 能引用到它。任何需要跨 step 存活、或者需要显式释放的资源
     (尤其是 COM 对象)都必须注册。
  4. **(第四轮)同一个名字重复注册,如果它当前仍然活着,是运行期
     失败,不是静默覆盖**。runner 在真正执行 `with.as: "<名>"` 之前先查
     `$Ctx.Session` 里这个名字是否已经存在,存在就直接失败,不进
     `Invoke-Step`。不选"runner 自动释放旧的再注册新的"——理由和第 1 点
     否掉"runner 自动清理"一样:不同种类资源释放方式不同,不该让 runner
     替工作流做这个决定。
  5. **(第四轮)新增 `"once": "groupEnd"`,和 `"once": "group"` 对偶**:
     在同组最后一条 item 处理完之后跑一次,专门释放 `once:"group"` 在
     组开头注册的组级资源。ledger 键和 `once:"group"` 一样是
     (group, step);`source.groupBy` 没设时用它是配置错误,`ebi lint`
     报错。于是 compose 工作流的正确写法是 `open` 用 `once:"group"`、
     `close` 用 `once:"groupEnd"`,两者在 `each` 段内配对;`teardown`
     只管 `setup` 里注册的东西,不再兼管组级资源。
- **完成**:`STEP-CONTRACT.md` 新增 `releases` 字段(manifest 骨架 + §2.1
  字段表 + §3.4 新增第 5 点 + §4 needs/sessionKind 去重说明 + §6.2 补
  teardown 幂等要求 + §7 Run-Tests.ps1 清单新增一条);`WORKFLOW-SCHEMA.md`
  新增 §1.1"退出路径 x teardown 保证"表格、§7.5 补一段 Ctrl+C 场景下
  Session 资源的说明、§9 `ebi lint` 清单新增一条;`BACKLOG.md` P0-06 /
  P1-08 两张卡各追加一条对应的检查项。`grep -c 'releases' spec/STEP-CONTRACT.md
  spec/WORKFLOW-SCHEMA.md` 两个文件合计 ≥ 8 处命中(字段定义、示例骨架、
  §2.1 表格、§3.4 决定、§4 衔接句、§6.2 幂等句、§7 清单、§9 清单——不是
  只在一处提了一句就算数)。**(第四轮追加)**:`grep -rn 'groupEnd'
  docs/ebi-dance/` 命中 `STEP-CONTRACT.md`(§3.4 第 3/5 点、§6.3)、
  `WORKFLOW-SCHEMA.md`(§7.2 定义+例子、§9 两条 lint)、`BACKLOG.md`
  (本卡 + P1-04 + P1-08 + P4-01 + P4-20)——不是只在 §7.2 写了语法;
  `STEP-CONTRACT.md` §3.4 新增第 6 点的 `mustRelease` 种类表存在,
  `WORKFLOW-SCHEMA.md` §9 的释放判据条目改成读这张表,不再读"catalog
  里有没有恰好带 `releases` 的 step"。

### [ ] P0-01 打冻结标签
- **估** 10min | **依赖** — | **读** `Plan.md` §11 P0
- **做**:`git tag freeze/pre-ebi-dance <当前 main tip>` 并推送。作为整个重构期的回滚点。
- ⚠ **不用 `spec/gift-gfix`**:远端已经存在一个同名**分支**
  `refs/heads/spec/gift-gfix`(指向旧提交 `0f5343e`,PR #103)。git 允许
  同名 branch + tag 共存,但那样 `git checkout spec/gift-gfix` 会变成
  歧义引用,`git show spec/gift-gfix` 也会警告 —— 换成 `freeze/pre-ebi-dance`
  彻底避开冲突,不要图省事换回 `spec/gift-gfix`。
- **完成**:远程能看到该 tag;`git show freeze/pre-ebi-dance --stat` 正常

### [ ] P0-02 建目录骨架
- **估** 30min | **依赖** P0-01 | **读** `Plan.md` §3
- **做**:建 `kernel/ modules/{browser,screen,file,excel,table,verify,human} workflows/ profiles/ legacy/`,
  每个目录放一个 `README.md` 说明放什么。更新根 `.gitignore` 加 `.ebi/`、`run/`、`capture/`。
- **完成**:目录存在;`git status` 干净;每个目录的 README 说清楚"什么该进来、什么不该"

### [ ] P0-03 kernel/Trace.ps1
- **估** 60min | **依赖** P0-02 | **读** `spec/STEP-CONTRACT.md` §3.2
- **做**:从 `ProgressLog.ps1` 搬过来并**字段泛化** —— 去掉硬编码的 `correl_id_s` / `job_name`,
  改成 `key` + `tags{}`。保留 UTF-8 无 BOM 追加写(`UTF8Encoding($false)`,`Set-Content -Encoding UTF8` 会在每次追加时插 BOM,毁掉 jsonl)。
- **完成**:`Write-TraceEvent` / `Read-TraceEvents` 可用;单测覆盖「追加 3 条后读回 3 条且无 BOM」

### [ ] P0-04 搬纯函数库进 modules/verify/
- **估** 45min | **依赖** P0-02 | **读** `Plan.md` §12
- **做**:`SnapVerify.ps1` `GfixLog.ps1` `GfixJobList.ps1` `ScreenRegion.ps1` `OwnerFilter.ps1`
  原样搬进 `modules/verify/`(**这一步不改任何逻辑**,只挪位置 + 改 dot-source 路径)。
  对应的 `Tests/Test-*.ps1` 跟着改路径。
- **完成**:`Tests/Run-Tests.ps1` 全绿,和搬之前的测试数一致

### [ ] P0-05 归档 legacy/
- **估** 30min | **依赖** P0-02 | **读** `Plan.md` §6.3
- **做**:`TimeDigitVerify.ps1` `PixelDigitMatch.ps1` `OldSnapPixelVerify.ps1` `OldSnapVerify.ps1`
  移入 `legacy/`,加 `legacy/README.md` 说明:**这些只用来清历史老快照,清完即弃,
  不进 catalog,不许被新工作流依赖**。测试跟着移。
- **完成**:测试全绿;`legacy/README.md` 写清楚退役条件

### [ ] P0-06 Tests 适配 + 契约检查器
- **估** 90min | **依赖** P0-04, P0-05, P0-R2, P0-R5 | **读** `spec/STEP-CONTRACT.md` §7
- **做**:`Tests/Run-Tests.ps1` 支持新目录树;新增 `Tests/Test-StepContract.ps1`,对
  `modules/**` 的每个 step 检查:能 dot-source、无 `param()`、`$Manifest.id` == 文件名、
  `required` 与 `default` 不共存、`failures` 非空、`example` 的参数都声明过、源码纯 ASCII。
  评审追加的检查:**outputs 声明的类型必须 JSON-可序列化**(句柄/COM 走 Session,
  P0-R2);`failures` 每项有 `transient` 布尔(P0-R5);step 文件里除 `Invoke-Step`
  外的辅助函数**必须带 step 前缀**(如 `BrowserFind-*`)—— 所有 step 会被同一
  runspace 依次 dot-source,`Invoke-Step` 靠注册表捕获解决(P1-02),裸名辅助函数
  则会互相覆盖且无人发现;`inputs` 里 `type='session'` 的参数都带 `sessionKind`
  (P0-R2);`provides` 最多一项(P0-R2 §3.4——一次调用最多注册一个资源);
  `releases` 声明的种类都能在该 step 自己的某个 `type='session'` 输入的
  `sessionKind` 里找到(P0-R10)。
- **完成**:对一个故意写错的 fixture step 能报出每一类错误(含新增六类)

### [ ] P0-07 [整块] 最小 runner spike
- **估** 90min | **依赖** P0-06, P0-R2 | **读** `spec/WORKFLOW-SCHEMA.md` §1-2
- **做**:**只做能跑通的最小版**:读 workflow JSON → 按顺序 dot-source 并调用 step →
  打印结果。不做模板求值、不做 foreach、不做 onError。**但 `$Ctx.Session` 从
  第一天就要在**(哪怕只是个空 hashtable)—— spike 的目的就是验证 ensure→capture
  的句柄传递走 Session 而不是全局变量。
- **完成**:能跑一条只有 `setup` 三步的 JSON

### [ ] P0-08 三个 step + 端到端验收
- **估** 90min | **依赖** P0-07 | **读** `spec/STEP-CONTRACT.md` §8(完整示例)+ Session 节(P0-R2)
- **做**:`human.prepare`(从 `Common.ps1 Wait-PagePrepared`)、`browser.ensure`
  (从 `Common.ps1 Activate-EdgeWindow`,进程句柄优先/标题回退)、
  `screen.capture_window`(从 `Common.ps1 Take-WindowScreenshot`)。
- **完成**:**办公 PC 上,一条 5 行的 workflow JSON 真的存下一张 PNG**;
  窗口句柄经 `$Ctx.Session` 流转,`grep -rn 'Global:' modules/` 为 0,
  三个 step 的 outputs 全部可 `ConvertTo-Json`
- ⚠ 这是 P0 的唯一验收标准。做不到就别进 P1。

---

# P1 — 内核 + 25 个 MVP step + 文档生成(34 张,4 张整块)

## kernel(6 张)

### [ ] P1-01 [整块] kernel/Context.ps1
- **估** 90min | **依赖** P0-08, P0-R6 | **读** `spec/WORKFLOW-SCHEMA.md` §4
- **做**:`{{}}` 模板求值。作用域 `vars` / `profile` / `run` / `item` / `steps.<id>.out.<f>`
  / **`page`(P0-R6 的绑定间接)** / **`group`(分组遍历时,§7.2 用到但初版作用域表漏了)**。
  **纯函数,先写单测再写实现。**
  规则:只有取值和字符串拼接,**没有运算**;整个值就是一个 `{{}}` 时保留原类型;
  `\{\{` 转义;引用不存在的路径要能报出**具体是哪一段**解析不到;
  profile/page 子树**递归求值一次**(P0-R6);复合键的 `{{item.key}}` / `{{item.keySafe}}`
  按 P0-R4 的定义展开。
- **完成**:单测覆盖 7 种作用域 + 类型保留 + 转义 + 递归求值 + keySafe + 3 种解析失败的报错信息

### [ ] P1-02 kernel/Registry.ps1
- **估** 75min | **依赖** P0-06 | **读** `spec/STEP-CONTRACT.md` §2
- **做**:扫描 `modules/**`、加载 `$Manifest`、按 `inputs` schema 校验一次调用的参数
  (类型、required、enum、default 填充)。纯函数,单测。
  **加载机制**:所有 step 文件定义同名 `Invoke-Step`,依次 dot-source 会互相覆盖 ——
  Registry 必须在每次 dot-source 后**立刻**把 `${function:Invoke-Step}` 的 scriptblock
  捕获进按 id 索引的表,运行期从表里调,不再二次 dot-source。
- **完成**:对缺 required、类型不符、未声明的多余参数,都能报出**参数名**;
  加载两个 step 后各自的 Invoke-Step 仍能正确调用(单测)

### [ ] P1-03 [整块] kernel/Runner.ps1 主体
- **估** 120min | **依赖** P1-01, P1-02 | **读** `spec/WORKFLOW-SCHEMA.md` §1,3,7
- **做**:`setup` / `each` / `teardown` 三段;`source.select` 的五种 `pendingWhen`;
  `flow.foreach`(隐式)、`flow.if`(`when` 的四种形式)、`once: group`
  (outputs 对组内后续 item 可见,按 P0-R3 的重放规则)。
  `source.select` 的筛选实现和 `table.select` step **共用同一个函数**,不写两份。
- **完成**:能跑通一条有 setup+each 的 JSON,遍历 3 行 fixture 数据;
  含一条 `once: group` 的用例(组内第 2 个 item 能引用第 1 个 item 时跑出的输出)

### [ ] P1-04 [整块] Runner 的 onError + ledger
- **估** 120min | **依赖** P1-03, P0-R3, P0-R5 | **读** `spec/WORKFLOW-SCHEMA.md` §6;`STEP-CONTRACT.md` §6
- **做**:四种 policy(`retry` 退避 / `ask` / `skip` / `fail`)+ **`byFailure` 按失败
  id 覆盖;`retry` 只重试 manifest 标了 `transient` 的失败**(P0-R5);
  `destructive` 自动插确认关卡;step 返回的 `warnings` 进 trace + 末尾汇总;
  ledger 写 `run/<runId>/ledger.jsonl`,粒度 **(item, step)**(`once: group` 为
  (group, step)),**每条连同 outputs 持久化**;重跑跳过已完成的并**重放其 outputs**,
  `setup`/`teardown` 每次 resume 重跑(P0-R3)。**(第四轮追加)**
  `once: "groupEnd"`(`WORKFLOW-SCHEMA.md` §7.2,P0-R10)在同组最后一条
  item 处理完之后触发一次,ledger 键同样是 (group, step);runner 按
  `groupBy` 排序遍历 `source`,组切换(或整个遍历结束)时触发上一组的
  `groupEnd`。**运行期同名重复注册检查**(`STEP-CONTRACT.md` §3.4 第 3
  点,P0-R10):执行 `with.as: "<名>"` 之前,先查 `$Ctx.Session` 里这个
  名字是否已注册且尚未释放,是则直接失败(不进 `Invoke-Step`),不静默
  覆盖——这是运行期检查,`ebi lint`(P1-08)静态走一遍 JSON 时看不出
  `once:"group"` 会在运行时对同一个名字重复调用几次,只能检查组头/组尾
  的注册-释放配对结构对不对,不能替代这条运行期检查。
- **完成**:中断后重跑不重复执行已完成的 (item, step),且被跳过 step 的输出仍可被
  后续步引用;`timeout`(transient)会 retry 而 `not_found` 不会;`confirm:false`
  能跳过自动关卡;`once:"groupEnd"` 在同组最后一条 item 后触发且仅触发一次;
  对一个已注册且未释放的名字重复 `with.as` → 运行期失败,不静默覆盖

### [ ] P1-05 kernel/Gate.ps1
- **估** 75min | **依赖** P1-03 | **读** `Plan.md` §3.3
- **做**:统一的 ASCII 关卡面板 —— **发生了什么 / 下一步会做什么 / 证据在哪 / 可选动作**。
  替代现在 27 个文件、77 处各写各的 `Read-Host`。
- **完成**:面板在 80 列终端下不折行;`r/s/q/m` 四个动作都通

### [ ] P1-06 kernel/Docs.ps1
- **估** 75min | **依赖** P1-02 | **读** `Plan.md` §5
- **做**:扫 manifest → 生成 `docs/ebi-dance/CATALOG.md`(人读)+ `docs/ebi-dance/catalog.json`(Agent 读)。
- **完成**:两份产物都生成;CATALOG.md 按 group 分节;catalog.json 能被 `ConvertFrom-Json` 读回

## CLI(4 张)

### [ ] P1-07 ebi help
- **估** 45min | **依赖** P1-06 | **做**:`ebi help` 分组列全部 step;`ebi help <id>` 渲染单个 manifest
- **完成**:输出纯 ASCII,80 列不折行

### [ ] P1-08 ebi lint
- **估** 90min | **依赖** P1-01, P1-02, P0-R6 | **读** `spec/WORKFLOW-SCHEMA.md` §9
- **做**:§9 的 9 项静态检查全实现,包括 fallback tier 警告和 `confirm:false` 警告。
  评审追加:`page` 绑定解析得到(P0-R6);`inputs` 的 `sessionKind`/`provides`
  的 Session 资源配平(「用了 browser 没人 ensure」,不读 `needs`——P0-R2 的
  配平算法本来就只看 `type='session'` 输入,P0-R10 把 `needs:session:<kind>`
  从 manifest 里整个删掉之后更是如此);`byFailure` 引用的失败 id 在 manifest 里
  存在(P0-R5);`with.as` 注册过、且种类在 catalog 里有 `releases` 覆盖的资源名,
  `teardown` 里必须有对应的释放调用(P0-R10)。
- **完成**:对一份故意写错的 workflow,全部检查项都能报出来

### [ ] P1-09 ebi explain
- **估** 90min | **依赖** P1-03 | **读** `Plan.md` §9(输出样例)
- **做**:渲染成 ASCII 执行计划,标出每步的 effects、人工关卡数、破坏性操作数、是否用到降级层
- **完成**:输出和 `Plan.md` §9 的样例形状一致;显示 `list(転送状態一覧)` 这种中性名+显示名

### [ ] P1-10 ebi dryrun / run / doctor
- **估** 75min | **依赖** P1-04 | **做**:三个子命令接线;`doctor` 检查 PS 版本、Excel COM、Edge、编码策略
- **完成**:`dryrun` 不碰真实系统就能走完全流程

## browser 组(7 张,11 个 step)

### [ ] P1-11 browser.ensure + browser.focus_body
- **估** 60min | **抄** `Common.ps1` `Activate-EdgeWindow` / `Click-PageBody`
- **注意**:进程句柄优先、标题匹配只做回退、两条路都失败要 `[WARN]`(旧版静默"激活"了随便哪个前台窗口)
- **完成**:manifest 过 lint;dryrun 打印正确

### [ ] P1-12 browser.send_keys + tab_to + fill + submit
- **估** 75min | **抄** `Common.ps1` `Send-Key` / `Send-Tab` / `Send-ShiftTab` / `Paste-Replace` / `Send-Enter`
  ⚠ 少不了 `Send-ShiftTab`(`Common.ps1:174`):`spec/PROFILE-SCHEMA.md` §3.0
  写明 HM 的按键序列是 `Tab n → 粘贴 → Shift+Tab m → 回车`,没有它这条序列
  实现不出来
- **注意**:时序参数(`waitMs`)走 step 输入,**不要用 `$Global:Timing`**
- **完成**:4 个 manifest 过 lint;全局变量依赖为 0

### [ ] P1-13 browser.read_text
- **估** 45min | **抄** `Read-PageText.ps1`
- **完成**:能把 Ctrl+A 文本返回,并可选归档到指定路径

### [ ] P1-14 browser.wait_for
- **估** 75min | **抄** `MqSnap.ps1 Wait-MqPageReady`(**去掉 MQ 特有的硬编码**)
- **做**:轮询页面文本直到 `contains` 匹配或超时;可选 `archiveTo` 同时留档
- **注意**:归档是**强制的最佳实践** —— 有文本就永远不用 OCR(`Plan.md` §6.1)
- **完成**:超时返回 `failure='timeout'` 而不是抛异常

### [ ] P1-15 browser.assert_page
- **估** 60min | **抄** `SnapVerify.ps1 Get-SnapPageKind`;**读** `spec/PROFILE-SCHEMA.md` §3.1
- **做**:按 fingerprint 判 `ok` / `loading` / `empty` / `expired` / 未知
- **注意**:**未知页面必须失败,绝不允许继续截图** —— 这是最坏的一类失败(看起来成功)
- **完成**:5 种页面状态各有一个 fixture 单测

### [ ] P1-16 browser.navigate
- **估** 45min | **做**:Ctrl+L 粘贴 URL 回车;URL 为空时降级为提示人工打开

### [ ] P1-17 browser.find
- **估** 60min | **做**:Ctrl+F 查找**精确串**,返回是否命中;可选 Esc 关闭
- **注意**:Ctrl+F 是**子串搜索,会停在页面列出的第一行** —— 所以调用方必须传完整的、
  已经由 `verify.match_record` 选定的那一行的标识,不能传裸 key(旧工具在这栽过)

## screen 组(4 张,5 个 step)

### [ ] P1-18 screen.capture_window + capture_region
- **估** 60min | **抄** `Common.ps1 Take-WindowScreenshot` + `ScreenRegion.ps1 Resolve-ScreenRegion`
- **完成**:region 越界时自动 clamp 并在返回值里报告 clamped

### [ ] P1-19 screen.fit_window
- **估** 45min | **抄** `MqSnap.ps1 Move-EdgeAwayFromBorder` + `WinAPI MoveWindow`

### [ ] P1-20 screen.crop —— **消掉 4 份重复**
- **估** 60min | **抄** `ScreenRegion.ps1 Resolve-DirectionalCrop` + 任一份 `Invoke-CropPng`
- **做**:四边裁剪;per-role 覆盖走 profile
- **完成**:`grep -c "function Invoke-CropPng" *.ps1` 在迁移完成后为 0(现在是 4)
- **参考**:`spec/STEP-CONTRACT.md` §8 就是这张卡的完整答案

### [ ] P1-21 screen.save
- **估** 45min | **做**:按命名模板定位保存;支持 `<keySafe>__<tag>.png` 的多张形式
  (文件名一律用 P0-R4 的 `keySafe`,不用裸 key)
- **读** `spec/VOCABULARY.md` §2.5

## file 组(2 张)

### [ ] P1-22 file.find
- **估** 75min | **依赖** P0-R4 | **抄** `WorkbookResolver.ps1 FullWidthFilenameResolver` + `MappingStore.ps1 Resolve-CorrelFilePath`
- **做**:glob/key 查找,全角回退 + key 变体容忍(规范化调 `kernel/Key.ps1`,见 P1-27,
  自己不写比较)。**匹配到多个时按 P0-R4 的标准候选形状返回全部候选 + 证据**,不自己挑
- **完成**:同名多文件时返回标准候选数组而不是单个

### [ ] P1-23 file.assert_exists
- **估** 30min | **做**:存在性断言,不存在按策略走 gate

## table / progress 组(6 张)

### [ ] P1-24 table.load + table.save
- **估** 60min | **依赖** P0-R4 | **抄** `MappingStore.ps1 Import-Mapping` / `Export-MappingAtomic`
- **注意**:CSV 是 UTF-8 **带 BOM**(Excel 需要);写入必须原子(临时文件 + 改名)。
  `table.load` 必须对全表算一遍 `keySafe`(`PROFILE-SCHEMA.md` §6.6),撞车的行
  直接判失败并列出来——`keySafe` 的规范化规则本身会制造新的重名
  (`A_B`+`C` 和 `A`+`B_C` 都拼成 `A_B_C`),不在加载时挡住就会在 capture
  阶段静默互相覆盖截图

### [ ] P1-25 table.ensure_columns
- **估** 45min | **抄** `MappingStore.ps1 Ensure-MappingColumns`;列 schema 来自 profile

### [ ] P1-26 table.select
- **估** 60min | **抄** `MappingStore.ps1 Get-PendingRows`
- **注意**:`ng` **仍算 pending**(`spec/WORKFLOW-SCHEMA.md` §3.2)—— 旧的
  `Get-PendingRows` 把任何非 `0` 都当已完成,会把 NG 行藏起来。
  同一份筛选实现同时供 runner 的 `source.select` 用(P1-03),不写两份

### [ ] P1-27 [整块] kernel/Key.ps1 + table.key
- **估** 120min | **依赖** P1-26, P0-R4 | **读** `spec/PROFILE-SCHEMA.md` §6 全节
- **做**:核心是 **`kernel/Key.ps1` 纯库**(无 param(),可被任何 step dot-source):
  复合主键(`key.columns` 数组)规范化、`confirmedRules` 应用、候选排序 + 证据
  收集,歧义时按 P0-R4 的标准形状返回**全部候选 + 每个候选的全部证据**,附建议、
  理由、**不确定点**。`table.key` 只是它的 step 包装。
- **注意**:这是全项目最容易做错的一张。旧工具在**七个地方**各写 `-eq`,症状
  五花八门。规则必须**只有这一处** —— 做成库而不是只做成 step,正是为了让
  `file.find` / `file.newest` / `verify.match_record` / `excel.find_anchor` 能直接
  调它:step 不能调 step,但都能 dot-source 同一个 kernel 库。学习到的新规则按
  P0-R4 落 `<WorkDir>/ebi.local.json` 待回填。
- **完成**:单测覆盖 —— 单列键、复合键、后缀变体、全角、大小写、
  「4 个候选无法确定」返回完整候选表;`grep -rn '\-eq' modules/` 里没有 key 比较

### [ ] P1-28 table.set + flow.checkpoint
- **估** 60min | **抄** `MappingStore.ps1 Update-MappingRows` / `Set-MappingBit`
- **做**:位定义来自 profile 的 `bits`,不硬编码 1/2/4
- **注意**:`pendingWhen` 的位掩码写法从 `"bit !3"`(数字)改成 **`"bit !<位名>"`**
  (如 `bit !before`)—— checkpoint 用名字、pendingWhen 用数字是两套口径,
  必然抄错;顺手改 `spec/WORKFLOW-SCHEMA.md` §3.1

### [ ] P1-29 progress.event + progress.status
- **估** 60min | **抄** P0-03 的 Trace + `VerifyTool.ps1 Show-Status`
- **做**:ASCII 进度表

## verify 组(4 张)

### [ ] P1-30 verify.parse_text —— delimited
- **估** 75min | **依赖** P0-R5 | **抄** `GfixJobList.ps1 ConvertFrom-GfixJobListText`;**读** `spec/PROFILE-SCHEMA.md` §4
- **做**:分隔符表格,靠 `rowWhen` 正则识别数据行
- **⚠ 必须**:未识别行走 P0-R5 的标准 **`warnings` 通道**(带行数和内容),
  不发明私有输出字段 —— runner 才会把它进 trace、进末尾汇总(静默丢行是旧工具
  最恶劣的 bug,光「返回了」不够,必须**有人看见**)

### [ ] P1-31 verify.parse_text —— labeled + columns + regex
- **估** 90min | **抄** `SnapVerify.ps1 ConvertFrom-HmPageText` / `ConvertFrom-JenkinsListText`
- **⚠ 必须**:时间格式用 `H:mm:ss` 单字符说明符,**不要 `HH`**
  (单位数小时的行曾被整批静默丢弃,导致"文件不在列表里"的误判)
- **完成**:单位数小时的行有专门的回归单测

### [ ] P1-32 verify.match_record
- **估** 75min | **依赖** P1-27 | **抄** `SnapVerify.ps1 Get-MatchedRowIndex` / `Select-JenkinsFileCandidate`
- **做**:按 key 找行(dot-source `kernel/Key.ps1`,不自己写比较)、`tieBreak: newest`、
  多候选按 P0-R4 标准候选形状返回,走歧义流程

### [ ] P1-33 verify.assert
- **估** 90min | **读** `spec/PROFILE-SCHEMA.md` §5
- **做**:规则表引擎,`op` 的 12 种(`equals`/`notEquals`/`in`/`notIn`/`matches`/
  `present`/`empty`/`within`/`gt`/`lt`/`gte`/`lte`);**`else` 只能是 `ng` 或
  `unknown`,校验时拒绝 `ok`**
- **完成**:单测覆盖每种 op;`else: ok` 的规则表被拒绝并报错

## human 组(1 张)

### [ ] P1-34 human.prepare + human.gate + human.choose
- **估** 75min | **依赖** P1-05, P0-R4 | **做**:三个 step 接到 `kernel/Gate.ps1` 的面板
- **注意**:`human.choose` 渲染 P0-R4 的**标准候选形状**(file.find / table.key /
  verify.match_record 返回的是同一个形状,渲染器只写一份);`human.gate` 的
  outputs 要在 manifest 里声明(`action`: enter/n/s/q、`note`),后续 step 才能
  `when` 到它 —— 旧规格没定义 gate 的返回值

---

# P2 — 对拍验证(8 张,1 张整块)

目标:**用新引擎重跑一条现有流程,产出和旧脚本逐项一致。**
这一步会暴露契约的全部错误 —— **P1 的设计不必完美,P2 之后重构一次是计划内的。**

### [ ] P2-01 profiles/host-open 骨架
- **估** 60min | **依赖** P0-R1, P0-R4 | **读** `spec/PROFILE-SCHEMA.md` §1,2,6
- **做**:`vocabulary.json`(side/列名映射)+ `worklist.json`(列 schema、复合键、
  位定义;key 只声明在这里)+ `pages.json` 骨架 —— **按 page 名建条目**
  (如 `transferStatus`(role list)、`fileList`(role list)),不按 role 建

### [ ] P2-02 ebi grammar tune
- **估** 120min | **依赖** P2-08 | **读** `spec/PROFILE-SCHEMA.md` §4.1
- **做**:交互式解析器调试器 —— 喂真实页面文本 → 渲染解析结果表格 →
  改参数即时重解析 → `s` 存进 profile **同时存成 fixture**
- **⚠ 必须**:未识别的行显式列出,不能藏;`s` 保存 fixture 前**必须过 P2-08 的
  脱敏门禁**(fixture 的原料是真实内网页面文本 —— 这一步会自动积累敏感文件,
  掩码不能等到 P5)
- **完成**:调一次 grammar 自动留下一个回归测试

### [ ] P2-03 用 grammar tune 调出 list 页解析
- **估** 60min | **依赖** P2-02 | **做**:拿一份真实的 `list` 页 Ctrl+A 文本调到未识别行为 0

### [ ] P2-04 pages.json + rules.json
- **估** 90min | **抄** `SnapVerify.ps1 Test-MqRecord` 的判定语义**翻译成规则表**
- **⚠ 判定语义一行不改**,靠 `Tests/Test-SnapVerify.ps1` 的既有 fixture 护住

### [ ] P2-05 workflows/before.transferStatus.capture.json
- **估** 60min | **读** `spec/WORKFLOW-SCHEMA.md` §8(完整示例)
- **完成**:`ebi lint` 全绿;`ebi explain` 的输出人工逐行确认过

### [ ] P2-06 [整块] 办公 PC 首跑 + 对拍
- **估** 120min | **依赖** P2-05
- **完成判据(缺一不可)**:
  - [ ] 同一批 key,新旧两条路各跑一遍:PNG 尺寸/裁剪一致
  - [ ] CSV 标记逐项一致
  - [ ] 故意造一个 NG 页面,两边都判 NG
  - [ ] 中途 Ctrl+C,重跑从断点续上,不重复截图
- ⚠ 如果旧流程已无真实环境可跑,改用任意一条还能跑的。**对拍验证的是引擎,不是业务。**

### [ ] P2-07 human.input + run.timeWindow 接线
- **估** 60min | **依赖** P1-05, P0-R6
- **做**:`human.input` step(默认值 + 校验 + 批量一次问,抄旧 Expected_Time 批量
  提示的交互方式)+ CLI `--time-window`,写入 run 作用域的 `run.timeWindow`
  (形状 `{ "from": "<ISO8601>", "to": "<ISO8601>" }`,见
  `spec/WORKFLOW-SCHEMA.md` §4.1;字段叫 `timeWindow` 不叫 `window`,
  避免和 `profile.window`、Session 窗口句柄名撞概念)
- **为什么在 P2**:模块表里它排 P2 但原 backlog 漏了卡 —— 而 MqSnap 对拍的判定
  规则里有 `within {{run.timeWindow}}`(时间窗),没有这张卡 P2-04/P2-06 跑不了
- **完成**:rules.json 里 `within` + `{{run.timeWindow}}` 的规则在 fixture 单测里可判

### [ ] P2-08 mask-lite:脱敏门禁前移
- **估** 60min | **依赖** —(可与 P2-01 并行)
- **做**:只做规则版 `ebi mask check`(员工号 / 邮箱域 / UNC / `C:\Users\<id>` /
  内网 URL 的正则 + 一个词典文件),扫 `profiles/**/fixtures/` 和 tracked 文件,
  命中即失败;挂进 `Tests/Run-Tests.ps1`。交互式决策、一致性替换仍留在 P5。
- **为什么前移**:P5-03/04 原计划里掩码在 Agent 循环阶段才有,但 fixture 从
  P2-02 起就在自动积累真实页面文本 —— 等到 P5,git 历史里已经躺满了没洗过的
  内网数据,再洗要改历史
- **完成**:对一份埋了 4 类敏感项的 fixture 全部报出;`Run-Tests.ps1` 因此变红

---

# P3 — 新工作实战(8 张)

目标:**全程不写 PowerShell,只写 profile JSON + workflow JSON,搭起下一份工作的第一条流程。**
做不到就说明抽象漏了 —— 补的 step 记入 catalog,这是正常的成长。

### [ ] P3-01 素材收集(操作员做)
- **估** 操作员 60min | **读** `INTERVIEW.md` §2
- **做**:手工做一遍,每页留下 Ctrl+A 文本 + 截图 + 一句话说明;
  外加工作清单样例、**样本证据簿**、一个失败/异常的例子

### [ ] P3-02 访谈:边界 + 页面盘点
- **估** 60min | **读** `INTERVIEW.md` §3, §4
- **⚠ 必问**:主键是几列(1.3-之一);每页要做几件事(2.6)

### [ ] P3-03 访谈:判定
- **估** 45min | **读** `INTERVIEW.md` §5
- **⚠ 必问**:3.3「有没有你也看不出来的情况」—— 这一问决定容错质量

### [ ] P3-04 访谈:样本证据簿逐格
- **估** 60min | **读** `INTERVIEW.md` §6
- **做**:打开样本证据簿,**逐 sheet、逐图、逐标注**问一遍
- 这是效率最高的一段 —— 交付物结构就是流程的逆向说明书

### [ ] P3-05 访谈:变数 + 关卡
- **估** 45min | **读** `INTERVIEW.md` §7, §8
- **⚠ 必做**:§7 的 11 个问题**全部问过**,不是跳过

### [ ] P3-06 生成 profile + 调解析器
- **估** 90min | **做**:填 5 份 profile JSON;用 `ebi grammar tune` 调 grammar + 存 fixture
- **完成**:`ebi profile check` 全绿 —— **这一步在家里就能做完,不用去办公 PC**

### [ ] P3-07 生成 workflow + 复核
- **估** 60min | **完成**:`ebi lint` 绿;`ebi explain` 输出逐行念给操作员确认过

### [ ] P3-08 五级跑通
- **估** 120min(分次) | **读** `INTERVIEW.md` §9.3
- **做**:`dryrun` → `--limit 1 --guided` → `--limit 5`(**必须含一条异常样本**)→ `--limit 30` → 全量
- **⚠** 第 3 级的异常样本不能省。只在成功样本上验证过的工具,遇到第一个异常时行为完全未知

---
---

> **以下 36 张不属于「能接新工作」的最小集**,按需推进。

# P4 — Excel / 文件组(22 张)

### [ ] P4-01 excel.open + close + save — 60min — 抄 `ExcelHelpers.ps1 New-ExcelApp/Open-Workbook/Close-Workbook`
  ⚠ `$xl.Visible=$true` 要在 `DisplayAlerts=$false` 之前;COM 对象反序释放
  ⚠ (第四轮)`close` 必须能安全面对"这个名字在 Session 里不存在"——照抄
    `ExcelHelpers.ps1 Close-Workbook` 的写法:`$null` 直接 `return`,不报错。
    典型调用点是 `once:"groupEnd"`(每组一个工作簿的 compose 场景),不是
    `teardown`——`teardown` 只管 `setup` 里注册的资源(`STEP-CONTRACT.md`
    §3.4 第 5 点,`WORKFLOW-SCHEMA.md` §7.2,P0-R10)
### [ ] P4-02 excel.find_workbook — 45min — 抄 `WorkbookResolver.ps1 Find-WorkbookByExcelName`
### [ ] P4-03 excel.find_sheet + list_sheets — 45min — 抄 `ExcelHelpers.ps1 Get-SheetByName/Unhide-AllSheets`
### [ ] P4-04 excel.find_anchor — 60min — 抄 `ExcelHelpers.ps1 Get-NextAnchorRow/Get-RowAtOrBelow`
### [ ] P4-05 excel.read_cell + read_range — 45min
### [ ] P4-06 excel.write_cell + write_lines — 60min — 抄 `Write-PlainText/Write-LogLines`
### [ ] P4-07 excel.clear_below — 45min — 抄 `Reset-SheetBelowRow`
### [ ] P4-08 excel.insert_picture — 75min — 抄 `Insert-PictureSendToBack/BringToFront`
### [ ] P4-09 excel.export_pictures — 75min — 抄 `EvidenceImageExport.ps1`(⚠ 会清空剪贴板)
### [ ] P4-10 excel.draw_rect + remove_shapes — 75min — 抄 `Add-RedRectangle` + `Set-ShapeMetadata` + `Remove-MarkShapes`
### [ ] P4-11 excel.set_format — 90min — 抄 `ProcessTime.ps1` 格式化块
  ⚠ 每个关注点一个 try/catch,**不许一个裸 `catch {}` 吞掉后面所有步骤**(v2.15.1 的教训)
### [ ] P4-12 excel.add_hyperlink — 45min
### [ ] P4-13 file.wait_for_download — 75min — 监视目录直到新文件出现**且大小稳定**
### [ ] P4-14 file.rename + move + copy — 60min
### [ ] P4-15 file.newest — 45min — 抄 `JenkinsDownload.ps1 Sort-JenkinsFilesNewestFirst`
  ⚠ 多候选走 P1-27 的歧义流程,不静默选最新
### [ ] P4-16 file.unzip — 60min — 抄 `DfSnap.ps1 Expand-DfZip`
  ⚠ **保留 entry 原名**放进 per-key 子目录(用 key 命名会把时间戳后缀当扩展名,文件打不开)
### [ ] P4-17 file.backup — 45min — 抄 `BackupJ4.ps1`
### [ ] P4-18 file.stat + hash — 45min — 用于共享文件的外部改动检测
### [ ] P4-19 browser.download_link + browser.verify_action — 90min
  ⚠ `verify_action` 原设想是「包装器」,但契约里 **step 不能调 step** ——
  实现前先在 STEP-CONTRACT 补一段,二选一:做成 ui step 的可选 `verifyChange`
  输入(fill/submit 自带前后对比),或做成 runner 的 flow 构造。不要发明第三种
### [ ] P4-20 layout.json + workflows/*.compose.json — 90min — 读 `spec/PROFILE-SCHEMA.md` §7
  ⚠ (第四轮)每个交付物一个工作簿 = `once:"group"` 开 + `once:"groupEnd"` 关,
    两者在 `each` 段内配对;**别写成 `teardown` 里关**——`teardown` 只跑一次,
    只关得掉最后一组,前面的组全部泄漏(`WORKFLOW-SCHEMA.md` §7.2,P0-R10)
### [ ] P4-21 workflows/*.annotate.json — 75min
  ⚠ 红框位置会随记录条数上下移动(`baseRow`/`rowHeight`),旧工具在这框错过行
### [ ] P4-22 办公 PC 冒烟 compose + annotate — 120min [整块]

# P5 — 掩码 + Agent 循环 + 校准(14 张)

### [ ] P5-01 [整块] kernel/Mask.ps1 一致性替换 — 120min — 读 `Plan.md` §7.2
  ⚠ 同一原文永远映射到同一占位符(`<HOST_1>`),否则 Agent 看 trace 推不出关联
### [ ] P5-02 掩码规则库 + 词典 — 75min — 员工号/邮箱域/UNC/`C:\Users\<id>`/内网 URL/人名/公司名
### [ ] P5-03 ebi mask scan(交互式) — 90min — 决策存 `.ebi/redaction.json`(**gitignore**)
### [ ] P5-04 ebi mask check(CI 门禁) — 60min — **在 P2-08 的规则版上扩展**(词典、白名单、误报处理),不是新建
### [ ] P5-05 step 级 trace 落盘 — 75min — 输入/输出/耗时/产物路径/页面文本哈希/判定详情
### [ ] P5-06 ebi trace(ASCII 时间线) — 75min
### [ ] P5-07 ebi bundle — 75min — 打包 trace + 页面文本 + 截图采样 + workflow/profile,**自动套掩码**
### [ ] P5-08 ebi apply — 90min — 备份 + lint + explain 三道闸
### [ ] P5-09 transport: git 模式 — 60min — 读 `Plan.md` §8.4(当前环境是这一种)
### [ ] P5-10 docs/ebi-dance/AGENTS.md — 90min — **Agent 只能改 workflow/profile,不能改 step 代码**
### [ ] P5-11 docs/ebi-dance/COOKBOOK.md — 75min — 10 个可直接抄的组合配方
### [ ] P5-12 ebi calibrate ocr — 90min — 读 `Plan.md` §6.2;输出逐字符混淆矩阵
### [ ] P5-13 verify.ocr_read + screen.ocr + screen.preprocess — 90min — 抄 `OcrWindows.ps1`,标 `tier: fallback`
### [ ] P5-14 fallback tier 门禁接线 — 60min — 读 `spec/STEP-CONTRACT.md` §5.2
  未校准 / 过期 / 机器不符 / 准确率不达标 → **拒绝运行**

---

# 未排期(想到了记在这,不承诺做)

- `screen.stitch_scroll` 滚动拼接长图(纵/横)
- `verify.pixel` + `screen.render_html` mock HTML vs snap 像素比对
- `browser.tab_to_labeled` 用焦点元素文本定位,替代盲数 Tab
- `browser.scroll` / 列表翻页
- `table.derive` 从外部表生成工作清单
- `excel.replace_sheet` / `excel.probe_format`
- HTML 报告面板
- `ebi run --guided` 的引导文案打磨
- 清完历史老快照后删除 `legacy/`
