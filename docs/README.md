# docs/ 索引

按**你是谁、要干什么**找文档,不用全读。

---

## 我要继续 ebi-dance 的重构开发

**新会话从这里开始,按顺序读三样,就够开工了:**

1. `ebi-dance/BACKLOG.md` — **挑一张卡**。每张卡自带上下文,告诉你要读哪一节、抄哪个函数、什么算完成
2. 卡里指定的那 1-2 节规格(**不要读完整份**)
3. `ebi-dance/BACKLOG.md` 顶部的「全局铁律」7 条 —— 这个每次都要看

一次做一张卡,一张卡一个 commit。

**当前共 103 张卡**(到 P3 可接新工作的最小集是 67 张)。开工前有一版契约
审查追加的 **`P0-R1`…`P0-R9` 修订卡**(在 BACKLOG.md 最前面,`P0-R` 前缀,
全部是纯文档的规格修订,不写代码)—— 它们**先于任何写码卡**,已经全部
完成;后续会话直接从 `P0-01` 开始写码即可。

---

## ebi-dance 文档地图

| 文件 | 读者 | 什么时候读 |
|------|------|-----------|
| `ebi-dance/Plan.md` | 人 | 想知道**为什么**这么设计、整体长什么样 |
| `ebi-dance/BACKLOG.md` | 人 + Agent | **每次开工**。任务卡 + 铁律 |
| `ebi-dance/INTERVIEW.md` | **Agent** | 要给一份新工作做流程访谈时。**不依赖代码,现在就能用** |
| `ebi-dance/spec/VOCABULARY.md` | 人 + Agent | 起名字的时候;搞不清 role / action / verb 区别的时候 |
| `ebi-dance/spec/STEP-CONTRACT.md` | 写 step 的人 | 写任何一个 `modules/**` 文件之前 |
| `ebi-dance/spec/WORKFLOW-SCHEMA.md` | 写工作流的人 + Agent | 写 `workflows/*.json` 之前 |
| `ebi-dance/spec/PROFILE-SCHEMA.md` | 写 profile 的人 + Agent | 接一份新工作、或改判定规则的时候 |
| `ebi-dance/CATALOG.md` | 人 | 有哪些 step 可用(**自动生成**,别手改) |
| `ebi-dance/catalog.json` | **Agent** | 组装工作流的唯一数据源(**自动生成**) |

### 读哪一份,看你要干什么

| 你要… | 读 |
|-------|-----|
| 加一个新的 step | `spec/STEP-CONTRACT.md` §1-3, §9 |
| 搭一条新工作流 | `spec/WORKFLOW-SCHEMA.md` §1-4, §8 |
| 接一份全新的工作 | `INTERVIEW.md` 全文 + `spec/PROFILE-SCHEMA.md` §10 |
| 改判定规则 | `spec/PROFILE-SCHEMA.md` §5 |
| 页面解析不对 | `spec/PROFILE-SCHEMA.md` §4 |
| 主键匹配不上 / 匹配到多个 | `spec/PROFILE-SCHEMA.md` §6 |
| 起名字拿不准 | `spec/VOCABULARY.md` |
| 搞懂 OCR 为什么被降级 | `Plan.md` §6 |
| 知道敏感信息怎么处理 | `Plan.md` §7 |

---

## 我在用现在的工具干活(旧版 VerifyTool)

重构期间旧工具**继续可用**,冻结在 `freeze/pre-ebi-dance` 标签。

| 文件 | 内容 |
|------|------|
| `Operations.md` | 日常操作手册 —— 各 phase 怎么跑 |
| `Configuration.md` | 配置分层规则(psd1 / 工作目录 overlay / session) |
| `SendVsGift.md` | SendVsGift phase 的详细说明 |
| `SnapVerify-Plan.md` | SnapVerify 判定的设计文档 |
| `Versioning.md` | 版本号规则 |

---

## 历史与决策记录

| 文件 | 内容 | 状态 |
|------|------|------|
| `Generalization-Roadmap.md` | 早期的分层重构路线图(M0–M6) | **已被 `ebi-dance/Plan.md` 取代**。保留是因为它的沙盒化审计和分支策略仍然有效 |
| `Sanitization-Audit.md` | 敏感信息审计报告 | 有效 —— `ebi-dance` 的掩码模块设计基于它 |
| `Parked-Ideas.md` | 设计过但主动搁置的东西 | 有效。**不在待办列表上** |
| `ProcessTime-OldSnap-MockMatch-Plan.md` | D2 图像检查方案 | 已搁置(见 `Parked-Ideas.md`) |
| `ProcessTime-OldSnap-Verify-Plan.md` | 老快照人工验证方案 | 已实现(v2.17.0) |
| `ProcessTime-OcrBenchmark-Plan.md` | OCR 基准测试方案 | 未做。`ebi-dance` 用 `ebi calibrate ocr` 取代了它的大部分意图 |

---

## 一句话版本

> 旧工具按**业务步骤**切文件,同一个能力抄了好几遍(裁剪抄了 4 份,交互写了 77 处)。
> ebi-dance 按**能力**切,用 JSON 描述工作流。
> 换一份工作只改 profile,不改代码。
