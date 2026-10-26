# tools/review — 交付前自查(J4 证迹簿)

交给加瀬さん之前,把当天要交的 `J4検証資料(...)_<job>.xlsx` 过一遍。
**只读**,不改工作簿。规则来自 Check Sheet_J4 的历史指摘和 10/23 的自查,
见 `docs/ebi-dance/KNOWLEDGE.md` 第 5 节(K-36 ～ K-46)。

不是 VerifyTool 的 phase,也不是 ebi-dance 的 step:是 Python 3 的旁路工具,
在云端会话(或任何装了 Python 的机器)上跑。

## 准备

```
pip install openpyxl pillow
```

## 用法

```
python3 tools/review/check_book.py <book.xlsx> [<book.xlsx> ...]
python3 tools/review/drawings.py list   <book.xlsx>
python3 tools/review/drawings.py render <book.xlsx> <输出文件夹>
```

- `check_book.py`:自动检查,每本一段,`NG` = 交之前要改,`WARN` = 看一眼。
  有 NG 时退出码 1。
- `drawings.py list`:列出每张表的图片和图形(锚定的单元格、像素位置、红框颜色)。
- `drawings.py render`:每张有图的表画成一张 PNG(图按锚定位置贴好,红框画在上面),
  用来**人眼**确认红框落在哪一行。单元格里的文字不画。

## 自动检查的项目(check_book.py)

| 项目 | 对应 |
|---|---|
| 保存状态:第一张表、各表滚到顶、光标 A3 | Check Sheet「セル選択をA3にした上で保存」 |
| 字体全是 ＭＳ ゴシック | Check Sheet「文字フォントが変わってしまってます」 |
| 日志标签粗体;标签 = 实际日志名(`GIFTReceiveUnzip.log` 等) | K-36 |
| 日志段:以 START 开头,START / END 成对、不同名,END + 日期之后不混进别的运行 | K-37 |
| GFIX受信ログ:多份 job log 的顺序跟着 DF 的组(空文件 = 最小的上传) | K-44 |
| GFIX受信ログ:Renaming 一行涂黄 | K-45 |
| 没有文字的行不留黄色;黄色涂到文字结尾、不过长 | K-38 |
| GIFTデータvsGFIXデータ:每组的相关ID 在 B、STEP 名在 G,和送信データ一一对应 | K-39 |
| 红框落在某张图上(不是孤零零浮着) | Check Sheet「最背面に赤枠の残骸」 |
| 乱码(`�`、`縺繧` 一类) | Check Sheet「文字化けしてます」 |

## 还要人眼看的(render 出来的图)

- GIFT 一览的红框行 ↔ Detail 的 RECVFILENAME ↔ 日志的 START/END 是同一个文件。
- Jenkins 红框行的文件名 / 时刻 / 大小 ↔ 日志时刻 ↔ DF 标题里的文件名。
- DF 红框行号 = 送信データ红框行号 − 1;空文件的 DF 只框「同一内容」,框要压在字上(K-40)。
- 比对结果是「相違ファイル」的,有没有吹き出し说明(K-41)。

## 不算错的(别报)

- B3 快照的ジョブ列是**正式环境**的作业名(`RJODJCP2`、`KJODJWBV`,W → J)。
- B3 快照的相関ID是 `JJPCRS13`(10/21 带着它过了审)。
- 送信データ B3 是空的 JOD 证迹簿,GIFTデータvsGFIXデータ 的 G3 也空(K-27)。

## 为什么按单元格锚定算位置

Excel 按 `<xdr:from>/<xdr:to>`(oneCellAnchor 是 `<xdr:from>` + `<xdr:ext>`)摆图;
图形自己的 `<a:xfrm>` 在插入行、改行高之后常常是旧值。10/23 用 xfrm 算,
两次把对的红框看成了错的(K-42)。
