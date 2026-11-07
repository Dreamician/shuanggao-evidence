# shuanggao-evidence

按学校统一母版，把建设任务与凭证原件组装成可交付的佐证材料 docx，并在交付前批量复核。

## 背景

高职院校"双高计划"按年度下达建设任务，每项任务验收要交一册佐证材料。学校给一份空白 Word 母版，每册固定由「抬头与年度、任务标题、100–150 字描述、佐证材料清单、逐份材料图片」组成；正文不抄录原件文字，佐证一律以图片呈现。

评审看的是"这份凭证能否证明这条任务完成"。因此既要结果件（证书、通报、立项名单、结题证书），也要过程件（工作方案、筹备分工、任务书、日程、抽签、会议纪要、课程标准）和效果件（验收、技术报告、应用证明）。只放一叠获奖证书，或只放一份孤立方案，都不完整。

任务由合并总表分发，佐证负责人常与领办人不同；凭证散落在 PDF、DOC、OFD、XLS、截图里，一个任务少则一两张、多则几十张图。同一批人要同时管多册、上百张图，且每册格式必须一致。

## 问题

1. **找不齐**：任务与凭证分属不同目录、不同格式，容易漏用可用附件，或把汇总台账误当凭证。
2. **判不准**：结果凭证只认出具单位原始件，过程件按责任主体、时间、正式版式认定，两者口径不同；红头/公章、文件首页、年份、数量词、描述承诺，错一处就可能被退回。
3. **排不好**：母版带合并单元格和提示文字，手工改容易多空块、图片错位、标题继承成蓝色；8cm 图宽、同行等高、清单表行高 482 twips、表格行 `cantSplit` 等参数靠肉眼核不动。
4. **复核慢**：多册上百张图，逐份检查字体、页边距、图宽、清单序号、占位残留；改一处要重看全部。
5. **易回退**：在 WPS、预览或 editor_sdk 里保存一次，`cantSplit` 就丢；为改一句话重建整册，之前的增量修改会全部回退。

## 方法

**规则与实现分离**：`SKILL.md` 写流程与铁律，`references/evidence-rules.md` 写凭证认定与材料构成判据，`references/format-spec.md` 写母版 OOXML 参数。判定先于排版——结果凭证只认出具单位原始件；正式形成、有责任主体的过程文件必须入册；自说自话的汇总 Excel、周报、新闻稿不入册。

**一份 spec 驱动一册**：`spec.json` 声明专业群、年度、任务标题、描述、佐证材料清单，以及每个材料对应的图片路径。字段见 `references/spec-example.json`。

**四步流水线**（均在技能根目录执行）：

1. 定任务：读合并总表，按「佐证负责人」列筛出本人任务，用「专业群 + 任务编号」联合定位。
2. 做图：`imgprep.py` 把原件渲染成 8cm 凭证图，支持裁边、红框、表格线检测、多页公文 A4 画框；`.doc` / `.ofd` / `.xls` 的转换链见 `SKILL.md` 第 2 步。
3. 成册：`build_doc.py` 复制母版、改抬头/描述/清单/材料标题、按图片数重塑材料表、倒序插图、属性收尾。
4. 复核：`audit.py` 批量检查，`--content` 追加内容口径提示，`fix_doc.py` 修复编辑器保存丢失的属性。

**增量改动直接改 OOXML**：只改文字或单张图时，替换对应 `w:t`，或替换 `word/media` 图片并同步 `wp:extent`，不要重建整册。

## 效果

- 任务、清单、材料标题、图片路径由同一份 spec 对应，改哪一册只动哪一册，不牵连其它册。
- `audit.py` 一条命令复核整个目录，覆盖 17 项结构检查：抬头与年度、描述字数、清单序号、材料标题字号、图片数与宽度、显示宽高比（变形）、同行图片等高、`cantSplit`、页边距、占位残留等；`--content` 另给 6 类内容口径提示。
- 格式参数固化在脚本与规则文档里，不靠记忆；编辑器保存丢属性后可一键修复。
- 材料构成按「结果 + 过程 + 效果」成套判定，避免整册只有证书、或方案孤零零没有执行记录。
- 复核可重复执行，改动后重跑一次即可确认没有回退。

## 用法

先备好空白母版 docx 与凭证原件。

```bash
cp references/spec-example.json spec.json        # 改 group/year/task/desc/list/materials；template 指向母版
python3 scripts/imgprep.py 原件.pdf --out ./img --tag m1   # 多页公文加 --frame；每材料换 tag 重跑
python3 scripts/build_doc.py spec.json           # 需 editor_sdk
python3 scripts/fix_doc.py 目标目录 --margins    # 仅在编辑器保存过之后需要
python3 scripts/audit.py 目标目录
python3 scripts/audit.py 目标目录 --content
```

图片路径须与 `materials[].images` 一致；改文字或单张图时直接改 OOXML，不必重建整份。

## 目录

```
SKILL.md                     规则、流程、踩坑
references/
  evidence-rules.md          凭证认定与材料构成判据
  format-spec.md             母版 OOXML 参数
  spec-example.json          spec 字段示例
scripts/
  build_doc.py               按 spec 从母版生成整份文档
  imgprep.py                 原件 → 8cm 凭证图（裁边 / 红框 / A4 画框）
  audit.py                   批量复核；--content 追加内容口径提示
  fix_doc.py                 属性收尾（黑字 / 16pt / cantSplit / 居中；页边距需 --margins）
  sdk.py                     editor_sdk 封装
```

## 关键口径

- 结果件、过程件、效果件成套：制度配执行记录，方案配过程或效果材料
- 内部台账（汇总 Excel、周报、新闻稿）不入册；正式形成、有责任主体的过程文件必须入册
- 红头或公章二选一；竞赛类附件须有文件首页 + 公章；与本院无关的整页可省略
- 图片一律 8cm 宽、同行等高、居中；佐证材料一律以图片形式呈现
- 中文仿宋、西文 Times New Roman；描述 100–150 字；清单行数 = 材料数

## 依赖

运行依赖：Python 3 + Pillow、ghostscript（`gs`，渲染 PDF）、editor_sdk（build_doc.py 插图用）。
旧格式手工转换另需：`.doc` → textutil + 无头 Edge/Chrome + PyMuPDF；`.ofd` → WPS 另存 PDF；`.xls/.xlsx` → 见 SKILL.md 第 2 步。

## 注意

- 在编辑器里保存会丢 `cantSplit`，属性修复须放在最后一次保存之后（见 SKILL.md）
- 替换成品文件先移入废纸篓再复制，不要直接覆盖
