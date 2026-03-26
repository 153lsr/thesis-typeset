# 论文排版工具测试报告（2026-03-26）

## 1. 测试目标

本轮测试围绕 `E:\本科毕业论文\任务\03_封装命令` 展开，目标是：

- 验证工具主要功能在当前 Windows 环境下可用。
- 补齐 GUI 内部交互测试。
- 记录已确认缺陷、已修复问题和剩余风险。

## 2. 测试环境

- 操作系统：Windows
- Python：`py -3.10`
- Word：已安装，可通过 COM 调用
- GUI 测试方式：Tkinter harness（拦截 `mainloop()`，对按钮、复选框和启动流程做内部交互验证）

## 3. 主要测试输入

### 3.1 用户提供的正式文件

- 论文：`C:\Users\li'si'rong\Desktop\论文.docx`
- 自定义封面：`C:\Users\li'si'rong\Desktop\封面.docx`

### 3.2 辅助测试夹具

位于：`C:\Users\li'si'rong\.codex\memories\thesis_tool_test`

用于覆盖：

- `论文.docx`
- `封面.docx`
- `sample.txt`
- `sample.md`
- `sample_valid.tex`
- `unsupported.rtf`
- `custom_config.yaml`
- 配置自动发现与外部封面专用目录用例

## 4. 结果总览

### 4.1 CLI / 启动层

| 项目 | 结果 | 说明 |
| --- | --- | --- |
| EXE `dump-config` | 通过 | 可导出默认配置 |
| 源码 CLI `dump-config` | 通过 | 与 EXE 行为一致 |
| EXE GUI 启动冒烟 | 通过 | 可启动窗口 |
| `run_gui.py` 启动冒烟 | 通过 | 可启动窗口 |

### 4.2 输入格式与输出流程

| 项目 | 结果 | 说明 |
| --- | --- | --- |
| `docx` 默认论文格式化 | 通过 | 主流程可完成 |
| `doc` 输入兼容 | 通过 | `.doc` 可转入流程 |
| `txt` 输入 | 通过 | 可生成输出文档 |
| `md` 输入 | 通过 | 已做辅助兼容验证 |
| `tex` 输入 | 通过 | 已做辅助兼容验证 |
| 自定义配置 `--config` | 通过 | 配置项被写入输出 |
| `--toc-only` | 通过 | 仅更新目录流程正常 |
| 输入目录自动发现 `thesis_config.yaml` | 通过 | 配置自动发现有效 |
| `cover.only_insert` 外部封面流程 | 通过 | 可单独插入外部封面 |
| 不支持格式报错 | 通过 | `unsupported.rtf` 被拦截 |
| 缺失输入文件报错 | 通过 | 能正确拦截 |

### 4.3 GUI 内部交互

| 项目 | 结果 | 说明 |
| --- | --- | --- |
| 输入/输出/Logo/自定义封面浏览 | 通过 | 路径写入正确 |
| 保存配置 / 加载配置 / 恢复默认 | 通过 | YAML 往返正常 |
| `skip` / `toc-only` / `cover-only` 互斥联动 | 通过 | 状态切换符合设计 |
| 封面字段 添加/删除末行 | 通过 | 按钮级事件正常 |
| 特殊标题 添加/删除末行 | 通过 | 按钮级事件正常 |
| 导航切换到“单独设置” | 通过 | 面板切换正常 |
| `start` 无效输入校验 | 通过 | 可拦截并恢复按钮状态 |
| `start` 缺少输出路径校验 | 通过 | 可拦截 |
| `start` 参数值无效校验 | 通过 | 非法整数项可触发报错 |
| `start` cover-only 未选自定义封面 | 通过 | 可拦截 |
| `start` cover-only 已选自定义封面 | 通过 | worker 启动且配置传递正确 |
| 头部文件摘要 trace 刷新 | 通过 | 输入/输出变更会自动刷新 |
| 状态文本 `_set_status()` | 通过 | 状态变量正常更新 |
| 日志队列 `_append_log()` / `_poll()` | 通过 | 日志消费与轮询续调正常 |

## 5. 本轮已修复问题

### 5.1 GUI 残留导航分支引用不存在 `_cat_list`

- 文件：`thesis_gui.py`
- 位置：`_on_cat_select()`
- 原问题：方法内部直接访问 `self._cat_list`，但当前 GUI 并没有创建该控件。
- 风险：如果该残留入口被调用，会触发 `AttributeError`。
- 处理：改为兼容式安全分支；当 `_cat_list` 不存在时直接返回，存在时再复用现有 `_on_cat_click()` 流程。
- 回归测试：`tests/test_thesis_gui_navigation.py`

### 5.2 多节正文文档只在“最后一节”生成正文页眉 / 奇偶页码

- 文件：`thesis_formatter/page.py`、`thesis_formatter/headers.py`
- 现象：当文档在正文开始后还包含多个现有分节时，工具原先把 `doc.sections[-1]` 当成唯一正文节。结果是正文前面的多个节虽然已经进入正文内容，但仍沿用前置部分的页眉/页脚策略，直到最后一节才出现正文页眉与奇偶页码位置。
- 黑盒复现证据：
  - 合成论文用例中，正文第 1 节未拿到 `ODD_HEADER / EVEN_HEADER`，只有最后一个正文节生效。
  - 用户正式论文 + 正式封面用例中，修复前第一页正文页眉直到第 75 页才首次出现。
- 根因：`setup_page_numbers()` 与 `setup_headers()` 都把“正文节”硬编码成最后一节，而不是“正文起点所在节及其后续所有节”。
- 处理：新增 `get_body_start_section_index()`，按 `find_first_body_heading()` 定位正文起点所在节，并把该节及其后续全部节统一按正文处理。
- 回归测试：`tests/test_multi_section_body_headers.py`
- 验证命令：`py -3.10 -m unittest discover -s tests -v`

### 5.3 前置页摘要 / Abstract 缺失标签时不再强行补写，且摘要不套一级标题样式

- 文件：`thesis_formatter/formatter.py`、`thesis_formatter/_titles.py`、`thesis_formatter/headings.py`
- 原问题：
  - `auto` / `format` 模式下，只要进入前置页格式化分支，就会把首个非空前置页段落强行改成“摘要”。
  - 当英文摘要是独立 `Abstract` 标题时，工具未按标题处理，后续英文摘要正文还可能被误判为标题区。
  - `摘要` 在自动识别标题阶段可能先被套上 `Heading 1`，这既不符合常见论文要求，也会干扰前置页识别与正文起点判断。
- 本轮处理：
  - 若前置页缺少明确 `摘要` / `Abstract` 标签，保留原文，不再自动补写，只追加警告。
  - 独立 `Abstract` 现在按英文摘要标题处理：居中、标题字号，但不使用 `Heading 1` 样式。
  - 中文 `摘要` 同样仅保留标题外观，不再使用 `Heading 1` 语义样式。
  - `auto` 模式下的前置页识别改为跳过前置页特殊标题，按“正文首个标题”截断，避免被 `摘要` / `目录` 等前置页标题短路。
  - 分页逻辑未改动，仍保持在识别到 `关键词` / `Key words` 后插入分页。
- 新增回归测试：`tests/test_front_matter_recognition.py`
- 回归覆盖点：
  - 识别独立 `Abstract` 为前置页。
  - `auto` 模式下将独立 `Abstract` 排成标题，但不套 `Heading 1`。
  - `auto` / `format` 模式下，缺少中文摘要标签时只告警、不补写。
  - 缺少英文 `Abstract` 标签时只告警、不注入 `Abstract`。
- 验证命令：`py -3.10 -m unittest discover -s tests -v`


### 5.4 前置部分多分节时，目录尾页被错误切换为阿拉伯页码

- 文件：`thesis_formatter/page.py`
- 现象：在用户正式论文的真实黑盒输出里，正文页眉直到第 17 页才开始出现，但修复前第 13-15 页这些仍属于目录尾部的页面已经提前显示 `PN1ZZ / PN2ZZ / PN3ZZ`，导致“页眉边界”和“页码边界”不一致。
- 根因：`setup_page_numbers()` 原先只给第 0 节设置前置页格式、只给正文起始节设置正文格式；如果正文前还存在多个前置分节，后续前置节会保留 Word 默认阿拉伯页码格式。
- 处理：
  - 将正文起始节之前的所有节统一设置为前置页码格式；
  - 将正文起始节及其后续全部节统一设置为正文页码格式；
  - 对连续节显式清理旧的 `w:start`，避免页码在中途被意外重置。
- 回归测试：`tests/test_multi_section_body_headers.py`
- 黑盒复验结果：
  - 第 12 页页码为 `PNXZZ`；
  - 第 13-15 页仍分别为 `PNXIZZ / PNXIIZZ / PNXIIIZZ`；
  - 第 17 页才首次出现 `ODDHDRTOKEN` 与 `PN1ZZ`，第 18 页出现 `EVENHDRTOKEN` 与 `PN2ZZ`。
- 验证命令：`py -3.10 -m unittest discover -s tests -v`

## 6. 已确认但未修复的问题

### 6.1 `--no-postprocess` 不是绝对生效

- 文件：`thesis_runner.py`
- 位置：约第 117-125 行
- 现象：当 `caption_mode_effective == dynamic` 时，即使请求跳过后处理，也会强制走 Word COM 后处理。
- 已验证情况：
  - `md` 场景可尊重 `--no-postprocess`
  - `txt` + dynamic 题注场景会忽略该选项
- 当前结论：这是一个真实功能不一致问题，尚未修复。

## 7. 最新 fresh 验证（使用正式论文与封面路径）

本报告写入前，已再次用 GUI harness 验证以下路径：

- 输入文件：`C:\Users\li'si'rong\Desktop\论文.docx`
- 自定义封面：`C:\Users\li'si'rong\Desktop\封面.docx`

验证结果：

- 两个路径均存在。
- `cover-only` 启动分支无报错。
- `run_format()` 收到的 `inp`、`custom_docx` 与上述正式路径一致。
- GUI 最终状态为“格式化完成”，按钮恢复为 `normal`。

### 7.1 用户式黑盒验收：前置部分 / 正文页眉页码

本轮新增了更贴近用户感知的黑盒验收：先通过 Word COM 导出 PDF，再用 `pdftotext -tsv` 抽取各页真实渲染文本与坐标，不再只看 XML 结构。

用例 1：合成论文 + 用户正式封面（`scope=body`，正文奇偶页眉，正文奇偶页码左右切换）

- 第 9 页仍只有前置部分页码 `PN7ZZ`，无页眉。
- 第 11 页首次出现 `ODDHDRTOKEN`，同页页码 `PN1ZZ` 位于右侧（`left≈512.5`）。
- 第 12 页出现 `EVENHDRTOKEN`，同页页码 `PN2ZZ` 位于左侧（`left≈68.06`）。
- 结论：封面 / 前置部分无页眉，正文开始后奇偶页眉与奇偶页码位置都按配置生效。

用例 2：`first_page_no_header=true`

- 第 2-3 页无页眉、无页码。
- 第 4 页开始出现 `EVENHDRTOKEN` 与 `PN2ZZ`，第 5 页出现 `ODDHDRTOKEN` 与 `PN3ZZ`。
- 当前结论：该选项的实际效果是“正文首页同时不显示页眉和页码”，这一点已确认，但是否符合具体学校模板仍需按院校规范判断。

### 7.2 用户正式论文黑盒复验（同正式验收配置）

输入：

- 论文：`C:\Users\li'si'rong\Desktop\论文.docx`
- 配置：`C:\Users\li'si'rong\.codex\memories\thesis_tool_test\blackbox_acceptance_front_matter_2026-03-26\config.yaml`

fresh 证据：

- `摘    要` 保留在第 7 页，英文 `Abstract` 保留在第 10 页，未被强行改写为 `Abstract:`。
- fresh CLI 日志中未再出现“前置页缺少明确的摘要 / Abstract 标题”这类误报告警。
- 第 12 页页码为前置部分罗马页码 `PNXZZ`。
- 第 13-15 页仍分别为 `PNXIZZ / PNXIIZZ / PNXIIIZZ`，说明目录尾页没有被提前切到正文阿拉伯页码。
- 第 17 页首次同时出现 `ODDHDRTOKEN` 与 `PN1ZZ`。
- 第 18 页出现 `EVENHDRTOKEN` 与 `PN2ZZ`。

结论：对用户提供的正式论文，前置部分页码与正文页眉/页码边界现在已经对齐，正文页码从实际正文起始页第 17 页开始重新计为 1；修复前“目录尾页已切入阿拉伯页码、但正文页眉尚未开始”的问题已消除。

## 8. 剩余风险

- GUI 原生文件选择框本身未做人工逐点录屏式测试，当前以内部 harness 为主。
- GUI 成功分支中的真实后台线程 + Word COM 结束回调，主要通过 stub/harness 验证参数传递与状态恢复；真实文档流程已在 CLI/主流程侧验证。
- `md` / `tex` 属于辅助兼容检查，不是本轮用户重点输入格式。

## 9. 当前结论

当前工具在 Word / TXT 主路径、配置管理、目录处理、外部封面插入、GUI 核心交互，以及“前置部分无页眉 / 正文开始后出现页眉 / 正文奇偶页码左右切换”这些用户可感知特性上，已经完成了可复核验证。

本轮新增确认并修复了 2 个真实排版缺陷：

- 多节正文文档会把正文页眉与奇偶页码错误地推迟到最后一节；
- 当前置部分本身存在多个分节时，目录尾页会被提前切到阿拉伯页码。

修复后，用户正式论文的正文页眉与正文页码都从实际正文起始页第 17 页开始正常生效。

当前仍需要继续跟进的主要问题只有 1 个：`--no-postprocess` 在 dynamic 题注模式下会被覆盖。
