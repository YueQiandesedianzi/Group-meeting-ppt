---
name: group-meeting-ppt
description: 生成可编辑的科研组会汇报 PPT（.pptx）。按 brief.yaml 规范整理内容并渲染为 16:9 幻灯片，支持导入/配置模板（Tkinter 向导）、单页局部修改（--slide）与轻量历史快照（--history-dir）。当用户需要组会/进展/开题/综述汇报 PPT 时使用；仅统一现有 PPT 的字体配色等视觉样式（不涉及内容编排）时不用。
---

# 组会 PPT 自动化

把一次组会汇报的内容素材确定性转成可编辑 `.pptx`。内容与结构以 `brief.yaml` 为唯一事实源，脚本负责校验与渲染；模板决定配色/字体/页面尺寸。

## 何时使用

- 用户要「组会汇报 / 进展汇报 / 开题 / 综述 slides」等。
- 已有文字要点、图表、指标，需要排成 PPT。

**不用**的场景：只改现有 PPT 的视觉样式（不涉及内容编排）；学术海报、期刊排版。

## 文件

- `scripts/generate_ppt.py`：生成/校验/单页重建/快照。
- `scripts/template_wizard.py`：Tkinter 图形界面向导，导入/选择/配置模板。
- `scripts/workbench.py`：统一工作台（推荐 GUI），集成 brief 编辑 / 生成 / 历史预览 / 模板管理。
- `templates/*.yaml`：模板档案（页面尺寸 + 配色 + 字体），内置 `default.yaml`。
- `references/layout-library.md`：10 类版式字段与推荐规则（起草 brief 时查阅）。
- `assets/demo_brief.yaml`：完整示例 brief，可作为模板。

## 工作流

推荐用 `python scripts/workbench.py` 打开图形工作台；以下为脚本化 / Agent 化流程。

1. **首次使用**：若 `templates/` 下没有用户模板且用户想用自己的风格，运行
   `python scripts/template_wizard.py`（无图形环境时对话式询问后直接写 `templates/<name>.yaml`）。
2. **收集内容**：确认汇报类型（进展/开题/综述）与使用的模板名。
3. **起草 brief**：按下方规范写 `brief.yaml`，向用户确认要点。
4. **生成**：
   `python scripts/generate_ppt.py <brief.yaml> -o <out.pptx> [--template <name>] [--history-dir <dir>]`
5. **回读检查**：检查分页、文本、图片占位，提出局部修改建议。

## 编辑与保存模型（避免反复全量重建 / 手改被覆盖）

- `brief.yaml` 是内容/结构的**唯一事实源**，`.pptx` 是派生产物；改动内容一律改 brief，不从 pptx 反向解析。
- **默认同名覆盖**：重跑固定写 `-o` 路径，不自动生成 `_v2/_v3`。
- **单页重建**：只改某一节时用 `--slide N`（N 即 brief 第 N 个 section），原地重绘该页、保留其余页（含 PPT 里的视觉微调）。
- **输入保护**：输出不能覆盖 brief、所选模板或引用图片；单页重建时现有 PPT 页数必须与 brief 一致，页序仍需自行核对。
- **变更检测**：开启 `--history-dir` 时，brief、模板、图片、生成器、输出身份及实际输出文件哈希都匹配才跳过全量重建。单页重建后标记为需要全量重建，即使 brief 随后恢复原内容也不会误跳过。
- **收敛**：生成→检查→改 brief 的循环最多 2–3 轮，修改建议一次列全。

## 历史版本（轻量，不反复保存 pptx）

每次运行仍写入 `-o` 指定的 PPT；历史目录另外保存可重建的内容、模板、图片和轻量预览，不为每个版本另存 PPT：

```text
history/
  versions/<时间戳>-<hash8>/
    brief.yaml       # 图片已改为历史目录内的相对引用
    template.yaml    # 当次生成所用的完整模板
    manifest.json    # 格式版本、依赖哈希、生成方式等
  previews/<id>/preview.png  # 内容结构预览，非逐像素还原
  assets/<完整SHA-256>        # 跨版本按图片内容去重
```

- 引用图片按完整内容哈希去重，不再依赖原始路径。移动或备份历史时，必须整体保留 `versions/` 和 `assets/`。
- 用历史版本重建完整 PPT（即使该版本由单页操作产生，也按 brief 全量重建）：
  ```bash
  python scripts/generate_ppt.py <history>/versions/<id>/brief.yaml -o <restored.pptx>
  ```
- 新快照默认使用同目录的 `template.yaml`，即使另外指定了 `--template-dir` 也不会改变这一默认；只有显式传入 `--template <name>` 才覆盖历史模板。普通 brief 未指定模板时仍使用内置或指定目录下的 `default`。工作台的「从该版本重建」遵循同一规则，不会覆盖历史版本。
- 重建会核对快照 brief、归档模板和图片的 SHA-256；资源损坏或被替换时拒绝生成。显式改用其他模板时不要求读取已损坏的旧模板。
- `.last_generated.json` 记录生成状态及实际输出文件哈希。单页操作将它标记为 `partial` 和 `requires_full_rebuild`；全量输出已保存但历史尚未完成时标记为 `history_pending`，旧全量信息在 `last_full` 中保留。只有全量生成且快照、预览成功后才标记为 `full`，并核对输出文件哈希后才允许跳过。旧 `.last_generated_hash` 不含完整信息，不再用于跳过生成。
- **旧格式限制**：原有的 `versions/<时间戳>-<hash8>.yaml` 没有归档图片和模板，只能继续浏览、预览；重建仍需原始图片并明确选择当前模板，移动历史或删除源文件后不保证可重建。
- 同一个输出文件和历史目录应串行生成，不支持多个进程同时写入。

## brief 规范（摘要）

顶层 `meta`（`title` 必填）+ `sections`（有序列表）。每段一个 `type`，字段见 `references/layout-library.md`。

- 默认章节骨架：封面 → 目录 → 背景 → 方法 → 结果 → 讨论 → 结论 → 下一步。
- 逐块分类：有图 → `figure`；成组数字 → `metrics`（多行对照 → `table`）；带时间 → `timeline`；纯文字 → `points`。
- 校验失败会非零退出并给出具体错误；`figure.image` 必须存在，`table` 每行长度与列数一致。
- 目录、时间线、下一步计划每页最多 12 项，超出时拆页。长文字仍需在生成的 PPT 中检查是否溢出。

## 模板（主题级）

模板档案 YAML 示例：

```yaml
name: my-lab
description: 课题组组会模板
slide: { aspect: "16:9", width_in: 13.333, height_in: 7.5 }
colors:
  primary: "#1F3864"
  accent: "#2E74B5"
  light: "#D6E4F0"
  text: "#333333"
  muted: "#808080"
  line: "#C9C9C9"
  background: "#FFFFFF"
fonts: { cjk: "微软雅黑", latin: "Calibri" }
default: true
```

向导从 `.pptx` 的 `theme1.xml` 提取配色方案（dk1/lt1/dk2/lt2/accent1–6）与字体（含 ea 中文字体），缺省时用默认值。

## 验证

历史功能与基本版式回归测试（不读取或改写真实历史）：

```bash
python -m unittest discover -s tests -v
```

