# 内置版式库（10 类）

本文件供 Agent 在起草 brief 时选择版式。所有版式统一 16:9、空白版式 + 文本框/形状绘制，不依赖母版，跨机可编辑。字段定义以 `scripts/generate_ppt.py` 中的 `REQUIRED` 为准。

| type | 用途 | 必填字段 | 可选字段 |
|---|---|---|---|
| `cover` | 封面 | `title` | `subtitle, presenter, date` |
| `agenda` | 目录/预告 | `items[]` | `heading` |
| `section_header` | 章节分隔页 | `heading` | `subtitle` |
| `points` | 要点/方法/结论 | `heading, points[]` | `emphasis` |
| `figure` | 图 + 说明 | `heading, image` | `caption, takeaways[], metrics[]` |
| `metrics` | 指标卡片 | `heading, metrics[]{label,value}` | `metrics[].note` |
| `table` | 表格 | `heading, columns[], rows[][]` | — |
| `timeline` | 时间线/路线 | `heading, steps[]{time,title}` | `steps[].detail` |
| `conclusion` | 总结 | `heading, points[]` | — |
| `next_steps` | 下一步计划 | `heading, items[]{task}` | `items[].target, items[].deadline` |

## 布局要点

1. **cover**：中部标题（36pt 加粗）+ 副标题（18pt）+ 底部「汇报人 · 日期」；底部主题色横线。
2. **agenda**：标题 + 编号条目，编号为主题色圆角小方块。
3. **section_header**：居中大标题（32pt 加粗）+ 可选副标题，下方短横线。
4. **points**：标题（24pt）+ 项目符号列表（16pt，≤6 条）；`emphasis` 渲染为浅底 + 主题色左边线的强调框。
5. **figure**：标题 + 左侧约 58% 宽图片（等比缩放居中）+ 图注；右侧列放 takeaways 与关键指标。
6. **metrics**：标题 + 卡片网格（1/2/3 列）：大数值（28pt）+ 标签 + 说明。
7. **table**：标题 + 表格；表头主题色底白字，正文行交替浅底。
8. **timeline**：标题 + 横向轴线与节点；时间（加粗）+ 标题 + 说明。
9. **conclusion**：标题 + 编号结论（16pt）。
10. **next_steps**：标题 + 任务卡片：任务（加粗）+ 目标/截止（11pt）。

## 内容 → 版式推荐规则

对每块内容依次判断，命中即用：

1. 含图片/曲线/示意图 → `figure`
2. 含成组数字指标/精度对比 → `metrics`（行数多需对照 → `table`）
3. 含时间节点/路线图 → `timeline`
4. 纯文字要点/方法步骤/结论 → `points`
5. 章节切换 → `section_header`；开场 → `cover`；≥3 个章节 → 加 `agenda`；收尾 → `conclusion`；后续工作 → `next_steps`

排版约束：一页一个点；`points` 每页建议 ≤ 6 条；`figure` 每页一张主图；每页标题建议 ≤ 20 字。目录、时间线、下一步计划最多 12 项，超出时拆页。长文字可能在文本框内溢出，生成后请打开 PPT 检查。

