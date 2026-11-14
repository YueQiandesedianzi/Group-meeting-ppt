# M0 技术验证：阶段结论（初步）

- 日期：2026-11-14
- 分支：`v2-m0`（基线 `9e1a53d`），v1 路径未改动
- 样例：`v2/samples/reference-template.pptx`、`v2/samples/hand-edited.pptx`
- 客户端：**LibreOffice Kit**（本机未安装 PowerPoint 或 WPS）
- 证据：`v2/reports/*.json`（结构、往返、部件保留）

> 全部结论都基于合成样例与 LibreOffice 代理，**不能等同于 PowerPoint 验收**。凡标注
> 「代理盲区」的条目，必须在装有 PowerPoint 的目标机上复核后才能进入 M1。

## 1. 结论摘要

| # | 验证项 | 结果 | 证据 |
|---|---|---|---|
| 1 | 真实模板加载（母版/Logo/页脚/非默认尺寸） | 通过 | 生成结果 13.333×8.333，母版含 Logo 图片、页脚、强调线；渲染 Fig 区域均值前后一致 |
| 2 | 指定对象最小修改、其余内容保留 | 通过（我方引擎） | 定向编辑后 50 个部件中 43 个字节不变，仅 6 个被重写、4 个新增 |
| 3 | 自定义身份标记（备用文本 `descr`） | 通过 | 客户端保存后幻灯片对象的 `gmppt:*` 标记全部保留 |
| 4 | SlideID → 部件映射稳定 | 通过 | 往返前后 `id→part` 一一对应（256→slide1 … 262→slide7） |
| 5 | 页序调整 | **未验证（代理盲区）** | 文件内 `sldIdLst` 顺序正确，但 LibreOffice 一律按物理部件顺序渲染与写回 |
| 6 | 复制页面产生的身份冲突 | 检测到 | 复制页与源页携带同一 UUID，需 M1 的重复检测与重新分配 |
| 7 | 自由元素（箭头/说明框/新增自由页） | 通过 | 往返后对象与标记均保留 |
| 8 | 母版形状的结构保留 | **不保留（结构变化）** | 客户保存把母版形状下放到全部 11 个布局，母版 `spTree` 变空；视觉一致 |
| 9 | 幻灯片尺寸 | **被归一化** | 12191695×7619695 → 12192000×7620000；`sldSz/@type` 丢失 |
| 10 | 形状 id / 类型 | 不稳定 | id 2–6 → 94–99；`TEXT_BOX` 变 `AUTO_SHAPE`；名称与标记保留 |
| 11 | 表格对象 | 通过 | 往返后表格与内容保留 |
| 12 | 复杂度兜底 | 未覆盖 | 图表 / SmartArt / OLE / 媒体 / 公式样例尚未纳入 |
| 13 | PowerPoint Tags 往返 | **代理上丢失** | LibreOffice 保存后 @B@ppt/tags/*@B@ 全部消失，9/9 对象的 Tags 归零；我们自己的读写保存 9/9 保留 |
| 14 | 对象级修改（文本 / 图片 / 几何 / 页序） | 通过（原型） | 未指向对象 0 变化；50 个部件中 45 个字节不变，仅 4 个重写、1 个新增 |

## 2. 关键测量

### 2.1 部件级保留（retention_check.py）

```text
参考模板 → 我方定向编辑：   identical 43 / rewritten 6 / added 4 / removed 0
参考模板 → LibreOffice 保存：identical  3 / rewritten 42 / added 1 / removed 4
```

- 定向编辑只重写被改的 3 张幻灯片，加上 `[Content_Types].xml`、`presentation.xml` 与
  一个 rels，**布局、母版、主题、媒体全部字节不变**（43/50）。
- 客户端保存几乎重写整个包，仅媒体 3 个部件字节不变。这直接印证规格 §10.3：
  不能以整包哈希判断保留性，也不能以“客户端保存成功”当作无损证据。

### 2.2 母版形状的下放

```text
保存前  母版 spTree: nvGrpSpPr, grpSpPr, sp×5, pic(Logo), sp, sp(页脚)
保存后  母版 spTree: nvGrpSpPr, grpSpPr
        每个布局 spTree: 各自多出 pic(Logo) + sp(页脚)，标记随形状一起搬走（11/11 布局）
```

Logo 与页脚区域在往返前后的渲染均值完全一致（`(109,135,163)` / `(253,253,253)`），
即**视觉保留、结构不保留**。对同步引擎的直接影响：

1. 不能以“母版 XML 未变”作为固定元素保留的证据；
2. 反向读取时必须同时扫描母版与布局，才能找到 Logo/页脚这类固定元素；
3. 布局数量会随固定元素复制而膨胀，比较时要按语义而非结构位置。

### 2.3 页序（代理盲区）

```text
文件内 sldIdLst（原始 XML，我方重排后）:
  256→slide1(封面), 260→slide5(小结), 257→slide2(正文), 258→slide3(图文), 259→slide4(双图), 261→复制页, 262→自由页
LibreOffice 保存后:
  256→slide1, 257→slide2, 258→slide3, 259→slide4, 260→slide5, 261, 262
```

`id→part` 映射没变，但 `sldIdLst` 的顺序被改回物理部件顺序；渲染结果同样按部件顺序，
逐页哈希对比显示“第 N 页 ↔ 第 N 页”。也就是说，**LibreOffice 既不按 `sldIdLst` 渲染，也不按它写回**，
无法用来证明或证伪“用户在 PowerPoint 里重排页面”。规格 §13 T06/T08 必须用 PowerPoint 复核。

### 2.4 身份标记与冲突

- 标记写在对象备用文本（`descr`）中，形如 `gmppt:picture:result-chart=<uuid>`，客户端保存后全部保留。
- 复制页保留源页 UUID（符合 PowerPoint 会复制自定义数据的预期），`diff` 能直接列出重复标记；
  M1 需要“重复 UUID → 分配新身份 / 无法区分则停止写入”的分支。
- 形状 ID 与名称不可作为主身份：ID 被重排，名称会随复制重名。

### 2.5 PowerPoint Tags 往返（规格 §6.1 建议的正式身份载体）

试验方式：给 9 个带身份标记的对象挂上 `gmppt.project / gmppt.slide / gmppt.object` 三组 Tags，
Tags 存放在独立部件 `ppt/tags/tag01.xml … tag05.xml`，由对象 `p:nvPr/p:custDataLst/p:tags` 以 `r:id` 引用。

```text
我方读写保存（python-pptx）：tags 部件 5 个保留，9/9 对象标签完全一致
LibreOffice 保存往返：        tags 部件 0 个，9/9 对象标签全部丢失
```

结论：**Tags 不能经 LibreOffice 这类客户端保存**。这意味着

1. 任何以 Tags 为唯一身份载体的设计，在链路中只要有一次非 PowerPoint 的保存就会失去全部身份；
2. 备用文本（`descr`）在两种路径下都保留，应作为与 Tags 并行的冗余载体；
3. `doctor` 必须能报告“当前客户端是否保留 Tags”，否则同步会在无提示的情况下退化为盲写。
4. PowerPoint 对 Tags 的真实行为仍未验证（本机无 PowerPoint），本结论只说明代理行为。

### 2.6 对象级修改原型（M1 起步）

`v2/engine/object_ops.py` 实现了四类按身份定位的操作：`set_text`（保留首 run 格式、多段落报能力不足）、
`replace_picture`（建立新媒体引用，不覆盖共享媒体）、`set_geometry`（可按比例约束）、
`reorder_slides`（按 slide_id 序列重排，参数不匹配即拒绝）。身份重复时直接拒绝写入。

`v2/tools/object_ops_demo.py` 对参考模板连续执行：改封面标题 → 换结果图 → 调整图片位置尺寸 → 把小节页移到第 2 位，
然后做两项检查：

```text
未指向对象：changed 0 / missing 0 / unexpected_added 0      → 通过
部件级保留：identical 45 / rewritten 4 / added 1（新增媒体）  → 只有目标页与包级清单被重写
```

被重写的 4 个部件是：两张目标幻灯片、`[Content_Types].xml`、`presentation.xml`（页序）；
母版、11 个布局、主题、既有媒体全部字节不变。这验证了规格 §3.2「普通更新采用对象级修改」在当前技术栈上可行。

### 2.7 复制页身份冲突（失败案例，已处理）

复制页面会连同身份一起复制。@B@hand-edited.pptx@B@ 里复制封面页后：

@B@@B@@B@text
cover-title 身份出现次数：2    唯一身份数：1
写入尝试 → CapabilityError: 身份 'gmppt:page-title:cover-title=...' 命中 2 个对象（疑似复制页导致重复身份），停止写入
@B@@B@@B@

规格 §6.1 要求“能可靠确认原件时为副本分配新身份，无法区分时停止相关写入”。
当前实现选择**停止写入**并给出命中数量；副本与新身份的分配留到 M1（需要用户确认或语义指纹）。
这条同时在自动化测试中断言（@B@IdentityConflictTests@B@）。

## 3. 失败案例与能力边界

| 案例 | 现象 | 处理 | 复现 |
|---|---|---|---|
| 复制页导致身份重复 | 同一 UUID 命中 2 个对象 | 拒绝写入并报告命中数 | @B@python -m unittest v2.tests.test_m0.IdentityConflictTests@B@ |
| 客户端保存丢弃 Tags | @B@ppt/tags/*@B@ 与引用全部消失 | 身份改为 Tags + 备用文本双载体；doctor 待实现 | @B@python v2/tools/tags_probe.py@B@ |
| 客户端保存重写整包 | 仅媒体 3 个部件字节不变 | 保留性检查必须按部件与语义进行，禁止整包哈希判断 | @B@python v2/tools/retention_check.py v2/samples/reference-template.pptx v2/samples/reference-template.roundtrip.pptx@B@ |
| 母版形状被下放至布局 | 母版 spTree 清空，11 个布局各多出 Logo 与页脚 | 反向读取需同时扫描母版与布局；不能以母版未变作为保留证据 | 见 §2.2 |
| 页序语义在代理上不可验证 | @B@sldIdLst@B@ 顺序被规范化 | 页序操作先挂起，等 PowerPoint 复核 | 见 §2.3 |
| 幻灯片尺寸被归一化 | 12191695×7619695 → 12192000×7620000 | 尺寸比较使用容差；记录 @B@sldSz/@type@B@ 丢失 | @B@python v2/tools/pptx_inspect.py <file>@B@ |
| 形状 ID / 类型不稳定 | id 2–6 → 94–99，@B@TEXT_BOX@B@ → @B@AUTO_SHAPE@B@ | 身份只用备用文本与 Tags，不用 id/名称 | 见 §2.4 |
| 多段落文本改写 | 能力不足 | 返回明确原因，不静默降级 | @B@v2/engine/object_ops.py@B@ 的 @B@set_text@B@ |
| 复杂对象（图表/SmartArt/OLE/媒体/公式） | 未覆盖 | 待补样例后逐类给出保留或拒绝结论 | 规划中 |

## 4. 自动化测试

@B@@B@@B@powershell
python -m unittest discover -s v2/tests -v            # 快速：11 项通过，2 项跳过
$env:GM_PPT_SLOW=1; python -m unittest discover -s v2/tests -v   # 含客户端往返：13 项全部通过
@B@@B@@B@

覆盖：模板结构与非默认尺寸、母版 Logo/页脚标记、每页身份标记、复制页身份冲突与拒绝写入、
定向修改后未指向对象 0 变化、部件保留阈值、页序参数不匹配拒绝、图片替换新增媒体部件、
Tags 在我方保存后 9/9 一致，以及（慢速）客户端保存重写整包与丢弃 Tags 的代理行为。

## 5. 对规格的反馈

1. §6.1 中“SlideID 在同一文稿内重排后保持不变”在 LibreOffice 上不成立（顺序被规范化），
   且我方合成样例的 ID 是连续分配的，尚不能证明 PowerPoint 行为；M1 前需在目标机复核。
2. §6.2 要求“原始保留信息”必须包含布局层：固定元素可能被搬到布局，只记母版会漏。
3. §10.3 的“未触碰部件字节哈希不变”在本工具链上**可以实现**（我方编辑 43/50 不变），
   但跨客户端保存时不可用，需要语义级比对（本次已提供部件级对比作为下限）。
4. 新增建议：把“渲染器/客户端的页序语义”写进 `doctor` 能力检查，否则 T06/T08 会产生假通过。

## 6. 工具清单

| 工具 | 作用 |
|---|---|
| @B@v2/tools/build_samples.py@B@ | 生成参考模板与模拟手改样例 |
| @B@v2/tools/pptx_inspect.py@B@ | 只读检查页面/对象/身份标记/Tags，支持两文件 diff |
| @B@v2/tools/roundtrip_check.py@B@ | 客户端往返并比较身份与结构保留 |
| @B@v2/tools/retention_check.py@B@ | 按 ZIP 部件比较字节哈希 |
| @B@v2/tools/tags_probe.py@B@ | PowerPoint Tags 创建 / 读回 / 往返存活试验 |
| @B@v2/tools/object_ops_demo.py@B@ | 对象级修改原型验证 |
| @B@v2/engine/object_ops.py@B@ | 对象级修改适配层（文本 / 图片 / 几何 / 页序 + 身份冲突保护） |
| @B@v2/tests/test_m0.py@B@ | M0 回归测试 |

## 7. 已知限制（必须随交付一并说明）

- 本机无 PowerPoint / WPS，无法做原生客户端往返与视觉验收。
- 样例为合成模板，不等同于真实模板（规格 §14 要求的真实输入仍未取得）。
- 渲染用 LibreOffice，与 PowerPoint 存在字体、字距与抗锯齿差异；文中所有“视觉一致”仅指同一渲染器下的对比。
- 未覆盖动画、超链接、多媒体、公式等语义保留。

## 8. 下一步建议

1. **PowerPoint 复核清单**：在装有 PowerPoint 的目标机上重跑 §2.3 的对照（导出页面图 + `pptx_inspect` 报告），
   确认页序是否按 `sldIdLst` 生效、Tags 是否保留、`SlideID` 是否稳定。
2. **复杂度支持矩阵**：补齐图表、SmartArt、OLE、媒体、公式样例，逐类给出保留或拒绝结论。
3. **CLI 与 doctor**：把 `inspect` / `doctor` 收敛为规格 §9.1 的稳定 JSON 协议，并声明客户端能力
   （无 PowerPoint 时页序不可验证、Tags 可能被丢弃）。
4. **M1 接入**：对象级操作已可用，先接文本与图片两条链路；页序操作等复核后再启用。

## 9. M0 完成度（对照规格 §15）

| 规格 §15 事项 | 状态 | 证据 |
|---|---|---|
| 1. 可复现模板与手工修改样例 | 完成（合成） | `v2/samples/reference-template.pptx`、`hand-edited.pptx`；真实模板待用户提供 |
| 2. 项目/页面/对象身份与只读检查命令 | 完成 | `pptx_inspect.py`（页面/对象/标记/Tags/diff）、`gmppt:*` 标记与 Tags 双载体 |
| 3. 真实模板加载、指定对象修改与不相关内容保留 | 完成 | §2.6：未指向对象 0 变化，50 个部件中 45 个字节不变 |
| 4. PowerPoint 保存/重排/复制后的身份行为 | 部分（代理） | 保存与复制已用 LibreOffice 验证；页序为代理盲区，需 PowerPoint 复核（§2.3） |
| 5. 支持矩阵与失败案例 | 完成 | §1 支持矩阵 + §3 失败案例（均附复现命令） |

结论：M0 在“可用客户端 + 明确标注”的范围内完成。唯一未闭环项是 **PowerPoint 原生客户端复核**，
原因是本机未安装 PowerPoint/WPS；该限制已在 §3 与 §7 显式标注，其余结论均由可复现实验与自动化测试支撑。
