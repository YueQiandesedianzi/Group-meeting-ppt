# v2 / M0 技术验证（进行中）

本目录是 v2 规格（`Group-meeting-ppt_v2_spec.md`）M0 阶段的开发分支产物，与 v1 使用路径完全隔离：
v1 的 `scripts/`、`SKILL.md` 未做任何修改。

## 目录

```text
v2/
├─ tools/
│  ├─ build_samples.py    生成参考模板与“手工修改”对照样例
│  ├─ pptx_inspect.py     只读检查：页面/对象/身份候选，可两文件 diff
│  ├─ roundtrip_check.py  用可用客户端重新保存，比较身份与结构保留
│  └─ retention_check.py  按 ZIP 部件比较字节哈希（保留性证据）
├─ samples/               生成的样例（PPTX 与映射 JSON）
└─ reports/               机器可读的验证报告
```

## 复现

```powershell
python v2\tools\build_samples.py
python v2\tools\pptx_inspect.py v2\samples\reference-template.pptx -o v2\reports\inspect-reference.json
python v2\tools\roundtrip_check.py --input v2\samples\reference-template.pptx --report v2\reports\roundtrip-reference.json
python v2\tools\retention_check.py v2\samples\reference-template.pptx v2\samples\reference-template.roundtrip.pptx
```

`roundtrip_check.py` 默认用 DSH 自带的 LibreOffice Kit 作为客户端代理；可用 `--node` /
`--cli` 指定其它渲染/转换器。它不支持 PowerPoint 自动化（当前机器未安装 PowerPoint）。

## 样例内容

`reference-template.pptx`：13.333×8.333 英寸（非默认尺寸），母版含 Logo 图片、页脚文本框与
强调线，5 个示例页（封面 / 正文 / 图文 / 双图 / 总结），第 2 页另有页级固定色带，
若干对象写入了 `gmppt:*` 身份标记（备用文本属性 `descr`）。

`hand-edited.pptx`：模拟用户在 PowerPoint 中的常见改动——改封面标题、移动并缩放第 3 页图片、
改图注、在第 4 页加箭头与说明框、把第 5 页移到第 2 位、复制封面页、末尾新增自由页。

`template-map.json`：页面用途 / 内容区域 / 固定区域 / 自由区域的最小映射草案（规格 §5.3）。

## 当前结论

见 reports/M0-preliminary.md。要点：身份标记与 SlideID→部件映射在客户端保存后仍可用；
母版形状会被 LibreOffice 下放到全部布局（视觉保留、结构变化）；客户端保存会重写几乎所有部件，
因此保留性必须按部件与语义检查；页序验证在当前代理客户端上存在盲区，需 PowerPoint 复核。
