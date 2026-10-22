---
name: scientific-mechanism-to-ppt
description: "Generate scientific mechanism or workflow diagrams with an image model, then reconstruct them as editable PowerPoint shapes, text and connectors while matching the reference layout. Use for 论文机制图生成后转绘PPT、科研流程图一比一还原、AI图片转可编辑PPT. Also supports reconstructing an existing scientific reference image. Excludes ordinary presentation decks, data plotting and Visio-native editing."
---

# 科研机制图生成与 PPT 原生转绘

把科研机制图的视觉设计与可编辑交付连起来：科学内容核对 → 图像模型生成视觉稿 → 按参考布局原生转绘 → PowerPoint 渲染 → 核查与导出。

已有参考图时直接转绘。用户只要求生成图片时交付图片，不强加 PPT；用户只要求 PPT 时不额外生成付费视觉稿。以当次明确要求为准。

## 先确定内容与保留范围

- 从手稿、代码或用户说明列出节点、输入输出、连接方向、正负作用、公式与分区。记录文字来源，不能靠图像模型补充科学结论。
- 优先沿用已给出的图号、命名、画布比例和布局；明确哪些原图或模块必须保留。不要把本项目的 Figure 1 排除规则推广到其他任务。
- 科学定义或因果方向确有歧义时，只询问必要信息；其余布局与实现选择自主完成。
- 另建任务输出目录，保留原文件。未要求修改论文时，只交付图件；获准插入时沿用现有短图注和指定尺寸。

## 视觉稿

需要新图时，读取 [references/image-generation.md](references/image-generation.md)。使用当前可用的图像生成工具及其实际参数；有用户指定模型或供应商时服从该选择。不要凭记忆编造工具、端点或凭据。

先画清信息层级和连线，再丰富示意图标与层次。涉及原图编辑时先查看原图，通过图像工具编辑；不能仅凭文件名猜内容。生成结果是视觉参考，公式、文字及箭头仍要对照来源检查。局部缺陷优先局部纠正，避免每次重做整图引入新差错。

## 原生转绘

读取 [references/native-rebuild.md](references/native-rebuild.md)，根据参考图创建对象清单和坐标计划。严格还原模式以现有参考的分区、相对位置、文字、颜色语义及箭头关系为准；美化模式才调整布局，且保持科学含义。

默认样式沿用本次作者偏好，可被用户或期刊模板覆盖：

- 英文正文、标签与可兼容的公式使用 Times New Roman。
- 小标题名称简短、小字号、非粗体、低调颜色；已有图注时不重复放置醒目大标题。
- 白底、清晰线条、有限的颜色角色，原有 Peak/Trough 等语义色保持一致。

所有文字、框体、箭头、矩阵与示意曲线尽量用 PowerPoint 原生对象重建。不能把整图贴进 PPT 后声称可编辑。公式采用真实文本上下标或明确支持的可编辑数学对象；不得把不受支持的 LaTeX 源码当成公式交付。

“一比一”指版式、元素内容和逻辑关系的对应。原生字体、渐变与手描图标通常无法逐像素复制 AI 位图纹理，交付时如实说明。若个别复杂素材经允许保留为图片，列出其位置与不可编辑范围。

## 工具与脚本

优先使用环境提供的原生演示文稿工具及其实际支持的编辑、导出功能。Windows 已安装 PowerPoint 时，可用以下自包含后备实现；不要为此自动安装 Office 或依赖。

1. 按 [references/drawing-plan.md](references/drawing-plan.md) 生成 `drawing_plan.json`。
2. `python scripts/validate_native.py --plan TASK/drawing_plan.json` 检查计划。
3. `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render_native.ps1 -PlanPath TASK/drawing_plan.json -OutputDir TASK/output -BaseName mechanisms` 在后台创建 PPTX、PDF、逐页 PNG。
4. `python scripts/validate_native.py --plan TASK/drawing_plan.json --pptx TASK/output/mechanisms.pptx --pdf TASK/output/mechanisms.pdf --report TASK/native_checks.json` 检查真实产物。需要单图 PDF 时加 `--split-dir TASK/output`；已有分图时改用新目录，或在允许覆盖时加 `--overwrite-splits`。

路径带空格时按实际 shell 正确引用。脚本要求输出名无同名文件；改用新版本目录，或确认当前任务允许覆盖后显式传 `-Overwrite`。COM 登录会话错误按运行环境的权限机制请求访问，不能声称渲染成功。只关闭本次创建的文稿，不关闭用户其他演示文稿。

## 验证与交付

检查内容完整性、箭头端点、公式符号和正负号，再检查对齐、间距、裁切、字体与实际论文尺寸下的可读性。图内示意曲线不可包装成实验数据；真实数据图必须来自原始数据或已确认的原图。

验证至少覆盖：

- PPTX 可正常打开；对象名称与计划对应；原生对象数量、真实文字和上下标一致。
- 严格原生模式无图片对象；例外必须明确记录，不能仅凭文件扩展名声称可编辑。
- 导出 PDF 的字体是否发生替换、图形是否被栅格化。检查器缺少 PDF 解析依赖时标记未检查，不宣称通过。
- 查看每张最终 PowerPoint 渲染图；密集小字、公式、曲线及输入输出连线需局部放大检查。程序检查通过不能替代视觉检查。

完成实质修正后重新验证相应图件；检查通过后停止无依据的重复生成。工具导出持续失败时保留可编辑源文件与明确失败记录，不交付未经验证的预览冒充最终图。

交付可编辑 PPTX、PDF、逐图 PNG、绘制计划及验证报告。按任务需要提供单图矢量 PDF 和来源记录。说明严格还原范围、保留的图片例外和未验证项。修改手稿时，保留原图注，重新编译并查看实际页面；不自行增加正文内容。
