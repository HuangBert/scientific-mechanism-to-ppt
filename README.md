# 科研机制图生成与 PPT 原生转绘

Codex skill：`scientific-mechanism-to-ppt`。

将科研机制图或流程图从图像模型视觉稿转绘为可编辑的 PowerPoint 原生文字、形状、连线和路径，也支持直接转绘已有参考图。

## 使用

安装后可在 Codex 中输入：

```text
使用 $scientific-mechanism-to-ppt，根据我的手稿生成机制图，
再按照生成图的布局转绘为可编辑 PPT，交付 PPTX、PDF 和高清 PNG。
```

已有图片时：

```text
使用 $scientific-mechanism-to-ppt，把这张科研流程图按参考布局转绘为原生可编辑 PPT。
保留原有文字、箭头方向和颜色语义；英文使用 Times New Roman，小标题简短低调。
```

完整工作流、适用范围和验证要求见 [SKILL.md](SKILL.md)。

## 安装

在已连接此仓库的 Codex 中请求：

```text
使用 skill-installer 安装 HuangBert/scientific-mechanism-to-ppt 仓库根目录的 skill。
```

也可将仓库中的 `SKILL.md`、`agents/`、`references/` 和 `scripts/` 放入个人 skills 目录下的 `scientific-mechanism-to-ppt` 文件夹。私有仓库需要具备访问权限。

## 流程

1. 根据原始材料核对节点、公式、作用方向及必须保留的模块。
2. 需要时使用可用的图像模型生成并修正视觉稿。
3. 按参考布局规划对象坐标，重建原生文字、图形、连线、矩阵和示意曲线。
4. 导出 PPTX、PDF 和逐图 PNG，检查原生对象、真实文本下标和字体替换。
5. 查看最终渲染图，检查箭头关系、重叠、裁切及实际使用尺寸下的可读性。

“一比一”指布局、元素与科学关系对应；原生字体、渐变和手绘图标不保证逐像素复制 AI 位图纹理。整图贴进 PPT 不能作为原生可编辑重绘交付。

## 环境

- 优先使用环境提供的演示文稿编辑工具。
- 附带后备渲染脚本需要 Windows 和已安装的 Microsoft PowerPoint。
- 计划及 PPTX 检查使用 Python 标准库；PDF 检查和单图 PDF 导出可使用 PyMuPDF。
- 脚本不自动安装依赖。完整公式排版不在简单 `_{...}` 文本下标渲染器的支持范围内。

## 内容

- `SKILL.md`：使用条件与主要工作流。
- `references/image-generation.md`：生成提示词与科学核对。
- `references/native-rebuild.md`：坐标、字体、公式、导出与范围控制。
- `references/drawing-plan.md`：绘制计划格式与示例。
- `scripts/render_native.ps1`：后台 PowerPoint 原生转绘和导出。
- `scripts/validate_native.py`：计划、PPTX 和可选 PDF 检查。
- `agents/openai.yaml`：Codex 的 skill 展示信息。
