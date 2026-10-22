# 绘制计划格式

顶层为非空 slide 数组，各页画布相同。单位是参考图上的逻辑单位；渲染时乘以 `Scale` 得到 point。PPTX 最大边长不能超过 PowerPoint 支持的尺寸。

每页字段：`figure`（唯一正整数）、`name`、`width`、`height`、`objects`、可选 `notes`。`objects` 按绘制顺序排列，每个对象有唯一 `id`。

| kind | 必需字段 | 可选字段 |
| --- | --- | --- |
| shape | shape=rect/roundRect/ellipse, x,y,w,h | fill,stroke,sw,round,grad,dash |
| text | text,x,y,w,h,size,color | font,bold,italic,align,angle |
| line | x1,y1,x2,y2,color,sw | arrow,startArrow,dash |
| path | points=[[x,y],...],color,sw | fill,grad,arrow,dash |

颜色为不带 `#` 的六位 RGB 字符串；无填充或无线框时省略相应字段。`font` 默认 Times New Roman，`align` 默认 left。`_{...}` 是附带渲染器支持的简单下标标记；不支持其余 LaTeX 命令或嵌套下标。需要分段配色时拆分成具有明确颜色的原生文字对象。

最小示例，供语法核对和脚本试运行使用，不能替代逐图规划：

```json
[
  {
    "figure": 1, "name": "Mechanism", "width": 1200, "height": 600,
    "notes": "Conceptual schematic; not measured data.",
    "objects": [
      {"id":"box","kind":"shape","shape":"roundRect","x":100,"y":200,"w":300,"h":140,"fill":"EAF2F8","stroke":"536D82","sw":2,"round":12},
      {"id":"label","kind":"text","text":"Input state","x":120,"y":245,"w":260,"h":40,"size":28,"color":"1C3043","align":"center"},
      {"id":"arrow","kind":"line","x1":410,"y1":270,"x2":700,"y2":270,"color":"536D82","sw":2,"arrow":true}
    ]
  }
]
```
