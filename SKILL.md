---
name: pptx-editable-figures
description: 生成「全可编辑」的 PPTX 插图（流程图、体系图、甘特图、架构图等）。用原生 PowerPoint 形状绘制而非贴图，用户可在 PowerPoint/WPS 中直接改文字、拖位置、换配色。当用户要求"做成可编辑的 PPT/幻灯片图""画流程图放到 PPTX""把 Word 里的图改成可编辑"时使用。内置浮点坐标防御、几何自检与字符画预览校验，杜绝箭头漂移。
agent_created: true
---

# 可编辑 PPTX 插图生成

把静态图（PNG/Word 插图/文字描述）转成**原生形状**的 PPTX，用户可自行编辑。
产出物是可用的 PPT 文件，不是图片。

## 一、核心陷阱（血泪教训，务必遵守）

### ⚠️ 1. 浮点 EMU 坐标 —— 致命，必翻车
OOXML 的 EMU 坐标（`a:off` / `a:ext`）**必须是整数**。Python 的 `/` 产生 float，
而 python-pptx **不做取整直接写入**，生成 `x="2414811.5"`。

控制变量实验（已实测）：
| 文件 | PowerPoint 打开 |
|---|---|
| 整数坐标 | ✅ 成功 |
| **浮点坐标** | ❌ **失败** |
| 同值取整回退 | ✅ 成功 |

WPS/PowerPoint 容错打开时不报错，而是让线条位置**回退默认值** → 表现就是**箭头飘到别处**。
这就是"画飘了"的根因，与坐标算得对不对无关。

**对策**：所有坐标经 `E()` / `inch()` 取整。禁止裸写 `(a-b)/2`、`x*0.5` 后直接传参。
本 skill 的 `pptxfig.py` 已在底层统一取整，用它的 API 即可免疫。

### 2. 中文字体
必须同时设 latin 与 eastAsia typeface，否则中文回落宋体：
`rPr.find(qn('a:ea')).set('typeface', '微软雅黑')`

### 3. a:ln 子元素顺序
严格为：fill → prstDash → headEnd → tailEnd。
箭头端点必须在设置完颜色/虚线**之后** append，否则 PowerPoint 报错。

### 4. 颜色用 LineFormat，不是 XML 节点
`conn.line.color.rgb = ...` ✅
`conn.line._get_or_add_ln().color` ❌（不存在）
只有 append headEnd/tailEnd 时才用 `_get_or_add_ln()`。

### 5. 虚线枚举位置
`from pptx.enum.dml import MSO_LINE_DASH_STYLE`（**不是** `pptx.enum.shapes`）。

### 6. 合并多页时的保存顺序
先用各页 `Fig.save()` 存单页文件，**再**合并。反之会让单页版本也变成多页。

## 二、标准工作流

```
1. 选模板（templates/figures.py）→ 改数据
2. 生成 PPTX
3. verify()         几何自检：浮点/超界/箭头端点贴合
4. check_render()   字符画预览 + 彩色元素位置报告   ← 必做，AI 靠它"看图"
5. 交付
```

**第 4 步不可省略**。几何自检只能验证数值关系，字符画才能发现视觉错位。

## 三、怎么用

### 用模板（推荐，改数据即可）
```python
import sys
sys.path.insert(0, r"C:\Users\yangc\.workbuddy\skills\pptx-editable-figures\templates")
from figures import make_flowchart, make_system_diagram, make_gantt, merge_figures

f1 = make_flowchart(
    [("信息采集", "学业数据、行为数据\n建立学生学业档案"),
     ("分级预警", "黄/橙/红三级\n教务初筛、辅导员复核"),
     ("预警谈话", "48小时内约谈\n共定改进计划")],
    title="图2-1  学业预警与帮扶工作流程图",
    loop_from=2, loop_to=1,                       # 闭环回流：末框→第2框
    loop_label="未改善：升级干预措施",
    note="闭环管理：预警不解除不销号")

f2 = make_system_diagram(
    [("物质帮助", "奖助勤贷免补", "#D6E6F2"),
     ("能力拓展", "勤工助学、技能培训", "#BBD6EA"),
     ("精神激励", "励志教育", "#D6E6F2")],
    "解困—育人—成才—回馈",
    [("价值引领", "志愿服务", "#BBD6EA"),
     ("精准认定", "定量+定性", "#D6E6F2"),
     ("隐私保护", "去标识化公示", "#BBD6EA")],
    title="图8-1  发展型资助育人体系构成",
    loop_pair=(0, 0), loop_label="回馈反哺")     # 下层0号 → 上层0号

f3 = make_gantt(
    [("一  大纲论证", 0, 3, "#2E5C8A"),
     ("二  分篇撰写", 3, 7, "#3E7CB1")],
    {3: "初稿完成", 10: "试用启动"},
    title="图3-1  进度甘特图（2026.09—2028.06）")

merge_figures([f1, f2, f3], "out.pptx")
```

### 从零画（图型不在模板里时）
```python
from pptxfig import Fig, verify
f = Fig(13.333, 7.5)                      # 16:9
f.box(1.0, 2.0, 2.15, 1.9, title="节点", desc="说明", fill="#E8F1F8", line="#2E75B6")
f.arrow_between(rec_a, rec_b, "h", color="#C0392B", width=2.5)   # 框到框，自动贴合
f.elbows([(x1,y1),(x2,y2),(x3,y3)], dash=MSO_LINE_DASH_STYLE.DASH, tail=True)  # 折线
f.save("out.pptx")
```

### API
| 方法 | 说明 |
|---|---|
| `.text/.box/.bar` | 文本框 / 圆角矩形 / 甘特条 |
| `.arrow(x1,y1,x2,y2,...)` | 直线箭头 |
| `.arrow_between(a,b,side)` | **框间箭头，自动贴合边缘**（side: h / v / v-up）|
| `.elbows(points,...)` | 折线箭头（回流、闭环）|
| `verify(path)` | 几何自检，返回问题列表 |
| `export_png(path,dir)` | PowerPoint COM 渲染 |
| `check_render.report(path)` | 字符画 + 元素位置报告 |

## 四、字符画预览（关键能力）

AI 看不到图片，`check_render.py` 把渲染结果转成文本：

```
python scripts/check_render.py <pptx路径>

图例: R=红(箭头) D=深蓝 B=中蓝 o=浅蓝框 *=文字 -=浅灰
|..DooooooooooooD.DooooooooooooD.DooooooooooooD...|   ← 三个框
|..DoooDDDDDDoooRRRoooDDDDDDoooRRRoooDDDDDDoooD..|   ← 箭头在框之间 ✓
```

**采样要点**：不要在缩放图上平均（细线会被抹掉），要对每个字符格取
"最偏离白色"的像素。色板顺序敏感，白色必须先判，否则背景被误判为浅蓝。

判读方法：箭头字符（`R`）应出现在**两个框之间的间隙**；若出现在标题区、
画布边缘或框内部，即漂移。

## 五、环境
Python 3.9+，依赖见 `requirements.txt`（`python-pptx` / `pywin32` / `pillow`）
COM 注意：同一进程内多次 Quit/Dispatch 会失败，需在**同一会话**内打开多个文件比对。

## 六、实战案例
申报书三图（流程图/体系图/甘特图）：
`D:\FDownload\商姚的申请书\辅导员管理-辅导员培训\output\fdygl-jc-20260913\build_pptx_figures.py`
同目录 `verify_pixels.py` 为定制像素校验示例。
