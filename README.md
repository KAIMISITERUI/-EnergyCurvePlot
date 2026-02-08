# EnergyCurvePlot

## 简介

EnergyCurvePlot 是一个用于生成化学反应能量变化曲线的可视化工具。该程序允许用户可视化整个反应过程中的能量演变，提供对反应机理和能量分布的洞察。输出为 CDXML 文件，可在 ChemDraw 中直接打开和编辑。

## 主要功能

### v1.3 新增功能

- 保存和读取配置选项，可将页面参数直接保存，避免反复修改
- 读取和保存数据支持 xlsx 和 csv 文件格式
- 保存 cdxml 文件可自定义保存位置
- 读取 ChemDraw 文件风格功能，可使用"Load CDXML Style"选择常用风格
- 增加显示数字和显示文字的选项
- 增加取色器功能，选择颜色单元格后可直接提取颜色信息
- 为每行参数增加虚实线选项，可单独设置虚线

### v1.2 核心功能

- 修复 ID 相同导致图案层级问题
- 优化页面大小和纵横比自由调整
- 表格区域右键菜单，支持删除行/列，添加行/列
- 显示网格选项，可设置是否显示网格线
- 打开表格按钮以显示表格编辑区域
- 多个标记位于同一位置时只绘制一个
- 底部工具栏优化，所有工具正常工作

### v1.1 核心功能

- 自动扩展页面，允许无缝添加内容
- 自定义标签位置
- 绘制单个标记以进行视觉强调
- 字体大小自定义
- 字体与标签中心位置距离设置
- 线条粗细自定义
- 键粗细自定义

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行方式

```bash
python EnergyCurvePlot_v1.3.py
```

或使用指定的 Python 解释器：

```bash
"C:/ProgramData/miniforge3/python.exe" EnergyCurvePlot_v1.3.py
```

## 使用教程

详细使用教程请访问: http://bbs.keinsci.com/forum.php?mod=viewthread&tid=50113&page=1#pid316941

## 兼容性说明

所有版本保存的 table_data.json 文件都是通用的，可被任何版本读取。

## 输出格式

- CDXML 文件，可在 ChemDraw 中打开和编辑
- 支持导出为 Excel (xlsx) 和 CSV 格式

## 开发环境

- Python 3.x
- Windows/Linux/macOS
