本人想实现word转ppt功能，但ppt软件自带的从大纲导入功能效果较差，我通过常见的免费AI生成Markdown大纲，然后把大纲转换成为可编辑的pptx文件。经过测试已经成功。
《大纲转 PPT 工具 v4.1 官方使用说明书》

（最专业、最好用、完全免费的 Markdown → PPT 一键生成神器）

一、软件概述
软件名称：大纲转 PPT 工具（Outline to PPT）
当前版本：v4.1（2025年最新正式版）
开发技术：Python + PyQt6 + python-pptx + lxml
适用人群：老师、学生、职场人、培训讲师、产品经理、运营、咨询顾问……所有需要快速做出漂亮PPT的人
核心优势：

完全支持 Markdown 语法（写 Markdown 就是写 PPT）(可使用deepseek等免费AI)
一键生成专业级 PPT，标题自动加粗、首行缩进 2 字符、1.5 倍行距、段前段后 0pt
中英文/数字字体完全分离（中文微软雅黑，英文数字自动 Times New Roman）
支持自定义任意 PPT 模板（公司品牌模板、年度汇报模板随便用）
自动生成封面页 + 目录页
6 套专业配色方案 + 深色模式 + 记住所有设置
支持拖拽 .md/.txt 文件导入
导出后可自动打开 PPT
二、安装与启动
Bash

# 1. 安装依赖（只需一次）
pip install PyQt6 python-pptx lxml

# 2. 保存下面完整的代码为：md文件转pptx（测试成功版）.py
# 3. 双击运行 或 命令行：
python "md文件转pptx（测试成功版）.py"
打包成免安装绿色 exe（推荐给同事/领导）：

Bash

pyinstaller --onefile --windowed --name="大纲转PPT v4.1" --icon=icon.ico "md文件转pptx（测试成功版）.py"
打包后 dist 文件夹里就是一个 exe，双击即用，无需任何环境。

三、界面总览与功能示例
<img width="475" height="716" alt="image" src="https://github.com/user-attachments/assets/f041ddb2-e00a-4ff4-880b-6ab327efd2f5" />
<img width="1901" height="1005" alt="image" src="https://github.com/user-attachments/assets/88cc5737-4c78-48be-a305-1d43be7d958e" />
<img width="1500" height="773" alt="image" src="https://github.com/user-attachments/assets/08c5f520-7e86-4b88-aaed-75cf092cdac0" />


Q&A 环节
只需 3 步：

把AI生成的大纲内容粘贴到左侧大纲区
点击右下角蓝色大按钮 → 选择保存路径
自动弹出专业级 PPT，领导看了直呼内行！
五、高级功能详解
选择公司模板（重磅！）
点击「选择」按钮 → 选中你公司的标准 PPT 模板（.pptx）
以后所有生成的 PPT 都会自动套用公司配色、Logo、页脚、字体规范

六、目前表格显示还是有点问题，需要编辑优化。
