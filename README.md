《大纲转 PPT 工具 v4.1 官方使用说明书》

背景：本人想实现word转ppt功能，但ppt软件自带的从大纲导入功能效果较差，我通过常见的免费AI生成Markdown大纲，然后把大纲转换成为可编辑的pptx文件。经过测试已经成功。

（目前是最专业、最好用、完全免费的 Markdown → PPT 一键生成神器）

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
<img width="963" height="866" alt="image" src="https://github.com/user-attachments/assets/7fd3bebc-9e04-4062-ab5d-2da85a982cae" />
<img width="1914" height="1012" alt="image" src="https://github.com/user-attachments/assets/7e4f4b2f-7325-4d41-bc0d-f5cff44a1b66" />
<img width="1876" height="810" alt="image" src="https://github.com/user-attachments/assets/013d0ac7-dcf4-406e-960a-6882791a86b0" />
<img width="1900" height="965" alt="image" src="https://github.com/user-attachments/assets/9eb8ad67-0bca-4df8-9203-0645817799f0" />






四、操作步骤：

把AI生成的大纲内容粘贴到左侧大纲区
点击右下角蓝色大按钮 → 选择保存路径
自动弹出专业级 PPT，领导看了直呼内行！

五、高级功能详解
选择公司模板（重磅！）
点击「选择」按钮 → 选中你公司的标准 PPT 模板（.pptx）
以后所有生成的 PPT 都会自动套用公司配色、Logo、页脚、字体规范

六、目前表格显示还是有点问题，需要手动编辑优化。

七、大纲生成提示词：
请作为资深 PPT 策划专家，将我提供的文本转换为 Markdown 格式的 PPT 文本大纲。
格式要求（严格遵守，不得偏差）：

深度清理（优先级最高）：
必须彻底过滤并删除所有类似 、`[cite_end]`、（XX为数字）以及任何形式的 `` 标记。
最终输出中严禁出现任何方括号包含的引用信息，只保留最纯净的文本。

标题分级规范：
使用 # 表示 PPT 总标题。（单独一页）
使用 ##  表示 PPT 目录。（单独一页）
使用 ### 表示每一页幻灯片的标题。（尽量整齐对仗，按 第1. 2. 3... 章递增）
使用 #### 表示页面内的子标题，尽量整齐对仗，使用 1.2，1.3... 等文章写作的标准的格式序号进行排序。
使用 #### # 表示页面内的子标题下面的正文，注意换行和使用（1）、（2）、等排序符号来排序。分号句号等标点符号不能不能省略。

物理分页与结构：
每一页幻灯片之间必须使用三个短横线 --- 进行物理分隔。

视觉平衡：每页字数要均匀，逻辑过长时请自动拆分为多页，每一页的字数不能超过100字（这是强制要求）。尽可能的美观。

表格转化：
若涉及数据对比、分类或流程，请自动将其转化为标准的 Markdown 文字对比表达，不要用使用表格。

输出限制：
直接输出 Markdown 源码块，严禁进行 UI 渲染。
确保我可以无损地复制为 .txt 文件，严禁删除任何要求的 # 符号。
