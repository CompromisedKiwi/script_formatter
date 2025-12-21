# 剧本格式化工具
# 介绍
- 这是一个用于格式化剧本文档的工具
- 支持的文件格式：`.docx, .doc`

# feature
- **标红并加粗**对话（对话段落中冒号后的内容，中英逗号均可）
    - 其他段落保持不变
    - 忽略所有非对话段落

# 使用方法
1. 启动
   1. 使用发行版: 直接运行 `dialogue_to_red.exe` 文件
   2. 从源代码运行: `python dialogue_to_red.py`
2. 在弹出的文件选择对话框中，选择要格式化的剧本文档
3. 程序将在同一目录下生成一个新文件，文件名前会添加 `标红_` 前缀
4. 格式化完成后，会弹出提示框显示保存路径

# 构建
1. 确保已安装 Python 3.x
2. 安装必要的库：
   ```
   pip install python-docx tkinter pyinstaller
   ```
3. `pyinstaller -D -w --version-file=version.txt --name "剧本格式化工具" .\dialogue_to_red.py`
