# LocalAgent

> **一个强大的本地 AI 智能体工具** —— 用于操作计算机和处理工作任务,针对本地小模型优化

<div align="center">

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![Ollama](https://img.shields.io/badge/Ollama-Latest-green.svg)](https://ollama.ai/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.55.0-red.svg)](https://streamlit.io/)

</div>

---
## Demos
#### demo1_优化日程表
<video src="demo/demo1_优化日程表.mp4" controls="controls" width="500" height="300"></video>

#### demo2_处理文件
<video src="demo/demo2_处理文件.mp4" controls="controls" width="500" height="300"></video>

#### demo3_数据可视化
<video src="demo/demo3_数据可视化.mp4" controls="controls" width="500" height="300"></video>

---

## 📖 项目简介

LocalAgent 是一个基于 **Ollama 本地大模型** 和 **Streamlit** 构建的智能体应用。它能够理解用户的自然语言指令，自动调用各种工具来完成复杂的电脑操作任务，如文件处理、文档转换、代码执行、图像识别等。

### 核心特性

- 🤖 **本地 AI 驱动** - 使用 Ollama 运行本地大模型（支持 Qwen、GLM 等）
- 🛠️ **丰富工具集** - 内置 18+ 种工具，涵盖文件操作、代码执行、文档处理等
- 💻 **系统交互** - 可执行 PowerShell 命令，操作文件系统
- 📄 **Office 支持** - 完整支持 Word、Excel、PPT 的读写和转换
- 🖼️ **图像识别** - 集成 OCR 功能，可识别图片中的文字
- 🔄 **自动化工具创建** - 支持将 Python 代码封装为可复用的 CLI 工具

---

## 🚀 快速开始

### 环境要求

- Python 3.10+
- Windows 系统
- [Ollama](https://ollama.ai/) (需预先安装并配置)

### 安装步骤

1. **克隆或下载项目**

```bash
cd 项目目录
```

2. **创建并激活虚拟环境**

```bash
python -m venv .venv
.venv\Scripts\activate
```

3. **初始化外部文件**
```
.\setup.cmd
```

4. **安装依赖**

```bash
pip install -r requirements.txt
```

5. **启动应用**

方式一：使用启动脚本(可以双击运行)
```cmd
.\launch.cmd
```

方式二：手动启动
```bash
.venv\Scripts\activate && streamlit run streamlit_app.py
```

---

## 🛠️ 功能模块

### 工具列表

| 类别 | 功能 | 描述 |
|------|------|------|
| **文件操作** | `read_file` | 读取文件内容 |
| | `create_file` | 创建新文件 |
| | `replace_file_content` | 替换文件部分内容 |
| **代码执行** | `run_powershell` | 执行 PowerShell 命令 |
| | `run_python_code` | 运行 Python 代码 |
| | `create_python_tool` | 创建可复用的 Python 工具 |
| **Word 文档** | `read_word_and_export_txt` | 读取 Word 导出文本 |
| | `convert_word_or_txt_to_pdf` | Word/TXT 转 PDF |
| | `convert_markdown_to_word` | Markdown 转 Word |
| **Excel 表格** | `read_excel_and_export_txt` | 读取 Excel 导出文本 |
| | `create_excel_from_2d_list` | 从二维列表创建 Excel |
| **PPT 演示** | `read_ppt_and_export_txt` | 读取 PPT 导出文本 |
| | `replace_ppt_content` | 替换 PPT 内容 |
| | `create_ppt_from_txt` | 从 TXT 创建 PPT |
| | `convert_ppt_to_pdf` | PPT 转 PDF |
| **图像识别** | `recognize_image_and_export_markdown` | OCR 识别图片文字 |
| **其他** | `add_knowledge` | 添加知识到知识库 |
| | `wait_user_do` | 等待用户完成操作 |

### 任务示例

```
✅ 计算 160968*(23516-75061)
✅ 创建一个 HTML 贪吃蛇游戏
✅ 制作一个 PPT 来介绍你自己
✅ 把桌面上的 Word 文档转换为 PDF
✅ 整理桌面上的单词表图片按词整理性到 Word 文件
✅ 根据 data.xlsx 绘制折线图和饼图
```

---

## 📁 项目结构

```
LocalAgent/
├── streamlit_app.py      # Streamlit 主应用
├── tools.py              # 工具函数库
├── ocr.py                # OCR 图像识别模块
├── markdown_to_word.py   # Markdown 转 Word 转换器
├── run.py                # 启动脚本
├── launch.cmd            # Windows 启动命令
├── requirements.txt      # Python 依赖
├── prompt/
│   └── prompt.md         # AI 系统提示词模板
├── images/               # UI 图标资源
├── output/               # 输出文件目录
└── demo/                 # 演示文件目录
```

---

## ⚙️ 配置说明

### Ollama 模型配置

确保 Ollama 服务在 `http://localhost:11434/` 运行，并已拉取所需模型：

```bash
ollama pull glm-4.7-flash
# 19G左右,建议一张4060+32G内存,效果最好

ollama pull qwen3:8b
# 5G左右,一张4060即可流畅运行

ollama pull qwen3:4b
# 2.5G左右,建议模型的参数不要小于4b
```

### 外部文件目录

项目依赖外部文件目录 `D:/ExternalFiles/`：

- 存放生成的临时文件
- 存储用户自定义工具（.py 文件）
- 存放工具使用说明（.md 文件）
- 知识库文件（KNOWLEDGE.txt）

### 系统提示词

`prompt/prompt.md` 包含 AI 的系统提示词模板，支持变量替换：

- `$EXTERNALFILES$` - 外部文件列表
- `$KNOWLEDGE$` - 用户知识库内容

---

## 💡 使用技巧

1. **停止任务**：点击 `Stop` 按钮，然后在 `...` 菜单中选择 `Rerun`

2. **创建可复用工具**：使用 `create_python_tool` 将常用代码封装为 CLI 工具，提高效率

3. **修复代码错误**：使用 `replace_file_content` 快速修复 Python 代码的 bug

4. **添加知识**：使用 `add_knowledge` 记录解决问题的技巧、系统信息等

5. **路径规范**：所有文件路径请使用绝对路径，Windows 系统使用反斜杠 `\`

---

## 🔧 开发说明

### 添加新工具

1. 在 `tools.py` 中定义工具函数，使用标准 docstring 格式：

```python
def my_new_tool(param1, param2):
    """
    工具描述

    Args:
      param1 (str): 参数 1 说明
      param2 (int): 参数 2 说明

    Returns:
      str: 返回状态信息
    """
    # 实现代码
    return "操作成功"
```

2. 在 `streamlit_app.py` 中注册工具：

```python
st.session_state.tools.append(my_new_tool)
st.session_state.available_functions["my_new_tool"] = my_new_tool
# 注意要更新st.session_state.tools和st.session_state.available_functions
```

### 依赖说明

主要依赖库：

| 库 | 用途 |
|----|------|
| `streamlit` | Web UI 框架 |
| `ollama` | 本地大模型客户端 |
| `python-docx` | Word 文档处理 |
| `openpyxl` | Excel 文档处理 |
| `python-pptx` | PPT 文档处理 |
| `pywin32` | Windows COM 接口（PDF 转换） |
| `paddleocr` | OCR 文字识别 |

---

## ⚠️ 注意事项

1. **本地模型**: 请选择支持工具调用的模型比如`qwen3`系列模型,建议设置16k以上的上下文长度

2. **OCR功能**: OCR功能需要额外安装[PaddleOCR-VL](https://www.paddleocr.ai/main/version3.x/pipeline_usage/PaddleOCR-VL.html)
   
3. **安全性**：工具可执行任意 PowerShell 命令和 Python 代码，请在可信环境中使用

4. **输出限制**：部分工具输出有长度限制（如 2500/3000 字符）

5. **编码格式**：文件读写默认使用 UTF-8 编码

6. **管理员权限**：部分 PowerShell 命令可能需要管理员权限

---

## 📝 更新日志

- **v1.0** - 初始版本
  - 基础工具集完成
  - Streamlit UI 实现
  - Ollama 集成
  - Office 文档处理
  - OCR 图像识别

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---


## 📧 联系方式

QQ邮箱: 2297468967@qq.com

---

<div align="center">

**⭐ 如果这个项目对你有帮助，请给一个 Star！**

</div>
