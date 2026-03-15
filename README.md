# LocalAgent

**[English](README.md) | [中文](README_zh.md)**

> **A powerful local AI agent tool** — designed for computer operations and task automation, optimized for small local models

<div align="center">

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![Ollama](https://img.shields.io/badge/Ollama-Latest-green.svg)](https://ollama.ai/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.55.0-red.svg)](https://streamlit.io/)

</div>

---

## 📖 Project Overview

LocalAgent is an intelligent agent application built on **local large models via Ollama** and **Streamlit**. It understands natural language instructions from users and automatically invokes various tools to perform complex computer operations, such as file handling, document conversion, code execution, image recognition, and more.

### Core Features

- 🤖 **Local AI-Driven** - Runs local large models using Ollama (supports Qwen, GLM, etc.)
- 🛠️ **Rich Tool Set** - Includes over 18 built-in tools covering file operations, code execution, document processing, and more
- 💻 **System Interaction** - Can execute PowerShell commands and manipulate the file system
- 📄 **Office Support** - Full support for reading, writing, and converting Word, Excel, and PPT files
- 🖼️ **Image Recognition** - Integrated OCR functionality to extract text from images
- 🔄 **Automated Tool Creation** - Supports packaging Python code into reusable CLI tools

---

## 🚀 Quick Start

### System Requirements

- Python 3.10+
- Windows operating system
- [Ollama](https://ollama.ai/) (must be installed and configured beforehand)

### Installation Steps

1. **Clone or download the project**

```bash
cd project_directory
```

2. **Create and activate virtual environment**

```bash
python -m venv .venv
.venv\Scripts\activate
```

If you encounter the error "Execution of scripts is disabled on this system," run the following command as administrator and retry:
```
set-ExecutionPolicy RemoteSigned
```

3. **Initialize external files**
```
.\setup.cmd
```

4. **Install dependencies**

```bash
pip install -r requirements.txt
```

5. **Launch the application**

Option 1: Use the launch script (can be double-clicked)
```cmd
.\launch.cmd
```

Option 2: Manual startup
```bash
.venv\Scripts\activate && streamlit run streamlit_app.py
```

---

## 🛠️ Functional Modules

### Tool List

| Category               | Function                              | Description                     |
| ---------------------- | ------------------------------------- | ------------------------------- |
| **File Operations**    | `read_file`                           | Read file content               |
|                        | `create_file`                         | Create a new file               |
|                        | `replace_file_content`                | Replace part of file content    |
| **Code Execution**     | `run_powershell`                      | Execute PowerShell commands     |
|                        | `run_python_code`                     | Run Python code                 |
|                        | `create_python_tool`                  | Create reusable Python tool     |
| **Word Documents**     | `read_word_and_export_txt`            | Read Word and export as text    |
|                        | `convert_word_or_txt_to_pdf`          | Convert Word/TXT to PDF         |
|                        | `convert_markdown_to_word`            | Convert Markdown to Word        |
| **Excel Spreadsheets** | `read_excel_and_export_txt`           | Read Excel and export as text   |
|                        | `create_excel_from_2d_list`           | Create Excel from 2D list       |
| **PPT Presentations**  | `read_ppt_and_export_txt`             | Read PPT and export as text     |
|                        | `replace_ppt_content`                 | Replace PPT content             |
|                        | `create_ppt_from_txt`                 | Create PPT from text            |
|                        | `convert_ppt_to_pdf`                  | Convert PPT to PDF              |
| **Image Recognition**  | `recognize_image_and_export_markdown` | OCR: recognize text in images   |
| **Others**             | `add_knowledge`                       | Add knowledge to knowledge base |
|                        | `wait_user_do`                        | Wait for user to complete task  |

### Task Examples

```
✅ Calculate 160968*(23516-75061)
✅ Create an HTML Snake game
✅ Make a PPT to introduce yourself
✅ Convert Word documents on the desktop to PDF
✅ Organize word image files on the desktop by part of speech into a Word document
✅ Generate line and pie charts based on data.xlsx on the desktop
```

---

## 📁 Project Structure

```
LocalAgent/
├── streamlit_app.py      # Main Streamlit application
├── tools.py              # Tool function library
├── ocr.py                # OCR image recognition module
├── markdown_to_word.py   # Markdown to Word converter
├── run.py                # Startup script
├── launch.cmd            # Windows launch command
├── requirements.txt      # Python dependencies
├── prompt/
│   └── prompt.md         # AI system prompt template
├── images/               # UI icon resources
├── output/               # Output file directory
└── demo/                 # Demo file directory
```

---

## ⚙️ Configuration Guide

### Ollama Model Configuration

Ensure the Ollama service is running at `http://localhost:11434/` and that the required models have been pulled:

```bash
# Recommended models:

ollama pull qwen3.5:35b # 23G, runs on one 4060 GPU + 32GB RAM, best performance
ollama pull glm-4.7-flash # 19G, runs on one 4060 GPU + 32GB RAM, best performance
ollama pull qwen3-coder:latest # 18G, runs on one 4060 GPU + 32GB RAM, no reasoning output

ollama pull qwen3.5:9b # 6.6G, runs smoothly on one 4060 GPU
ollama pull qwen3:8b # 5.2G, runs smoothly on one 4060 GPU

ollama pull qwen3.5:4b # 3.4G, runs on CPU
ollama pull qwen3:4b # 2.5G, runs on CPU
ollama pull qwen3:4b-instruct # 2.5G, runs on CPU, no reasoning output

# Suggested: model size >= 4B, context length >= 8K. If download stalls, press Ctrl+C to stop and resume.
```

### External Files Directory

The project depends on the external directory `D:/ExternalFiles/`:

- Stores generated temporary files
- Holds user-defined tools (`.py` files)
- Contains tool usage instructions (`.md` files)
- Stores knowledge base file (`KNOWLEDGE.txt`)

### System Prompt

`prompt/prompt.md` contains the AI system prompt template with variable substitution support:

- `$EXTERNALFILES$` - List of external files
- `$KNOWLEDGE$` - User's knowledge base content

---

## 💡 Usage Tips

1. **Stop a task**: Click the `Stop` button, then select `Rerun` from the `...` menu.

2. **Create reusable tools**: Use `create_python_tool` to encapsulate frequently used code into CLI tools for improved efficiency.

3. **Fix code errors**: Use `replace_file_content` to quickly patch bugs in Python code.

4. **Add knowledge**: Use `add_knowledge` to record problem-solving techniques, system information, etc.

5. **Path conventions**: Always use absolute paths. On Windows, use backslashes `\`.

---

## 🔧 Development Guide

### Adding New Tools

1. Define the tool function in `tools.py` using standard docstring format:

```python
def my_new_tool(param1, param2):
    """
    Tool description

    Args:
      param1 (str): Description of parameter 1
      param2 (int): Description of parameter 2

    Returns:
      str: Status message
    """
    # Implementation code
    return "Operation successful"
```

2. Register the tool in `streamlit_app.py`:

```python
st.session_state.tools.append(my_new_tool)
st.session_state.available_functions["my_new_tool"] = my_new_tool
# Remember to update both st.session_state.tools and st.session_state.available_functions
```

### Dependency Overview

Key libraries used:

| Library       | Purpose                                |
| ------------- | -------------------------------------- |
| `streamlit`   | Web UI framework                       |
| `ollama`      | Local large model client               |
| `python-docx` | Word document processing               |
| `openpyxl`    | Excel document processing              |
| `python-pptx` | PPT document processing                |
| `pywin32`     | Windows COM interface (PDF conversion) |
| `paddleocr`   | OCR text recognition                   |

---

## ⚠️ Important Notes

1. **Local Models**: Choose models that support tool calling, such as the `qwen3` series. Set context length to 16K or higher.

2. **OCR Functionality**: OCR requires additional installation of [PaddleOCR-VL](https://www.paddleocr.ai/main/version3.x/pipeline_usage/PaddleOCR-VL.html)

3. **Security**: Tools can execute arbitrary PowerShell commands and Python code — use only in trusted environments.

4. **Output Limits**: Some tools have output length restrictions (e.g., 2500/3000 characters)

5. **Encoding**: File I/O uses UTF-8 encoding by default

6. **Admin Privileges**: Some PowerShell commands may require administrator rights

---

## 📝 Changelog

- **v1.0** - Initial Release
  - Basic toolset complete
  - Streamlit UI implemented
  - Ollama integration
  - Office document processing
  - OCR image recognition

---

## 🤝 Contribution

We welcome issues and pull requests!

---

## 📧 Contact

QQ Email: 2297468967@qq.com

---

<div align="center">

**⭐ If this project helps you, please give it a Star!**

</div>