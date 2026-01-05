# AutoTable - 基于大语言模型的自动化智能表格生成应用

## 概述
传统的通过python来实现表格自动化操作的应用可能很难处理复杂的docx格式的表格，而大语言模型擅长处理非结构化的字段，从而增加了灵活性。本项目通过用户自定义的“知识库”文件，实现基于大语言模型对docx文件表格的自动填写。原理是使用大语言模型来分析表格结构，生成字段到单元格位置的映射。然后根据这个映射填充数据。支持本地Ollama调用和符合OpenAI协议的API。

本项目现已提供 **Web 可视化界面**，操作更加便捷。

我最近开发了另一款不依赖大语言模型的离线自动化填表应用，可以尝试https://github.com/Dukeway/Autowordtable

## 主要功能
✅ **智能填写表格**：基于大语言模型分析表格结构，自动映射并填充数据。  
✅ **多格式知识库支持**：除了 **Excel (.xlsx)**，现已新增支持从 **Word (.docx)** 文档中智能提取数据作为知识库，同时也支持 **JSON** 格式。  
✅ **Web 可视化界面**：提供基于 Streamlit 的 Web 界面，支持文件拖拽上传、在线配置模型、结果下载及历史记录管理。  
✅ **智能段落填充**：自动识别文档中的下划线占位符，智能填入内容并保持原有下划线格式；支持追加模式，防止覆盖原有标签（如“姓名：”）。  
✅ **上下文感知**：在处理多个相似表格或段落时，能够自动切换不同的实体数据（如不同的人员或项目），避免重复填充。  
✅ **双模引擎支持**：支持本地 **Ollama** 模型和云端 **OpenAI 协议 API**（如 DeepSeek, GPT-4）。  
✅ **局域网共享**：Web 模式启动后可自动获取本机 IP，方便局域网内其他设备访问。

## 快速开始

### 1. 安装
```bash
git clone https://github.com/Dukeway/Autotable.git
cd AutoTable
pip install -r requirements.txt
```

### 2. 运行应用
**推荐使用 Web 界面模式：**
```bash
streamlit run Autotable/app.py
```
启动后浏览器会自动打开，你也可以通过终端显示的 `http://localhost:8501` 或局域网 IP 访问。

**传统命令行模式：**
```bash
python Autotable/main.py
```

## 应用设置 (命令行模式)
如果你选择使用命令行模式，请按照以下步骤设置：
1. 找到文件夹里的“知识库.xlsx"文件，将表格中的""字段"和“字段值”分别填写为你需要的内容。
2. 文件夹有个“表格模版.docx"文件，替换为需要填写的表格文件，注意名称同样为"表格模版.docx”。
3. 在config.py文件中进行设置，选择“api”或者“ollama"模式进行项目运行；如果使用api，那么需要填写你的秘钥；如果使用ollama，确保ollama在本地运行。注意正确填写模型名称。

*(注：如果使用 Web 界面，上述步骤均可在网页中直接通过上传文件和填写表单完成)*

## 应用示例
### 知识库
![知识库示例](figures/知识库示例.png)

### 表格模版
![示例模板](figures/word模版示例.png)

自动化输出如下：
![自动化流程](figures/自动填表示例.png)

## 说明
由于依赖大语言模型解析表格的能力，因此小参数模型表现并不好。项目演示使用的是DeepSeek-V3，你也可以选择其他优秀的模型。如果通过ollama来运行，注意复杂的表格很可能无法识别，从而导致字段值插入错误位置。

如果想修改提示词，可以在autotable.py文件中修改。

---

# AutoTable - An Automated Smart Table Generation Application Based on Large Language Models

## Overview
Traditional applications that use Python to automate table operations may struggle to handle complex table structures in .docx format, whereas large language models excel at processing unstructured fields, offering greater flexibility. This project enables automatic table filling in .docx files based on a user-defined "knowledge base" file, leveraging large language models. The principle involves using a large language model to analyze the table structure, generate a mapping of fields to cell positions, and then populate the data according to this mapping. It supports local Ollama calls and APIs compatible with the OpenAI protocol.

The project now includes a **Web UI** for easier operation.

I recently developed another offline automated form filling application that does not rely on a large language model. You can try it https://github.com/Dukeway/Autowordtable

## Features
✅ **Smart Table Filling**: Analyzes table structures using LLMs to automatically map and fill data.
✅ **Multi-Format Knowledge Base**: Supports **Excel (.xlsx)**, **Word (.docx)** (via intelligent extraction), and **JSON** files.
✅ **Web Interface**: Provides a Streamlit-based Web UI for easy file uploads, configuration, downloading results, and managing history.
✅ **Smart Paragraph Filling**: Automatically detects underlined placeholders to preserve formatting; supports append mode to avoid overwriting labels (e.g., "Name: ").
✅ **Context Awareness**: Automatically switches entities (e.g., different people or projects) when processing multiple similar tables/sections to avoid duplication.
✅ **Dual Engine Support**: Supports local **Ollama** models and cloud **OpenAI-compatible APIs** (e.g., DeepSeek, GPT-4).
✅ **LAN Sharing**: Automatically detects local IP for easy access from other devices on the same network.

## Installation

### 1. Install
```bash
git clone https://github.com/Dukeway/Autotable.git
cd AutoTable
pip install -r requirements.txt
```

### 2. Run Application
**Web UI Mode (Recommended):**
```bash
streamlit run Autotable/app.py
```
The browser will open automatically. You can also access it via `http://localhost:8501` or your LAN IP.

**CLI Mode:**
```bash
python Autotable/main.py
```

## Application Setup (CLI Mode)
If you choose to use the CLI mode:
1. Locate the "知识库.xlsx" (knowledge base) file in the folder and fill in the "fields" and "field values" columns with the content you need.
2. Replace the "表格模版.docx" (table template) file in the folder with the table file you want to fill, ensuring the name remains "表格模版.docx".
3. Configure settings in the config.py file by selecting either "api" or "ollama" mode to run the project. If using an API, provide your API key; if using Ollama, ensure it is running locally. Make sure to correctly specify the model name.

*(Note: If using the Web UI, all the above steps can be done directly in the browser by uploading files and filling forms)*

## Application Example
Since the code comments and files are in Chinese, the pictures are not shown here (they are shown in the Chinese introduction). You can still optimize this project by looking at the code, because this project is not perfect.

## Notes
Since the project relies on the table-parsing capabilities of large language models, smaller-parameter models may underperform. The project demo uses DeepSeek-V3, but you can choose other high-performing models as well. When running via Ollama, note that complex tables may not be recognized correctly, potentially causing field values to be inserted into the wrong positions.

If you want to modify the prompts, you can edit them in the autotable.py file.
