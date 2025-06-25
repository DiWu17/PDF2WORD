# PDF-to-Word 智能转换工具

## 1. 项目简介

本项目是一个功能强大的 Python 工具，旨在将 PDF 文件转换为高质量、易于编辑的 Word (.docx) 文档。它深度整合了先进的文档布局分析技术和大型语言模型 (VLM)，能够智能处理包括复杂表格在内的各种文档元素，实现从内容到格式的精准转换。

与传统的转换工具不同，本项目能更好地理解文档结构，特别是将图片形式的表格重新结构化为真正的 Word 表格，而非简单的图片粘贴。

## 2. 主要功能

- **两种转换模式**:
    - **内容优先 (`content`)**: 专注于提取和重构文档的核心内容。此模式首先将 PDF 解析为结构化的 Markdown 文件，在此过程中利用大模型 API (如阿里云 DashScope) 识别并转换表格，最终生成格式流畅、内容准确的 Word 文档。非常适合需要二次编辑和内容复用的场景。
    - **格式优先 (`format`)**: 致力于通过自动化操作（如 `pywin32`），最大程度地保留原始 PDF 的视觉版式和布局。适用于追求高保真度视觉还原的场景。

- **强大的表格处理**: `content` 模式的核心亮点。它能够将 PDF 中的图片表格发送给大模型进行分析，并将其转换为标准的、可编辑的表格，而不是一张无法修改的图片。

- **完全开源**: 项目完全基于开源库（如 `pypandoc`, `python-docx`）实现核心的 MD 到 DOCX 转换，无任何商业库依赖，彻底告别水印问题。

- **自动化与易用性**: 提供简单的命令行接口和开箱即用的示例脚本，实现了从 PDF 解析到最终 Word 文档生成的一键式自动化流程。

## 3. 环境要求

- **Python**: 建议使用 `Python 3.8` 或更高版本。
- **Pandoc**: 本项目的核心依赖之一。您需要在系统中全局安装 Pandoc。
    - 可从官网下载安装：[https://pandoc.org/installing.html](https://pandoc.org/installing.html)
    - 安装后，请确保 `pandoc` 命令已添加到系统的环境变量中。
- **大模型 API Key**: `content` 模式下的表格处理功能需要调用云端大模型服务。
    - 本项目当前配置为使用阿里云的 **DashScope** 服务。
    - 您需要在根目录下的 `Config.py` 文件中配置您的 API Key。打开该文件并填入 `DASHSCOPE_API_KEY` 变量。
      ```python
      # Config.py

      # 在此文件中找到并更新您的 DashScope API Key
      DASHSCOPE_API_KEY = "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
      ```

## 4. 安装步骤

1.  **克隆项目**
    ```bash
    git clone <您的项目仓库地址>
    cd PDF2WORD
    ```

2.  **创建并激活虚拟环境 (推荐)**
    -   On macOS/Linux:
        ```bash
        python3 -m venv venv
        source venv/bin/activate
        ```
    -   On Windows:
        ```bash
        python -m venv venv
        .\venv\Scripts\activate
        ```

3.  **安装依赖包**
    ```bash
    pip install -r requirements.txt
    ```

## 5. 使用方法

您可以直接通过命令行调用主程序 `main.py` 来处理您自己的 PDF 文件。

-   **命令格式**:
    ```bash
    python main.py -i <输入PDF路径> -o <输出目录> -m <模式>
    ```

-   **参数说明**:
    -   `-i`, `--input`: **(必需)** 指定要转换的 PDF 文件的路径。
    -   `-o`, `--output`: **(必需)** 指定生成的 Word 文档要保存的目录。
    -   `-m`, `--mode`: **(可选)** 指定转换模式，可选值为 `content` 或 `format`。默认为 `content`。

-   **示例**:
    -   项目 `input/` 目录下已提供一个 `sample_2.pdf` 文件可供测试。
    -   **使用 `content` 模式进行转换**:
        ```bash
        python main.py -i "input/sample_2.pdf" -o "output/" -m content
        ```
    -   **使用 `format` 模式进行转换**:
        ```bash
        python main.py -i "input/sample_2.pdf" -o "output/" -m format
        ```

## 6. 注意事项

- 首次运行 `content` 模式时，程序会自动从 ModelScope 下载所需的本地模型文件，可能会花费一些时间，请耐心等待。
- `format` 模式依赖于 Windows 环境和 Office 软件的自动化接口，可能无法在非 Windows 系统上运行。
- 请确保您已按照 **环境要求** 部分的说明，正确配置了 `pandoc` 和大模型 API Key。

## 许可证

本项目仅供学习和个人使用。 