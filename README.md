# PrintALL: All-in-One Printing Assistant
# PrintALL：全能打印助手

PrintALL is a powerful desktop application built with Python and Tkinter, designed to simplify common document and image processing tasks, specifically adding watermarks and batch printing. It supports various file types including Word documents, PDFs, and common image formats.
PrintALL 是一个使用 Python 和 Tkinter 构建的强大桌面应用程序，旨在简化常见的文档和图像处理任务，特别是添加水印和批量打印。它支持包括 Word 文档、PDF 和常见图像格式在内的多种文件类型。

## Features
## 功能特性

### Watermark Functionality
### 水印功能

- **Batch Watermarking**: Add custom watermarks to multiple files in a specified folder.
- **批量水印**: 为指定文件夹中的多个文件添加自定义水印。
- **Supported File Types**:
- **支持文件类型**:
    - **Word Documents (.docx)**: Adds "Print Attachment Name: [filename]" to the header and includes page numbers in the footer (if not already present).
    - **Word 文档 (.docx)**: 在页眉添加“打印附件名称：[文件名]”，并在页脚添加页码（如果尚未存在）。
    - **Image Files (.jpg, .png, .bmp)**: Adds "Print Attachment Name: [filename]" as a text watermark.
    - **图片文件 (.jpg, .png, .bmp)**: 添加“打印附件名称：[文件名]”作为文字水印。
    - **PDF Documents (.pdf)**: Adds "Print Attachment Name: [filename]" as a text watermark on each page.
    - **PDF 文档 (.pdf)**: 在每页添加“打印附件名称：[文件名]”作为文字水印。
- **Customizable Image Watermarks**: Adjust font size, opacity, and position (top-left, top-right, center, etc.) for image watermarks.
- **图片水印自定义**: 可调整图片水印的字体大小、透明度和位置（左上角、右上角、居中等）。
- **Intelligent Watermark Color**: Automatically adjusts image watermark color (black/white) based on the background brightness for better visibility.
- **智能水印颜色**: 根据背景亮度自动调整图片水印的颜色（黑/白），以提高可见性。
- **Log Output**: Provides real-time processing logs within the application interface.
- **日志输出**: 在应用程序界面内提供实时处理日志。

### Batch Printing Functionality
### 批量打印功能

- **Batch Printing**: Print multiple documents and images from a specified folder.
- **批量打印**: 从指定文件夹批量打印多个文档和图片。
- **Supported File Types**:
- **支持文件类型**:
    - **Word Documents (.doc, .docx)**: Requires LibreOffice for conversion and printing.
    - **Word 文档 (.doc, .docx)**: 需要 LibreOffice 进行转换和打印。
    - **PDF Documents (.pdf)**.
    - **PDF 文档 (.pdf)**。
    - **Image Files (.jpg, .png, .bmp)**: All selected images are merged into a single PDF before printing, ensuring consistent output and reducing individual print jobs.
    - **图片文件 (.jpg, .png, .bmp)**: 所有选定的图片在打印前会合并成一个 PDF 文件，确保输出一致性并减少单独的打印作业。
- **Printer Selection**: Specify the target printer by name.
- **打印机选择**: 通过名称指定目标打印机。
- **Page Count Filtering**: Filter Word and PDF documents based on their page count (e.g., print only documents with 1-2 pages).
- **页数筛选**: 根据页数筛选 Word 和 PDF 文档（例如，仅打印页数在 1-2 页之间的文档）。
- **Customizable LibreOffice Path**: Allows users to specify the exact path to their LibreOffice executable (`soffice.exe`), which is necessary for printing Word documents.
- **LibreOffice 路径自定义**: 用户可以指定 LibreOffice 可执行文件 (`soffice.exe`) 的确切路径，这对于打印 Word 文档是必需的。
- **Image Margins**: Adjust the margins for images when converted to PDF for printing.
- **图片页边距**: 调整图片转换为 PDF 打印时的页边距。
- **Automatic Default Printer Detection (Windows only)**: Automatically detects and populates the default printer name on Windows systems using `pywin32`.
- **自动检测默认打印机 (仅限 Windows)**: 在 Windows 系统上使用 `pywin32` 自动检测并填充默认打印机名称。
- **Log Output**: Provides real-time printing logs within the application interface.
- **日志输出**: 在应用程序界面内提供实时打印日志。

## Prerequisites
## 前提条件

- **Python 3.x**: Download and install from [python.org](https://www.python.org/).
- **Python 3.x**: 从 [python.org](https://www.python.org/) 下载并安装。
- **LibreOffice (for Word/DOCX/PDF printing)**: Download and install from [libreoffice.org](https://www.libreoffice.org/download/download-libreoffice/). This is essential for the batch printing feature, particularly for Word documents. The application requires the path to `soffice.exe`.
- **LibreOffice (用于 Word/DOCX/PDF 打印)**: 从 [libreoffice.org](https://www.libreoffice.org/download/download-libreoffice/) 下载并安装。这对于批量打印功能至关重要，特别是对于 Word 文档。应用程序需要 `soffice.exe` 的路径。
- **Microsoft YaHei Font (for Chinese watermarks)**: The application attempts to use `msyh.ttc` (Microsoft YaHei) for Chinese watermarks in PDFs. If this font is not found, it will fall back to a default font, which might affect Chinese character display. It's recommended to have this font installed on your system or place `msyh.ttc` in the same directory as the script.
- **微软雅黑字体 (用于中文水印)**: 应用程序尝试使用 `msyh.ttc`（微软雅黑）作为 PDF 中文水印的字体。如果找不到此字体，将回退到默认字体，这可能会影响中文字符的显示。建议将此字体安装到您的系统或将 `msyh.ttc` 放在脚本的同一目录中。

## Installation
## 安装

1.  **Clone or Download**: Clone this repository or download the `PrintALLApp.py` file.
    **克隆或下载**: 克隆此仓库或下载 `PrintALLApp.py` 文件。
    ```bash
    git clone https://github.com/your-username/PrintALL.git
    cd PrintALL
    ```
2.  **Install Dependencies**: Open your terminal or command prompt and navigate to the project directory. Run the following command to install the required Python libraries:
    **安装依赖**: 打开您的终端或命令提示符，导航到项目目录。运行以下命令安装所需的 Python 库：
    ```bash
    uv sync # use uv or pip ↓
    pip install Pillow python-docx PyPDF2 reportlab pypdf pywin32
    ```
    *Note: `pywin32` is Windows-specific. If you are on Linux/macOS, this package might fail to install, but the watermark functions will still work. The batch printing feature that relies on `pywin32` (e.g., auto-detecting printers) will be disabled.*
    *注意: `pywin32` 是 Windows 特定的。如果您在 Linux/macOS 上，此包可能安装失败，但水印功能仍将有效。依赖 `pywin32` 的批量打印功能（例如，自动检测打印机）将被禁用。*

## How to Use
## 如何使用

1.  **Run the Application**:
    **运行应用程序**:
    ```bash
    python PrintALLApp.py
    ```
    A graphical user interface (GUI) window will appear.
    将出现一个图形用户界面 (GUI) 窗口。

2.  **Watermark Tab**:
    **水印选项卡**:
    *   **Step 1: Select Folder**: Click "Browse..." to choose the folder containing the files you want to watermark.
    *   **第一步: 选择文件夹**: 点击“浏览...”选择包含要添加水印的文件的文件夹。
    *   **Step 2: Select File Types**: Check the boxes for "Word Documents", "Image Files", and/or "PDF Documents" based on what you want to process.
    *   **第二步: 选择文件类型**: 根据您要处理的文件类型，勾选“Word 文档”、“图片文件”和/或“PDF 文档”的复选框。
    *   **Step 3: Watermark Settings (for Images)**: Adjust "Font Size" (1-500), "Opacity" (0-255), and "Watermark Position" for image watermarks.
    *   **第三步: 水印详细设置 (主要用于图片)**: 调整图片水印的“字体大小”(1-500)、“透明度”(0-255) 和“水印位置”。
    *   **Start Processing**: Click "Batch Add Watermark". A confirmation dialog will appear, reminding you to back up your files as this operation modifies original files directly.
    *   **开始处理**: 点击“批量添加水印”。将出现一个确认对话框，提醒您备份文件，因为此操作会直接修改原始文件。

3.  **Batch Printing Tab**:
    **批量打印选项卡**:
    *   **Document/Image Folder**: Click "Select Folder" to choose the directory containing the files you want to print.
    *   **文档/图片文件夹**: 点击“选择文件夹”选择包含要打印的文件的目录。
    *   **Printer Name**: Enter the exact name of your printer as it appears in your system's printer settings. On Windows, the default printer might be auto-detected.
    *   **打印机名称**: 输入打印机在系统打印机设置中显示的准确名称。在 Windows 上，默认打印机可能会被自动检测到。
    *   **File Types to Print**: Select the checkboxes for `.doc`, `.docx`, `.pdf`, `.jpg`, `.png`, and/or `.bmp` to include them in the print queue.
    *   **要打印的文件类型**: 勾选 `.doc`、`.docx`、`.pdf`、`.jpg`、`.png` 和/或 `.bmp` 的复选框，将它们添加到打印队列中。
    *   **Page Count Filter**: Check "Enable Filter" and set "Pages >=" and "Pages <=" values to print only documents within a specific page range.
    *   **页码筛选**: 勾选“启用筛选”并设置“页数 >=”和“页数 <=”的值，以仅打印特定页数范围内的文档。
    *   **LibreOffice Path**: **Crucial for Word/PDF printing.** The default path is `C:\Program Files\LibreOffice\program\soffice.exe`. If your LibreOffice is installed elsewhere, click "Browse..." to locate `soffice.exe`.
    *   **LibreOffice 路径**: **对于 Word/PDF 打印至关重要。** 默认路径为 `C:\Program Files\LibreOffice\program\soffice.exe`。如果您的 LibreOffice 安装在其他位置，请点击“浏览...”找到 `soffice.exe`。
    *   **Image Margins (Pixels)**: Adjust this value to control spacing around images when they are converted to PDF for printing.
    *   **图片页边距 (像素)**: 调整此值以控制图片转换为 PDF 打印时的间距。
    *   **Start Printing**: Click "Start Batch Printing".
    *   **开始打印**: 点击“开始批量打印”。

## Important Notes
## 重要提示

*   **Backup Your Files**: Both watermark and printing functions modify or interact with original files. **ALWAYS BACK UP YOUR IMPORTANT DATA** before using this tool.
*   **备份您的文件**: 水印和打印功能都会修改或与原始文件交互。在使用此工具之前，**务必备份您的重要数据**。
*   **LibreOffice Requirement**: The batch printing feature relies heavily on LibreOffice for converting and printing `Word` documents. Ensure it's installed and the correct `soffice.exe` path is configured.
*   **LibreOffice 要求**: 批量打印功能严重依赖 LibreOffice 来转换和打印 `Word` 文档。请确保已安装 LibreOffice 并配置了正确的 `soffice.exe` 路径。
*   **Windows Only for Pywin32**: The `pywin32` library is used for Windows-specific features like automatically detecting the default printer and for more robust subprocess handling. If you are not on Windows, these specific functionalities might be unavailable, but the core logic should still work for other file types if LibreOffice is correctly configured.
*   **Pywin32 仅限 Windows**: `pywin32` 库用于 Windows 特定的功能，例如自动检测默认打印机和更强大的子进程处理。如果您不在 Windows 上，这些特定功能可能不可用，但如果 LibreOffice 配置正确，核心逻辑仍应适用于其他文件类型。
*   **Temporary Files**: The application creates temporary files during image-to-PDF conversion and page count checks. These are automatically cleaned up after the process.
*   **临时文件**: 应用程序在图片转换为 PDF 和页数检查过程中会创建临时文件。这些文件在处理完成后会自动清理。
*   **Logging**: All operations are logged to `PrintALL.log` in the same directory as the script for debugging and review.
*   **日志记录**: 所有操作都记录在与脚本位于同一目录的 `PrintALL.log` 文件中，以便调试和审查。

## Development & Contribution
## 开发与贡献

Feel free to fork the repository, make improvements, and submit pull requests.
欢迎自由地 Fork 仓库，进行改进并提交 Pull Request。

## License
## 许可证

This project is open-source and available under the MIT License.
本项目是开源的，并根据 MIT 许可证发布。

---