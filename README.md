# 文档转PDF转换工具

一个Windows专用的Python工具，用于将Excel和Word文档转换为PDF格式。输出黑白PDF文件，支持自动环境检查和依赖安装。

[English Version](#english-version) | 中文版

---

## 功能特点

- **Excel转PDF**: 将Excel文件的第一张工作表转换为PDF
- **Word转PDF**: 将整个Word文档转换为PDF
- **页面提取**: 通过关键词搜索从Word文档中提取特定页面
- **黑白输出**: 所有PDF转换为黑白格式，适合打印
- **批量处理**: 一次处理多个文件
- **自动依赖**: 自动检查并安装所需的Python包
- **错误日志**: 出错时自动生成`error.txt`日志文件

## 系统要求

- **操作系统**: 仅支持Windows
- **Python**: 3.7或更高版本
- **Microsoft Office**: 必须安装Excel和Word

## 安装

1. 克隆仓库:
   ```bash
   git clone https://github.com/yourusername/doc2pdf-converter.git
   cd doc2pdf-converter
   ```

2. 安装依赖（首次运行会自动安装）:
   ```bash
   pip install pywin32 pypdf
   ```

## 使用方法

运行脚本:
```bash
python doc2pdf.py
```

### 菜单选项

```
==================================================
  Document to PDF Converter
==================================================
  Select file type to convert:

  1. Area Summary Table (Excel, Sheet 1)
  2. Application Form (Word document)
  3. Quality Commitment Page (from Survey Report)
  4. Acceptance Report Page (from Survey Report)

  Multi-select supported:
    1,2,3,4  -> Convert all
    34       -> Convert options 3 and 4
    Enter    -> Convert all (default)
    0        -> Exit
```

### 文件匹配规则

工具会搜索以下文件模式（可在脚本中自定义）:

| 选项 | 文件模式 |
|------|---------|
| 1 | `*面积汇总表*.xls*` (面积汇总表) |
| 2 | `*审查申请表*.doc*` (审查申请表) |
| 3 | `*地籍调查报告*.doc*` (质量承诺书页面) |
| 4 | `*地籍调查报告*.doc*` (地籍调查成果验收报告页面，第2次出现) |

### 输出

- PDF文件保存到 `./out/` 文件夹
- 错误日志保存到 `./error.txt`

## 自定义

修改 `main()` 函数中的文件模式以匹配您的命名规则:

```python
excel_files = find_files(os.path.join(ROOT_DIR, '**', '*你的模式*.xls*'))
word_files = find_files(os.path.join(ROOT_DIR, '**', '*你的模式*.doc*'))
```

对于基于关键词的页面提取，修改搜索词:

```python
page_num = find_page_by_keyword(f, "你的关键词", occurrence=1)
```

## 项目结构

```
doc2pdf-converter/
├── doc2pdf.py          # 主脚本
├── README.md           # 说明文档
├── LICENSE             # MIT许可证
├── requirements.txt    # Python依赖
└── out/                # 输出文件夹（自动创建）
```

## 依赖

- `pywin32` - Windows COM接口，用于Office自动化
- `pypdf` - PDF操作库，用于黑白转换

## 常见问题

### 错误: "Microsoft Excel/Word not detected"
- 确保已安装Microsoft Office
- 尝试修复Office安装

### 错误: "File open failed"
- 关闭所有Office应用程序
- 检查文件是否被其他程序锁定
- 等待几秒后重试

### 错误: "pywin32 installation failed"
```bash
pip install pywin32
python Scripts/pywin32_postinstall.py -install
```

## 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件。

## 贡献

1. Fork本仓库
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建Pull Request

## 更新日志

### v1.0.0
- 初始发布
- Excel工作表1转PDF
- Word文档转PDF
- 基于关键词的页面提取
- 黑白PDF输出
- 多选菜单
- 自动依赖安装
- 错误日志记录

---

# English Version

A Windows-only Python tool for converting Excel and Word documents to PDF format. Outputs grayscale PDF files with automatic environment checking and dependency installation.

## Features

- **Excel to PDF**: Convert first sheet of Excel files to PDF
- **Word to PDF**: Convert entire Word documents to PDF
- **Page Extraction**: Extract specific pages from Word documents by keyword search
- **Grayscale Output**: All PDFs are converted to grayscale for printing
- **Batch Processing**: Process multiple files at once
- **Auto Dependency**: Automatically checks and installs required packages
- **Error Logging**: Generates `error.txt` when errors occur

## Requirements

- **Operating System**: Windows only
- **Python**: 3.7 or higher
- **Microsoft Office**: Excel and Word must be installed

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/doc2pdf-converter.git
   cd doc2pdf-converter
   ```

2. Install dependencies (auto-installed on first run):
   ```bash
   pip install pywin32 pypdf
   ```

## Usage

Run the script:
```bash
python doc2pdf.py
```

### File Patterns

The tool searches for files matching these patterns (customize in script):

| Option | File Pattern |
|--------|-------------|
| 1 | `*面积汇总表*.xls*` (Area Summary Table) |
| 2 | `*审查申请表*.doc*` (Application Form) |
| 3 | `*地籍调查报告*.doc*` (Quality Commitment page) |
| 4 | `*地籍调查报告*.doc*` (Acceptance Report page, 2nd occurrence) |

### Output

- PDF files are saved to `./out/` folder
- Error logs are saved to `./error.txt`

## Customization

Edit the file patterns in the `main()` function to match your naming conventions:

```python
excel_files = find_files(os.path.join(ROOT_DIR, '**', '*your_pattern*.xls*'))
word_files = find_files(os.path.join(ROOT_DIR, '**', '*your_pattern*.doc*'))
```

For keyword-based page extraction, modify the search terms:

```python
page_num = find_page_by_keyword(f, "Your Keyword", occurrence=1)
```

## Project Structure

```
doc2pdf-converter/
├── doc2pdf.py          # Main script
├── README.md           # This file
├── LICENSE             # MIT License
├── requirements.txt    # Python dependencies
└── out/                # Output folder (created automatically)
```

## Dependencies

- `pywin32` - Windows COM interface for Office automation
- `pypdf` - PDF manipulation for grayscale conversion

## Troubleshooting

### Error: "Microsoft Excel/Word not detected"
- Ensure Microsoft Office is installed
- Try repairing Office installation

### Error: "File open failed"
- Close all Office applications
- Check if file is locked by another program
- Wait a few seconds and retry

### Error: "pywin32 installation failed"
```bash
pip install pywin32
python Scripts/pywin32_postinstall.py -install
```

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Changelog

### v1.0.0
- Initial release
- Excel Sheet 1 to PDF conversion
- Word document to PDF conversion
- Keyword-based page extraction
- Grayscale PDF output
- Multi-select menu
- Auto dependency installation
- Error logging
