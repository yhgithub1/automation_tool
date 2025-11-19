# Automation Tool

这是一个自动化工具项目，包含多个模块用于不同的自动化任务。

## 项目结构

```
automation_tool/
├── .gitignore              # Git忽略文件配置
├── README.md              # 项目说明文档
├── correction.py          # 主程序文件
├── correction.spec        # PyInstaller配置文件
├── file_converter.py      # 文件转换工具
├── main.py               # 主入口文件
├── requirements.txt      # Python依赖包列表
├── robot-solid-full.svg  # 项目图标
├── Automation tool使用说明.pdf  # 使用说明文档
├── modules/              # 功能模块目录
│   ├── __init__.py
│   ├── calibration_report_demo.py
│   ├── config.py
│   ├── excel_manager.py
│   ├── file_converter.py
│   ├── file_converter_always.py
│   ├── file_converter_ui.py
│   ├── findfile.py
│   ├── folder_creation.py
│   ├── memo_generator.py
│   ├── outlook_automation.py
│   └── pdf_extractor.py
├── test_files/           # 测试文件目录
│   ├── test_data.xlsx
│   ├── test_document.docx
│   ├── test_image.jpg
│   └── test_image.png
└── utils/               # 工具函数目录
    ├── __init__.py
    └── file_utils.py
```

## 功能模块

- **文件转换**: 支持多种格式的文件转换
- **Excel管理**: Excel文件处理和管理功能
- **PDF提取**: PDF文档内容提取
- **Outlook自动化**: 邮件和日历自动化
- **文件查找**: 快速查找文件工具
- **文件夹创建**: 自动化文件夹创建
- **备忘录生成**: 自动生成备忘录文档
- **校准报告**: 生成校准报告模板

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

运行主程序:
```bash
python correction.py
```

## 构建可执行文件

使用PyInstaller构建可执行文件:

```bash
pyinstaller correction.spec
```

构建产物将在`dist/`目录下生成。

## 注意事项

- 本项目使用了多个第三方库，请确保安装了所有依赖
- 构建产物（`build/`、`dist/`目录）和IDE配置文件（`.idea/`）已被.gitignore排除
- 测试文件仅用于功能验证，实际使用时请替换为真实文件

## 版本信息

- Python版本: 3.11+
- 主要依赖: 根据requirements.txt文件
