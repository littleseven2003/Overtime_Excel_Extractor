# 更新日志

## [1.0.0] - 2025-09-11

### 新增功能
- 图形用户界面(GUI)应用程序，用于从Excel文件中提取加班数据
- 支持拖放文件操作
- 可选择多个研究室进行数据提取
- 自动生成格式化的Excel结果文件
- Windows批处理启动脚本

### 技术特性
- 基于PyQt5的现代化图形界面
- 使用pandas进行高效数据处理
- 使用openpyxl进行Excel文件读写和格式化
- 支持.xlsx和.xls格式文件

### 系统要求
- Python 3.6或更高版本
- 支持Windows、macOS和Linux操作系统

### 文件结构
- `src/main.py` - 应用程序入口点
- `src/gui.py` - 图形界面模块
- `src/core.py` - 核心业务逻辑模块
- `start_gui.bat` - Windows启动脚本
- `requirements.txt` - 依赖包列表
- `README.md` - 项目文档