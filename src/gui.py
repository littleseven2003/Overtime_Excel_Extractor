# -*- coding: utf-8 -*-

"""
加班数据提取和处理 GUI 应用程序

该应用程序提供图形界面用于从Excel文件中提取指定研究室的加班数据，
并生成格式化的结果文件。

功能包括：
1. 选择源数据文件（.xlsx或.xls）
2. 拖放文件支持
3. 选择结果文件保存位置
4. 选择要提取的研究室
5. 执行数据提取和处理
"""

import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QFileDialog, QTextEdit, 
    QMessageBox, QGroupBox, QCheckBox
)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QFont


class DragDropLabel(QLabel):
    """支持拖放文件的标签控件类
    
    允许用户通过拖放方式选择Excel文件，提供直观的文件选择体验。
    """
    
    def __init__(self, parent=None):
        """初始化拖放标签控件
        
        Args:
            parent: 父级控件，默认为None
        """
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setText("将Excel文件拖放到这里\n或点击下方按钮选择文件")
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 20px;
                margin: 10px;
                background-color: #f0f0f0;
            }
        """)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """处理拖入事件
        
        当用户将文件拖拽到标签上时调用此方法，接受包含URL的拖拽数据。
        
        Args:
            event (QDragEnterEvent): 拖入事件对象
        """
        if event.mimeData().hasUrls():
            # 先接受拖拽动作，具体的文件类型检查在dropEvent中进行
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def dropEvent(self, event: QDropEvent):
        """处理放下事件
        
        当用户在标签上释放拖拽的文件时调用此方法，验证文件类型并设置源文件。
        
        Args:
            event (QDropEvent): 放下事件对象
        """
        if event.mimeData().hasUrls():
            file_path = event.mimeData().urls()[0].toLocalFile()
            if file_path.lower().endswith(('.xlsx', '.xls')):
                # 查找主窗口并调用set_source_file方法
                main_window = self.window()
                if hasattr(main_window, 'set_source_file'):
                    main_window.set_source_file(file_path)
            else:
                # 查找主窗口并显示错误消息
                main_window = self.window()
                QMessageBox.warning(main_window, "文件类型错误", "请选择Excel文件（.xlsx或.xls）")


class OvertimeExtractorApp(QMainWindow):
    """加班数据提取应用程序主窗口类
    
    提供完整的图形用户界面，包括文件选择、研究室选择、数据处理和日志显示功能。
    """
    
    def __init__(self):
        """初始化应用程序主窗口"""
        super().__init__()
        self.source_file = ""  # 源文件路径
        self.output_dir = ""   # 输出目录路径
        self.selected_labs = []  # 选中的研究室列表
        self.research_labs = ['智能通信', '数据算法', '新型能源', '新型材料']  # 研究室列表
        self.init_ui()
    
    def init_ui(self):
        """初始化用户界面
        
        创建和布局所有GUI组件，包括文件选择区域、研究室选择区域、
        处理按钮和日志显示区域。
        """
        self.setWindowTitle("加班数据提取工具")
        self.setGeometry(100, 100, 600, 500)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 标题标签
        title_label = QLabel("加班数据提取工具")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout(file_group)
        
        # 拖放区域
        self.drag_drop_label = DragDropLabel(self)
        file_layout.addWidget(self.drag_drop_label)
        
        # 文件选择按钮
        file_button_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("选择源文件")
        self.select_file_btn.clicked.connect(self.select_source_file)
        self.select_output_btn = QPushButton("选择保存位置")
        self.select_output_btn.clicked.connect(self.select_output_dir)
        file_button_layout.addWidget(self.select_file_btn)
        file_button_layout.addWidget(self.select_output_btn)
        file_layout.addLayout(file_button_layout)
        
        # 文件路径显示标签
        self.source_file_label = QLabel("源文件: 未选择")
        self.output_dir_label = QLabel("保存位置: 默认为源文件目录")
        file_layout.addWidget(self.source_file_label)
        file_layout.addWidget(self.output_dir_label)
        
        main_layout.addWidget(file_group)
        
        # 研究室选择区域
        lab_group = QGroupBox("研究室选择")
        lab_layout = QVBoxLayout(lab_group)
        
        # 创建复选框用于选择研究室
        self.lab_checkboxes = []
        
        # 使用左右布局形式显示研究室选择和按钮
        lab_main_layout = QHBoxLayout()
        
        # 左侧：研究室选择（2列网格布局）
        lab_grid_layout = QGridLayout()
        for i, lab in enumerate(self.research_labs):
            checkbox = QCheckBox(lab)
            self.lab_checkboxes.append(checkbox)
            
            # 分两列显示
            row = i // 2
            col = i % 2
            lab_grid_layout.addWidget(checkbox, row, col, Qt.AlignLeft)
        
        lab_main_layout.addLayout(lab_grid_layout)
        
        # 添加间隙
        lab_main_layout.addSpacing(20)
        
        # 右侧：全选/取消全选按钮（垂直布局）
        button_layout = QVBoxLayout()
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_labs)
        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.clicked.connect(self.deselect_all_labs)
        
        button_layout.addWidget(self.select_all_btn)
        button_layout.addWidget(self.deselect_all_btn)
        button_layout.addStretch()  # 添加弹性空间
        
        lab_main_layout.addLayout(button_layout)
        
        # 创建一个水平布局来居中整个研究室选择区域
        lab_container = QHBoxLayout()
        lab_container.addStretch()
        lab_container.addLayout(lab_main_layout)
        lab_container.addStretch()
        
        lab_layout.addLayout(lab_container)
        
        main_layout.addWidget(lab_group)
        
        # 开始处理按钮（单独一行）
        self.process_btn = QPushButton("开始处理")
        self.process_btn.clicked.connect(self.process_data)
        main_layout.addWidget(self.process_btn)
        
        # 操作按钮（清空日志和退出）
        button_layout = QHBoxLayout()
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.clear_log)
        self.exit_btn = QPushButton("退出")
        self.exit_btn.clicked.connect(self.close)
        
        button_layout.addWidget(self.clear_log_btn)
        button_layout.addWidget(self.exit_btn)
        main_layout.addLayout(button_layout)
        
        # 日志输出区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        main_layout.addWidget(log_group)
        
        # 状态栏
        self.statusBar().showMessage("就绪")
    
    def set_source_file(self, file_path):
        """设置源文件路径并更新界面显示
        
        Args:
            file_path (str): 源文件的完整路径
        """
        self.source_file = file_path
        self.source_file_label.setText(f"源文件: {os.path.basename(file_path)}")
        self.statusBar().showMessage(f"已选择源文件: {file_path}")
        self.log_message(f"已选择源文件: {file_path}")
    
    def select_source_file(self):
        """打开文件选择对话框选择源文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择源文件", 
            "", 
            "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.set_source_file(file_path)
    
    def select_output_dir(self):
        """打开目录选择对话框选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择保存位置")
        if dir_path:
            self.output_dir = dir_path
            self.output_dir_label.setText(f"保存位置: {dir_path}")
            self.statusBar().showMessage(f"已选择保存位置: {dir_path}")
            self.log_message(f"已选择保存位置: {dir_path}")
    
    def select_all_labs(self):
        """全选所有研究室复选框"""
        for cb in self.lab_checkboxes:
            cb.setChecked(True)
    
    def deselect_all_labs(self):
        """取消全选所有研究室复选框"""
        for cb in self.lab_checkboxes:
            cb.setChecked(False)
    
    def clear_log(self):
        """清空日志显示区域"""
        self.log_text.clear()
        self.log_message("日志已清空")
    
    def log_message(self, message):
        """在日志区域添加消息并滚动到底部
        
        Args:
            message (str): 要添加到日志的消息
        """
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        QApplication.processEvents()  # 确保界面及时更新
    
    def process_data(self):
        """处理数据的抽象方法
        
        子类必须实现此方法以提供具体的数据处理逻辑。
        
        Raises:
            NotImplementedError: 当子类未实现此方法时抛出
        """
        raise NotImplementedError("需要在子类中实现process_data方法")


def main():
    """GUI应用程序入口点
    
    创建Qt应用程序实例，初始化主窗口并启动事件循环。
    """
    app = QApplication(sys.argv)
    window = OvertimeExtractorApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    """程序入口点
    
    当脚本直接运行时执行main函数启动应用程序。
    """
    main()