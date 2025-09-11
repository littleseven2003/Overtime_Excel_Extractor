# -*- coding: utf-8 -*-

"""
加班数据提取和处理应用程序入口

该脚本是加班数据提取工具的主入口点。
它整合了GUI界面和核心业务逻辑，提供完整的应用程序功能。
"""

import sys
import os
from PyQt5.QtWidgets import QApplication, QMessageBox

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from gui import OvertimeExtractorApp
from core import OvertimeDataProcessor


class OvertimeExtractorGUI(OvertimeExtractorApp):
    """加班数据提取GUI应用程序"""
    
    def process_data(self):
        """处理加班数据"""
        # 检查是否选择了源文件
        if not self.source_file:
            QMessageBox.warning(self, "警告", "请选择源文件")
            return
        
        # 检查文件是否存在
        if not os.path.exists(self.source_file):
            QMessageBox.warning(self, "警告", "源文件不存在")
            return
        
        # 获取选中的研究室
        selected_labs = [cb.text() for cb in self.lab_checkboxes if cb.isChecked()]
        
        # 检查是否选择了研究室
        if not selected_labs:
            QMessageBox.warning(self, "警告", "请至少选择一个研究室")
            return
        
        try:
            # 创建数据处理器
            processor = OvertimeDataProcessor(
                source_file=self.source_file,
                output_dir=self.output_dir if self.output_dir else "",
                selected_labs=selected_labs
            )
            
            # 处理数据
            output_path = processor.process_data(log_callback=self.log_message)
            
            # 显示完成消息
            QMessageBox.information(self, "完成", f"数据提取完成！\n结果已保存到: {output_path}")
            self.statusBar().showMessage("处理完成")
            
        except Exception as e:
            error_msg = str(e)
            self.log_message(error_msg)
            QMessageBox.critical(self, "错误", error_msg)
            self.statusBar().showMessage("处理失败")


def main():
    app = QApplication(sys.argv)
    window = OvertimeExtractorGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()