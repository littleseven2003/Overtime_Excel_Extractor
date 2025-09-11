# -*- coding: utf-8 -*-

"""
加班数据提取和处理应用程序入口

该脚本是加班数据提取工具的主入口点。
它整合了GUI界面和核心业务逻辑，提供完整的应用程序功能。
"""

import sys
import os
from PyQt5.QtWidgets import QApplication, QMessageBox

# 添加项目根目录到Python路径，确保可以正确导入项目模块
# 由于当前文件在src目录下，需要添加上级目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gui import OvertimeExtractorApp
from core import OvertimeDataProcessor


class OvertimeExtractorGUI(OvertimeExtractorApp):
    """加班数据提取GUI应用程序类
    
    继承自OvertimeExtractorApp，实现了具体的数据处理逻辑。
    负责协调GUI界面和核心业务逻辑之间的交互。
    """
    
    def process_data(self):
        """处理加班数据
        
        执行完整的数据处理流程：
        1. 验证输入参数（源文件、研究室选择等）
        2. 创建数据处理器实例
        3. 调用核心处理逻辑
        4. 显示处理结果或错误信息
        
        该方法通过重写父类的抽象方法来实现具体的数据处理逻辑。
        """
        # 1. 验证必要条件 - 检查是否选择了源文件
        if not self.source_file:
            QMessageBox.warning(self, "警告", "请选择源文件")
            return
        
        # 2. 验证文件是否存在
        if not os.path.exists(self.source_file):
            QMessageBox.warning(self, "警告", "源文件不存在")
            return
        
        # 3. 获取选中的研究室列表
        selected_labs = [cb.text() for cb in self.lab_checkboxes if cb.isChecked()]
        
        # 4. 验证是否选择了研究室
        if not selected_labs:
            QMessageBox.warning(self, "警告", "请至少选择一个研究室")
            return
        
        try:
            # 5. 创建数据处理器实例
            processor = OvertimeDataProcessor(
                source_file=self.source_file,
                output_dir=self.output_dir if self.output_dir else "",
                selected_labs=selected_labs
            )
            
            # 6. 处理数据，并传入日志回调函数以便在GUI中显示处理进度
            output_path = processor.process_data(log_callback=self.log_message)
            
            # 7. 显示完成消息，通知用户处理成功
            QMessageBox.information(
                self, 
                "完成", 
                f"数据提取完成！\n结果已保存到: {output_path}"
            )
            self.statusBar().showMessage("处理完成")
            
        except Exception as e:
            # 8. 处理异常情况，显示错误信息
            error_msg = str(e)
            self.log_message(error_msg)  # 在日志区域显示错误信息
            QMessageBox.critical(self, "错误", error_msg)  # 弹出错误对话框
            self.statusBar().showMessage("处理失败")


def main():
    """应用程序主函数
    
    创建Qt应用程序实例，初始化主窗口并启动事件循环。
    这是程序的入口点。
    """
    # 创建Qt应用程序实例
    app = QApplication(sys.argv)
    
    # 创建主窗口实例
    window = OvertimeExtractorGUI()
    
    # 显示主窗口
    window.show()
    
    # 启动事件循环
    sys.exit(app.exec_())


if __name__ == "__main__":
    """程序入口点
    
    当脚本直接运行时执行main函数启动应用程序。
    """
    main()