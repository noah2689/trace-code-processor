import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QTextEdit, QPushButton, QFileDialog,
                             QMessageBox, QGroupBox, QScrollArea, QFrame)
from PyQt5.QtCore import Qt, QSize, QSettings, QCoreApplication
from PyQt5.QtGui import QFont
import pandas as pd
from datetime import datetime

class TraceCodeProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("码上放心追溯码批量处理工具")
        self.setGeometry(100, 100, 900, 700)

        # 默认存储路径
        self.default_path = os.path.join(os.path.expanduser("~"), "Documents", "追溯码输出")

        # 初始化设置
        self.settings = QSettings("TraceCodeTool", "Settings")
        self.saved_path = self.settings.value("output_path", self.default_path)

        self.init_ui()

    def init_ui(self):
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)

        # 企业名称输入区域
        company_group = QGroupBox("企业/医院信息")
        company_layout = QHBoxLayout(company_group)

        company_label = QLabel("企业/医院名称:")
        self.company_input = QLineEdit()
        self.company_input.setPlaceholderText("请输入企业或医院名称")

        company_layout.addWidget(company_label)
        company_layout.addWidget(self.company_input)

        # 存储路径设置
        path_group = QGroupBox("存储设置")
        path_layout = QHBoxLayout(path_group)

        path_label = QLabel("存储路径:")
        self.path_input = QLineEdit()
        self.path_input.setText(self.saved_path)
        self.path_input.setReadOnly(True)

        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_path)

        path_layout.addWidget(path_label)
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(browse_btn)

        # 追溯码输入区域
        code_group = QGroupBox("追溯码列表")
        code_layout = QVBoxLayout(code_group)

        code_info = QLabel("请粘贴追溯码（每行一个）:")
        code_info.setStyleSheet("font-weight: bold;")

        self.code_textarea = QTextEdit()
        self.code_textarea.setPlaceholderText("在此粘贴追溯码，每行一个...\n例如：\n82050180000352480033\n82050180000576128589")
        self.code_textarea.setMinimumHeight(300)

        code_layout.addWidget(code_info)
        code_layout.addWidget(self.code_textarea)

        # 按钮区域
        button_layout = QHBoxLayout()

        self.generate_btn = QPushButton("生成表格")
        self.generate_btn.setStyleSheet("background-color: #4CAF50; color: white; font-size: 14px; padding: 10px;")
        self.generate_btn.clicked.connect(self.generate_table)

        self.reset_btn = QPushButton("重置")
        self.reset_btn.setStyleSheet("background-color: #f44336; color: white; font-size: 14px; padding: 10px;")
        self.reset_btn.clicked.connect(self.reset_form)

        self.open_folder_btn = QPushButton("打开文件夹")
        self.open_folder_btn.setStyleSheet("background-color: #2196F3; color: white; font-size: 14px; padding: 10px;")
        self.open_folder_btn.clicked.connect(self.open_folder)

        button_layout.addWidget(self.generate_btn)
        button_layout.addWidget(self.reset_btn)
        button_layout.addWidget(self.open_folder_btn)
        button_layout.addStretch()

        # 状态栏
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: gray; font-style: italic;")

        # 添加所有组件到主布局
        main_layout.addWidget(company_group)
        main_layout.addWidget(path_group)
        main_layout.addWidget(code_group)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.status_label)

        # 设置整体样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QTextEdit {
                border: 1px solid #cccccc;
                border-radius: 3px;
                font-family: Monaco, Consolas, monospace;
                font-size: 12px;
            }
            QLineEdit {
                border: 1px solid #cccccc;
                border-radius: 3px;
                padding: 5px;
            }
            QPushButton {
                border-radius: 5px;
                font-weight: bold;
            }
        """)

    def browse_path(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择存储文件夹",
            self.path_input.text(),
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        if folder:
            self.path_input.setText(folder)
            self.settings.setValue("output_path", folder)

    def generate_table(self):
        company = self.company_input.text().strip()
        if not company:
            QMessageBox.warning(self, "警告", "请输入企业/医院名称！")
            return

        trace_codes_text = self.code_textarea.toPlainText().strip()
        if not trace_codes_text:
            QMessageBox.warning(self, "警告", "请输入追溯码！")
            return

        # 解析追溯码，去除空行和空白
        raw_codes = trace_codes_text.split('\n')
        trace_codes = []
        for code in raw_codes:
            cleaned = code.strip()
            if cleaned:
                # 检测并拆分粘连的追溯码（40位数字 = 两个20位追溯码粘在一起）
                if len(cleaned) == 40 and cleaned.isdigit():
                    # 拆分为两个追溯码
                    code1 = cleaned[:20]
                    code2 = cleaned[20:]
                    trace_codes.append(code1)
                    trace_codes.append(code2)
                else:
                    trace_codes.append(cleaned)

        if not trace_codes:
            QMessageBox.warning(self, "警告", "没有有效的追溯码！")
            return

        # 去重（保持顺序）
        seen = set()
        unique_codes = []
        for code in trace_codes:
            if code not in seen:
                seen.add(code)
                unique_codes.append(code)

        # 统计信息
        total_count = len(trace_codes)
        unique_count = len(unique_codes)
        duplicate_count = total_count - unique_count

        # 创建DataFrame
        df = pd.DataFrame({
            '企业名称': [company] * unique_count,
            '追溯码': unique_codes
        })

        # 确保输出目录存在
        output_dir = self.path_input.text()
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 生成文件名（时间戳）
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{company}_{timestamp}.xlsx"
        filepath = os.path.join(output_dir, filename)

        try:
            # 保存Excel文件
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='追溯码数据')
                worksheet = writer.sheets['追溯码数据']

                # 自动调整列宽
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            # 更新状态
            status_msg = f"成功生成: {filename} | 总数: {total_count}, 唯一: {unique_count}, 重复: {duplicate_count}"
            self.status_label.setText(status_msg)

            # 询问是否自动重置
            reply = QMessageBox.question(
                self,
                "完成",
                f"表格已生成！\n文件: {filename}\n\n是否清空当前输入，继续处理下一个企业？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                self.reset_form()
            else:
                QMessageBox.information(self, "提示", "文件已保存，您可以继续添加其他企业的数据。")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成表格时出错：{str(e)}")

    def reset_form(self):
        self.company_input.clear()
        self.code_textarea.clear()
        self.status_label.setText("已重置")

    def open_folder(self):
        folder_path = self.path_input.text()
        if os.path.exists(folder_path):
            import platform
            system = platform.system()
            if system == "Darwin":  # macOS
                os.system(f"open '{folder_path}'")
            elif system == "Windows":  # Windows
                os.startfile(folder_path)
            else:
                QMessageBox.warning(self, "警告", f"不支持的操作系统: {system}")
        else:
            reply = QMessageBox.question(
                self,
                "警告",
                "存储路径不存在，将创建默认路径。是否继续？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                os.makedirs(folder_path, exist_ok=True)
                import platform
                system = platform.system()
                if system == "Darwin":
                    os.system(f"open '{folder_path}'")
                elif system == "Windows":
                    os.startfile(folder_path)

def main():
    # 在 QApplication 创建之前设置 DPI 属性
    QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QCoreApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)

    app = QApplication(sys.argv)

    window = TraceCodeProcessor()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
