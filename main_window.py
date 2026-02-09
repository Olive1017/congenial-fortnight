import sys
import os

root_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_path not in sys.path:
    sys.path.append(root_path)

from PySide6.QtWidgets import (
    QWidget, QPushButton, QLabel, QFileDialog,
    QVBoxLayout, QHBoxLayout, QTextEdit, QMessageBox, QApplication
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon, QColor
from PySide6.QtCore import QSize
from tools.splitter1 import split_excel_by_row
from tools.writer2 import set_smart_print_titles
from tools.stamper3 import add_stamp_to_excel


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📊 文档自动化工具")
        self.resize(800, 600)
        self.excel_paths = []

        # 设置窗口样式
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #1084d7;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton#successBtn {
                background-color: #28a745;
            }
            QPushButton#successBtn:hover {
                background-color: #34a853;
            }
            QPushButton#outputBtn {
                background-color: #6c63ff;
            }
            QPushButton#outputBtn:hover {
                background-color: #7b75ff;
            }
            QLabel {
                color: #333;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 8px;
                font-family: 'Courier New';
                font-size: 11px;
            }
        """)

        self.excel_path = None
        self.output_files = []

        self.init_ui()

    def init_ui(self):
        # ===== 标题 =====
        title_label = QLabel("文档自动化处理工具")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)

        # ===== 文件选择区域 =====
        file_label_title = QLabel("输入文件:")
        file_label_title.setFont(self._get_section_font())

        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: #666; padding: 8px; background-color: white; border-radius: 3px;")

        self.select_btn = QPushButton("📁 选择 Excel")
        self.select_btn.setMinimumWidth(120)
        self.select_btn.clicked.connect(self.select_file)

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.select_btn)
        file_layout.addWidget(self.file_label, 1)

        file_group_layout = QVBoxLayout()
        file_group_layout.addWidget(file_label_title)
        file_group_layout.addLayout(file_layout)

        # ===== 处理按钮区域 =====
        button_label_title = QLabel("操作:")
        button_label_title.setFont(self._get_section_font())

        self.run_btn = QPushButton("▶️ 开始处理")
        self.run_btn.setObjectName("successBtn")
        self.run_btn.setMinimumHeight(40)
        self.run_btn.clicked.connect(self.run_process)

        self.export_btn = QPushButton("💾 导出文件")
        self.export_btn.setObjectName("outputBtn")
        self.export_btn.setMinimumHeight(40)
        self.export_btn.clicked.connect(self.export_files)
        self.export_btn.setEnabled(False)

        self.clear_log_btn = QPushButton("🗑️ 清空日志")
        self.clear_log_btn.setMinimumHeight(40)
        self.clear_log_btn.clicked.connect(self.clear_log)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.run_btn)
        button_layout.addWidget(self.export_btn)
        button_layout.addWidget(self.clear_log_btn)

        button_group_layout = QVBoxLayout()
        button_group_layout.addWidget(button_label_title)
        button_group_layout.addLayout(button_layout)

        # ===== 日志区域 =====
        log_label_title = QLabel("处理日志:")
        log_label_title.setFont(self._get_section_font())

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(250)

        log_group_layout = QVBoxLayout()
        log_group_layout.addWidget(log_label_title)
        log_group_layout.addWidget(self.log_box)

        # ===== 状态栏 =====
        status_layout = QHBoxLayout()
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: #28a745; padding: 5px;")
        status_layout.addWidget(self.status_label)
        status_layout.addStretch()

        # ===== 主布局 =====
        main_layout = QVBoxLayout()
        main_layout.addWidget(title_label)
        main_layout.addSpacing(10)
        main_layout.addLayout(file_group_layout)
        main_layout.addSpacing(10)
        main_layout.addLayout(button_group_layout)
        main_layout.addSpacing(10)
        main_layout.addLayout(log_group_layout)
        main_layout.addSpacing(10)
        main_layout.addLayout(status_layout)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(5)

        self.setLayout(main_layout)

    def _get_section_font(self):
        font = QFont()
        font.setPointSize(11)
        font.setBold(True)
        return font

    def select_file(self):

        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择 Excel 文件", "", "Excel Files (*.xlsx)"
        )

        if paths:
            self.excel_paths = paths  # 保存多个文件路径
            file_names = [os.path.basename(path) for path in paths]
            self.file_label.setText(
                f"已选择 {len(paths)} 个文件: {', '.join(file_names[:3])}{'...' if len(file_names) > 3 else ''}")
            self.log(f"✅ 已选择 {len(paths)} 个文件")
            self._update_status(f"已选择 {len(paths)} 个文件，可以开始处理", "#0078d4")

    def run_process(self):
        """核心处理逻辑：批量处理多个文件"""
        if not hasattr(self, 'excel_paths') or not self.excel_paths:
            QMessageBox.warning(self, "提示", "请先选择 Excel 文件")
            return

        # 获取当前文件 (main_window.py) 的绝对路径：c:\Users\xinan\PycharmProjects\excel_handle\
        root_dir = os.path.dirname(os.path.abspath(__file__))
        # 直接进入 pic 目录：c:\Users\xinan\PycharmProjects\excel_handle\pic\stamp.png
        stamp_path = os.path.join(root_dir, "pic", "stamp.png")

        self._update_status("正在批量处理文件...", "#ff9800")
        self.log("=" * 50)
        self.log("🚀 开始批量处理任务")

        try:
            import tempfile
            from tools.splitter1 import split_excel_by_row
            from tools.writer2 import set_smart_print_titles
            from tools.stamper3 import add_stamp_to_excel

            all_output_files = []  # 保存所有文件的输出路径

            # 循环处理每个文件
            for file_idx, excel_path in enumerate(self.excel_paths, 1):
                self.log(f"\n📁 处理文件 {file_idx}/{len(self.excel_paths)}: {os.path.basename(excel_path)}")
                self.log("-" * 40)

                temp_dir = tempfile.mkdtemp()

                # 获取原文件名（不含扩展名）用于输出命名
                input_filename = os.path.splitext(os.path.basename(excel_path))[0]
                temp_prefix = os.path.join(temp_dir, input_filename)

                # --- Step 1: 拆分 ---
                self.log("Step 1: 正在拆分 Excel...")
                split_files = split_excel_by_row(excel_path, temp_prefix)
                if not split_files:
                    self.log("❌ 未生成任何拆分文件")
                    continue
                self.log(f"✅ 拆分完成，生成 {len(split_files)} 个文件")

                # --- Step 2 & 3: 循环处理子文件 ---
                self.log("Step 2 & 3: 执行表头固定与自动盖章...")

                for idx, file_path in enumerate(split_files, 1):
                    f_name = os.path.basename(file_path)

                    # 1. 设置打印固定行 (writer2)
                    ok_h, msg_h = set_smart_print_titles(file_path)

                    # 2. 盖章 (stamp3)
                    ok_s, msg_s = add_stamp_to_excel(file_path, stamp_path)

                    # 日志记录
                    self.log(f"  [{idx}] {f_name}")
                    self.log(f"      └─ 表头: {'✅' if ok_h else '❌'} {msg_h}")
                    self.log(f"      └─ 印章: {'✅' if ok_s else '❌'} {msg_s}")

                all_output_files.extend(split_files)

            # --- 任务完成 ---
            self.output_files = all_output_files
            self.log("=" * 50)
            self.log(
                f"🎉 批量处理完毕！共处理 {len(self.excel_paths)} 个输入文件，生成 {len(all_output_files)} 个输出文件。")

            self.export_btn.setEnabled(True)
            self._update_status(f"✅ 处理完成，共生成 {len(all_output_files)} 个文件", "#28a745")
            QMessageBox.information(self, "完成",
                                    f"所有文件已处理完毕！\n输入: {len(self.excel_paths)} 个文件\n输出: {len(all_output_files)} 个文件")

        except Exception as e:
            self.log(f"❌ 流程中断: {str(e)}")
            self._update_status("❌ 处理失败", "#f44336")
            QMessageBox.critical(self, "错误", f"处理失败：\n{str(e)}")

    def export_files(self):
        """导出所有文件到指定目录"""
        if not self.output_files:
            QMessageBox.warning(self, "提示", "没有待导出的文件")
            return

        # 打开目录选择对话框
        output_dir = QFileDialog.getExistingDirectory(
            self, "选择导出文件夹", ""
        )

        if not output_dir:
            self.log("⚠️ 已取消导出")
            return

        try:
            self._update_status("导出中...", "#ff9800")
            self.log("=" * 50)
            self.log(f"开始导出到: {output_dir}")
            self.log("-" * 50)

            import shutil

            # 创建输出目录
            os.makedirs(output_dir, exist_ok=True)

            exported_files = []
            for idx, source_file in enumerate(self.output_files, 1):
                filename = os.path.basename(source_file)
                dest_file = os.path.join(output_dir, filename)
                shutil.copy2(source_file, dest_file)
                exported_files.append(dest_file)
                self.log(f"  {idx}. {filename}")

            self.log("-" * 50)
            self.log(f"✅ 导出完成 共导出 {len(exported_files)} 个文件")
            self.log("=" * 50)

            self._update_status(f"✅ 导出完成 ({len(exported_files)} 个文件)", "#28a745")
            QMessageBox.information(self, "✅ 导出完成", f"成功导出 {len(exported_files)} 个文件到:\n{output_dir}")

        except Exception as e:
            self.log(f"❌ 导出失败: {str(e)}")
            self.log("=" * 50)
            self._update_status("❌ 导出失败", "#f44336")
            QMessageBox.critical(self, "❌ 错误", f"导出失败：{str(e)}")

    def clear_log(self):
        """清空日志"""
        self.log_box.clear()
        self.log("日志已清空")

    def log(self, text):
        """添加日志"""
        self.log_box.append(text)

    def _update_status(self, text, color):
        """更新状态标签"""
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"color: {color}; padding: 5px; font-weight: bold;")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())