"""
激活码生成器 — 管理员专用 GUI
=============================
输入用户机器码 + 到期日期 → 一键生成激活码
"""

import sys
import os

# 确保项目根目录在路径中
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from datetime import datetime, date, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QGroupBox, QMessageBox,
    QDateEdit, QSizePolicy
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor, QPalette

from app.core.license_manager import generate_activation_key


class KeygenWindow(QMainWindow):
    """激活码生成器主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("大圣·快递物流派费结算系统 — 激活码生成器")
        self.setFixedSize(520, 420)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }
            QGroupBox {
                font-size: 13px;
                font-weight: bold;
                color: #1a1a2e;
                border: 1px solid #d0d5dd;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 20px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 16px;
                padding: 0 8px;
            }
            QLabel {
                color: #333;
            }
            QLineEdit {
                border: 1px solid #d0d5dd;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 14px;
                background: white;
            }
            QLineEdit:focus {
                border-color: #5b6af0;
            }
            QDateEdit {
                border: 1px solid #d0d5dd;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 14px;
                background: white;
            }
            QDateEdit:focus {
                border-color: #5b6af0;
            }
            QPushButton {
                border-radius: 6px;
                padding: 10px 24px;
                font-size: 14px;
                font-weight: bold;
                color: white;
                background-color: #5b6af0;
                border: none;
            }
            QPushButton:hover {
                background-color: #4a56d4;
            }
            QPushButton:pressed {
                background-color: #3b45b8;
            }
            QPushButton#btnCopy {
                background-color: #f0f2f5;
                color: #333;
                border: 1px solid #d0d5dd;
                font-size: 12px;
                padding: 6px 16px;
            }
            QPushButton#btnCopy:hover {
                background-color: #e0e3ea;
            }
            QLabel#lblResult {
                font-size: 18px;
                font-weight: bold;
                color: #5b6af0;
                background-color: #eef0ff;
                border-radius: 6px;
                padding: 12px;
                border: 1px solid #d0d5ff;
            }
            QLabel#lblHint {
                color: #888;
                font-size: 11px;
            }
        """)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        # ── 输入区 ──
        input_group = QGroupBox("输入信息")
        input_layout = QVBoxLayout(input_group)
        input_layout.setSpacing(12)

        # 机器码
        mc_layout = QHBoxLayout()
        mc_label = QLabel("机器码：")
        mc_label.setFixedWidth(70)
        self.mc_input = QLineEdit()
        self.mc_input.setPlaceholderText("XXXX-XXXX-XXXX-XXXX-XXXX（20位，从激活窗口复制）")
        self.mc_input.textChanged.connect(self._on_input_changed)
        mc_layout.addWidget(mc_label)
        mc_layout.addWidget(self.mc_input)
        input_layout.addLayout(mc_layout)

        # 到期日期
        date_layout = QHBoxLayout()
        date_label = QLabel("到期日期：")
        date_label.setFixedWidth(70)
        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDisplayFormat("yyyy-MM-dd")
        # 默认一年后
        default_date = date.today() + timedelta(days=365)
        self.date_input.setDate(QDate(default_date.year, default_date.month, default_date.day))
        self.date_input.setMinimumDate(QDate.currentDate())
        self.date_input.dateChanged.connect(self._on_input_changed)
        date_layout.addWidget(date_label)
        date_layout.addWidget(self.date_input)
        date_layout.addStretch()
        input_layout.addLayout(date_layout)

        layout.addWidget(input_group)

        # ── 生成按钮 ──
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.btn_generate = QPushButton("🔑 生成激活码")
        self.btn_generate.setFixedSize(200, 44)
        self.btn_generate.clicked.connect(self._on_generate)
        self.btn_generate.setEnabled(False)
        btn_layout.addWidget(self.btn_generate)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # ── 结果区 ──
        result_group = QGroupBox("激活码")
        result_layout = QVBoxLayout(result_group)
        result_layout.setSpacing(8)

        self.lbl_result = QLabel("")
        self.lbl_result.setObjectName("lblResult")
        self.lbl_result.setAlignment(Qt.AlignCenter)
        self.lbl_result.setWordWrap(True)
        self.lbl_result.setMinimumHeight(50)
        result_layout.addWidget(self.lbl_result)

        copy_layout = QHBoxLayout()
        copy_layout.addStretch()
        self.btn_copy = QPushButton("📋 复制激活码")
        self.btn_copy.setObjectName("btnCopy")
        self.btn_copy.setFixedSize(140, 34)
        self.btn_copy.clicked.connect(self._on_copy)
        self.btn_copy.setEnabled(False)
        copy_layout.addWidget(self.btn_copy)
        copy_layout.addStretch()
        result_layout.addLayout(copy_layout)

        layout.addWidget(result_group)

        # ── 底部提示 ──
        hint = QLabel("提示：激活码需配合机器码使用，换机器后需重新生成。到期日期过期后需续费。")
        hint.setObjectName("lblHint")
        hint.setAlignment(Qt.AlignCenter)
        hint.setWordWrap(True)
        layout.addWidget(hint)

        layout.addStretch()

        # 状态变量
        self._generated_key = ""

    def _on_input_changed(self):
        """输入变化时启/禁用生成按钮"""
        mc = self.mc_input.text().strip()
        has_input = len(mc) > 0
        self.btn_generate.setEnabled(has_input)

    def _on_generate(self):
        """生成激活码"""
        machine_code = self.mc_input.text().strip()
        expire_str = self.date_input.date().toString("yyyy-MM-dd")

        # 验证机器码格式
        clean_mc = machine_code.replace("-", "").upper()
        if len(clean_mc) != 20 or not all(c in "0123456789ABCDEF" for c in clean_mc):
            QMessageBox.warning(
                self, "格式错误",
                "机器码格式不正确！\n\n期望格式：XXXX-XXXX-XXXX-XXXX-XXXX（20位十六进制字符）\n\n请从激活窗口完整复制机器码。"
            )
            return

        # 验证日期
        try:
            parsed = datetime.strptime(expire_str, "%Y-%m-%d").date()
        except ValueError:
            QMessageBox.warning(self, "日期错误", "到期日期格式不正确！")
            return

        try:
            key = generate_activation_key(machine_code, expire_str)
        except Exception as e:
            QMessageBox.critical(self, "生成失败", f"激活码生成异常：\n{str(e)}")
            return

        self._generated_key = key
        self.lbl_result.setText(key)
        self.btn_copy.setEnabled(True)

        remaining = (parsed - date.today()).days
        QMessageBox.information(
            self, "生成成功",
            f"激活码已生成！\n\n有效期至：{expire_str}\n剩余天数：{remaining} 天\n\n请将激活码发给用户完成激活。"
        )

    def _on_copy(self):
        """复制激活码到剪贴板"""
        if self._generated_key:
            clipboard = QApplication.clipboard()
            clipboard.setText(self._generated_key)
            self.btn_copy.setText("✅ 已复制！")
            # 1.5秒后恢复
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1500, lambda: self.btn_copy.setText("📋 复制激活码"))


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = KeygenWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
