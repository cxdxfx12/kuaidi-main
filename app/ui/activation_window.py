"""
硬件绑定激活窗口
─────────────────
- 显示本机机器码
- 输入激活码（25字符，格式 XXXXX-XXXXX-XXXXX-XXXXX-XXXXX）
- 在线激活 / 离线激活
- 到期提醒
"""
import os
from typing import Optional

from PyQt5.QtWidgets import (
    QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QHBoxLayout, QMessageBox, QDialog,
    QFrame, QGraphicsDropShadowEffect,
    QApplication,
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QColor

from app.core.license_manager import (
    get_machine_code_display,
    verify_activation_key,
    save_license,
    get_license_info,
    check_license_on_startup,
    MSG_NOT_ACTIVATED,
    MSG_EXPIRED,
    MSG_MACHINE_CHANGED,
    MSG_TIME_TAMPERED,
)
from app.models.path_config import get_resource_path

APP_TITLE = "大圣快递账单结算系统"
COLOR_BG_TOP = "#6d4c2e"
COLOR_BG_BOTTOM = "#8b6239"


def _icon_path():
    p = os.path.join(get_resource_path("data", "icons"), "monkey-icon.png")
    if os.path.exists(p):
        return p
    return ""


class ActivationDialog(QDialog):
    """激活窗口 — 输入激活码完成硬件绑定"""

    activation_success = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_TITLE} - 软件激活")
        self.setFixedSize(480, 480)
        self.setModal(True)

        ip = _icon_path()
        if ip:
            self.setWindowIcon(QIcon(ip))

        self.setStyleSheet(f"""
            QDialog {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {COLOR_BG_TOP}, stop:1 {COLOR_BG_BOTTOM});
            }}
        """)

        self.machine_code = get_machine_code_display()
        self._build()

    def _build(self):
        main = QVBoxLayout(self)
        main.setContentsMargins(24, 20, 24, 20)
        main.setSpacing(12)

        # 标题
        title = QLabel("软件激活")
        f = QFont()
        f.setPointSize(16)
        f.setBold(True)
        title.setFont(f)
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color:#ffffff;")
        main.addWidget(title)

        subtitle = QLabel("请将本机机器码发送给管理员获取激活码")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color:rgba(255,255,255,180); font-size:12px;")
        main.addWidget(subtitle)

        # 白色卡片
        card = QFrame(self)
        card.setObjectName("activate_card")
        card.setStyleSheet("""
            QFrame#activate_card {
                background-color: #ffffff;
                border-radius: 12px;
            }
        """)
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(30)
        shadow.setYOffset(6)
        shadow.setColor(QColor(0, 0, 0, 80))
        card.setGraphicsEffect(shadow)

        cl = QVBoxLayout(card)
        cl.setContentsMargins(28, 20, 28, 18)
        cl.setSpacing(12)

        # ---- 机器码显示
        mc_label = QLabel("本机机器码")
        mc_label.setStyleSheet("color:#555; font-size:12px; font-weight:bold;")
        cl.addWidget(mc_label)

        mc_frame = QFrame()
        mc_frame.setStyleSheet("""
            QFrame {
                background-color: #f5f0eb;
                border: 1px solid #d4c5b5;
                border-radius: 6px;
            }
        """)
        mc_layout = QHBoxLayout(mc_frame)
        mc_layout.setContentsMargins(12, 10, 12, 10)

        self.mc_display = QLineEdit(self.machine_code)
        self.mc_display.setReadOnly(True)
        self.mc_display.setAlignment(Qt.AlignCenter)
        self.mc_display.setStyleSheet("""
            QLineEdit {
                border: none;
                background: transparent;
                font-size: 16px;
                font-weight: bold;
                letter-spacing: 2px;
                color: #5a3d24;
                font-family: 'Consolas', 'Courier New', monospace;
            }
        """)
        mc_layout.addWidget(self.mc_display, 1)

        copy_btn = QPushButton("复制")
        copy_btn.setFixedSize(60, 30)
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_btn.setStyleSheet("""
            QPushButton {
                background-color: #8b6239;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 11px;
            }
            QPushButton:hover { background-color: #a07850; }
        """)
        copy_btn.clicked.connect(self._copy_machine_code)
        mc_layout.addWidget(copy_btn)

        cl.addWidget(mc_frame)

        cl.addSpacing(4)

        # ---- 激活码输入
        ak_label = QLabel("激活码")
        ak_label.setStyleSheet("color:#555; font-size:12px; font-weight:bold;")
        cl.addWidget(ak_label)

        self.key_edit = QLineEdit()
        self.key_edit.setPlaceholderText("XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")
        self.key_edit.setFixedHeight(38)
        self.key_edit.setAlignment(Qt.AlignCenter)
        self.key_edit.setMaxLength(29)  # 25字符 + 4个横线
        self.key_edit.setStyleSheet("""
            QLineEdit {
                padding: 6px 12px;
                border: 1px solid #d7d7d7;
                border-radius: 6px;
                font-size: 14px;
                letter-spacing: 1px;
                font-family: 'Consolas', 'Courier New', monospace;
                background: #fafafa;
            }
            QLineEdit:focus { border: 2px solid #6d4c2e; background: #ffffff; }
        """)
        # 自动格式化：输入时每5个字符加横线
        self.key_edit.textChanged.connect(self._format_license_key)
        cl.addWidget(self.key_edit)

        cl.addSpacing(6)

        # ---- 提示
        hint = QLabel("激活码为25位字母数字组合，由管理员提供")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("color:#999; font-size:11px;")
        cl.addWidget(hint)

        cl.addSpacing(4)

        # ---- 激活按钮
        self.btn_activate = QPushButton("立即激活")
        self.btn_activate.setFixedHeight(40)
        self.btn_activate.setCursor(Qt.PointingHandCursor)
        self.btn_activate.setStyleSheet("""
            QPushButton {
                background-color: #6d4c2e;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #8b6239; }
            QPushButton:pressed { background-color: #5a3d24; }
            QPushButton:disabled { background-color: #ccc; }
        """)
        self.btn_activate.clicked.connect(self._on_activate)
        cl.addWidget(self.btn_activate)

        # ---- 试用按钮
        self.btn_trial = QPushButton("试用 7 天")
        self.btn_trial.setFixedHeight(34)
        self.btn_trial.setCursor(Qt.PointingHandCursor)
        self.btn_trial.setStyleSheet("""
            QPushButton {
                background-color: #eeeeee;
                color: #666;
                border: none;
                border-radius: 6px;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #e0e0e0; }
        """)
        self.btn_trial.clicked.connect(self._on_trial)
        cl.addWidget(self.btn_trial)

        # ---- 退出按钮
        self.btn_quit = QPushButton("退出软件")
        self.btn_quit.setFixedHeight(32)
        self.btn_quit.setCursor(Qt.PointingHandCursor)
        self.btn_quit.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #999;
                border: none;
                font-size: 11px;
            }
            QPushButton:hover { color: #555; text-decoration: underline; }
        """)
        self.btn_quit.clicked.connect(self.reject)
        cl.addWidget(self.btn_quit)

        main.addWidget(card, 1)

    def _format_license_key(self, text: str):
        """输入时自动格式化为 XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"""
        # 去掉已有横线和非字母数字字符
        raw = ''.join(c for c in text if c.isalnum())
        raw = raw.upper()
        # 限制25字符
        raw = raw[:25]
        # 每5字符加横线
        groups = [raw[i:i+5] for i in range(0, len(raw), 5)]
        formatted = '-'.join(groups)
        
        # 避免触发递归
        if formatted != text:
            self.key_edit.blockSignals(True)
            self.key_edit.setText(formatted)
            self.key_edit.blockSignals(False)

    def _copy_machine_code(self):
        """复制机器码到剪贴板"""
        clipboard = QApplication.clipboard()
        clipboard.setText(self.machine_code)
        QMessageBox.information(self, "已复制", "机器码已复制到剪贴板，请发送给管理员获取激活码")

    def _on_activate(self):
        """激活按钮点击"""
        key = self.key_edit.text().strip()
        if not key:
            QMessageBox.warning(self, "提示", "请输入激活码")
            return
        
        # 去掉横线验证长度
        clean = key.replace("-", "")
        if len(clean) < 25:
            QMessageBox.warning(self, "提示", "激活码不完整，请检查后重新输入")
            return
        
        # 显示等待提示
        self.btn_activate.setEnabled(False)
        self.btn_activate.setText("正在验证...")
        QApplication.processEvents()
        
        try:
            valid, msg, expire_date = verify_activation_key(self.machine_code, key)
            
            if valid and expire_date:
                # 保存授权
                ok, err = save_license(self.machine_code, key, expire_date)
                if ok:
                    QMessageBox.information(self, "激活成功", 
                        f"软件已成功激活！\n\n有效期至：{expire_date}\n\n感谢您的使用！")
                    self.activation_success.emit()
                    self.accept()
                else:
                    QMessageBox.critical(self, "激活失败", f"保存授权信息失败：\n{err}")
            else:
                QMessageBox.critical(self, "激活失败", msg)
        finally:
            self.btn_activate.setEnabled(True)
            self.btn_activate.setText("立即激活")

    def _on_trial(self):
        """开始7天试用"""
        from datetime import date, timedelta
        trial_expire = (date.today() + timedelta(days=7)).strftime("%Y-%m-%d")
        
        # 试用不需要激活码，直接保存授权
        ok, err = save_license(self.machine_code, "TRIAL", trial_expire)
        if ok:
            QMessageBox.information(self, "试用已开启",
                f"7天试用已开启！\n\n有效期至：{trial_expire}\n\n"
                "试用到期后请联系管理员获取正式激活码。")
            self.activation_success.emit()
            self.accept()
        else:
            QMessageBox.critical(self, "错误", f"无法保存试用信息：\n{err}")


# ===================== 对外入口 =====================

def show_activation_flow() -> bool:
    """
    显示激活流程
    返回 True 表示激活成功（或已激活），False 表示用户取消/退出
    """
    # 先检查是否已激活
    ok, _ = check_license_on_startup()
    if ok:
        return True

    # 未激活 → 显示激活窗口
    dlg = ActivationDialog()
    if dlg.exec_() == QDialog.Accepted:
        # 再次检查是否激活成功（activation_success 信号已触发 save_license）
        ok2, _ = check_license_on_startup()
        return ok2
    return False
