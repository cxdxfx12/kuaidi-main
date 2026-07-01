"""
自定义 UI 组件 - 从 main_window.py 提取
"""
import os

from PyQt5.QtWidgets import QFrame, QVBoxLayout, QLabel
from PyQt5.QtCore import Qt, pyqtSignal


class DropZone(QFrame):
    """文件拖拽区域 - 支持拖入文件，也可显示已选择的文件列表"""
    files_dropped = pyqtSignal(list)  # 拖入后发出信号，携带文件路径列表

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self._has_selection = False

        # 布局 - 垂直居中
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(20, 15, 20, 15)

        # 主标签
        self._main_label = QLabel("📄 请把账单Excel文件拖入框中")
        self._main_label.setAlignment(Qt.AlignCenter)
        self._main_label.setWordWrap(True)
        self._main_label.setStyleSheet("""
            font-size: 14px;
            color: #94a3b8;
            font-weight: 500;
        """)

        # 副标签
        self._sub_label = QLabel("或点击右侧「选择Excel文件」按钮（最多5个，支持 .xlsx / .xls / .csv）")
        self._sub_label.setAlignment(Qt.AlignCenter)
        self._sub_label.setStyleSheet("""
            font-size: 11px;
            color: #cbd5e1;
            margin-top: 10px;
        """)

        layout.addWidget(self._main_label)
        layout.addWidget(self._sub_label)

        self._apply_style(active=False)

    def set_selection(self, file_paths):
        """设置已选择文件 - 更新显示内容"""
        if file_paths:
            lines = [f"✅ 已选择 {len(file_paths)} 个文件："]
            for i, fp in enumerate(file_paths):
                try:
                    fsize = os.path.getsize(fp) / (1024 * 1024)
                except:
                    fsize = 0
                lines.append(f"  {i + 1}. {os.path.basename(fp)}  ({fsize:.1f} MB)")
            self._main_label.setText("\n".join(lines))
            self._main_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self._sub_label.setText("可继续拖入新文件（自动替换），或点击按钮添加更多")
            self._has_selection = True
            self._apply_style(active=True)
        else:
            self._main_label.setText("📄 请把账单Excel文件拖入框中")
            self._main_label.setAlignment(Qt.AlignCenter)
            self._sub_label.setText("或点击右侧「选择Excel文件」按钮（最多5个，支持 .xlsx / .xls / .csv）")
            self._has_selection = False
            self._apply_style(active=False)

    def _apply_style(self, active=False, dragging=False):
        if dragging:
            self.setStyleSheet("""
                QFrame {
                    background: #eef2ff;
                    border: 2px dashed #4f46e5;
                    border-radius: 12px;
                    min-height: 140px;
                }
            """)
        elif active:
            self.setStyleSheet("""
                QFrame {
                    background: #eef2ff;
                    border: 1px solid #c7d2fe;
                    border-radius: 12px;
                    min-height: 140px;
                }
            """)
        else:
            self.setStyleSheet("""
                QFrame {
                    background: #f8fafc;
                    border: 2px dashed #cbd5e1;
                    border-radius: 12px;
                    min-height: 140px;
                }
                QFrame:hover {
                    border-color: #94a3b8;
                    background: #f1f5f9;
                }
            """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            self._apply_style(active=False, dragging=True)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        # 离开时恢复到之前的状态（选中则显示选中样式，否则默认样式）
        self._apply_style(active=self._has_selection, dragging=False)
        event.accept()

    def dropEvent(self, event):
        file_paths = []
        seen = set()
        for url in event.mimeData().urls():
            if url.isLocalFile():
                path = url.toLocalFile()
                if path.lower().endswith(('.xlsx', '.xls', '.csv')) and path not in seen:
                    seen.add(path)
                    file_paths.append(path)

        if file_paths:
            # 恢复样式并发出信号
            self._apply_style(active=True, dragging=False)
            self.files_dropped.emit(file_paths)
            event.acceptProposedAction()
        else:
            self._apply_style(active=self._has_selection, dragging=False)
            event.ignore()
