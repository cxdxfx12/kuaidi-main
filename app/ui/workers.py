"""
工作线程模块 - 从 main_window.py 提取，避免 UI 卡死
所有耗时操作（计算、导入、导出、加载）都放在后台线程中执行
"""
import os
import json
import re as _re
import traceback

from PyQt5.QtCore import QThread, pyqtSignal


class CalculateWorker(QThread):
    """计算后台线程，避免UI卡死 + 支持进度反馈 + 多文件批量处理"""
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)  # (percent, stage_text)  percent: 0-100
    file_progress = pyqtSignal(int, int, str)  # (file_index, total_files, file_name)

    def __init__(self, file_paths, sheet_name=None):
        super().__init__()
        # 兼容单文件和多文件
        if isinstance(file_paths, (list, tuple)):
            self.file_paths = list(file_paths)
        else:
            self.file_paths = [file_paths]
        self.sheet_name = sheet_name

    def run(self):
        try:
            from app.services.calculate_service import CalculateService
            service = CalculateService()
            total_files = len(self.file_paths)
            all_results = []
            total_rows_all = 0
            total_fee_all = 0.0
            total_success_all = 0
            total_exception_all = 0
            failed_files = []

            for file_idx, file_path in enumerate(self.file_paths):
                file_name = os.path.basename(file_path)
                self.file_progress.emit(file_idx + 1, total_files, file_name)

                # 为每个文件定义进度回调，映射到总体进度
                def make_cb(fidx, total):
                    def cb(percent, stage_text):
                        # 每个文件占 100/total % 的总体进度
                        overall = int((fidx * 100 + percent) / total)
                        overall = min(max(overall, 0), 100)
                        self.progress.emit(overall, f"[第{fidx+1}/{total}个] {stage_text}")
                    return cb

                cb = make_cb(file_idx, total_files)

                try:
                    result = service.import_and_calculate(file_path, self.sheet_name, progress_callback=cb)
                    all_results.append({
                        "file_name": file_name,
                        "record_id": result["record_id"],
                        "total_fee": result["total_fee"],
                        "success_count": result["success_count"],
                        "exception_count": result["exception_count"],
                        "total_rows": result["total_rows"]
                    })
                    total_rows_all += result.get("total_rows", 0)
                    total_fee_all += result.get("total_fee", 0.0)
                    total_success_all += result.get("success_count", 0)
                    total_exception_all += result.get("exception_count", 0)
                except Exception as file_err:
                    failed_files.append(f"{file_name}: {file_err}")
                    continue

            # 发送100%完成信号
            self.progress.emit(100, "全部文件处理完成")

            # 返回汇总结果（第一个记录的id用于切换到结果tab）
            final_result = {
                "record_id": all_results[0]["record_id"] if all_results else None,
                "total_fee": total_fee_all,
                "success_count": total_success_all,
                "exception_count": total_exception_all,
                "total_rows": total_rows_all,
                "file_count": total_files,
                "files": all_results,
                "failed_files": failed_files
            }
            self.finished.emit(final_result)
        except Exception as e:
            self.error.emit(str(e))


class FilePreviewWorker(QThread):
    """文件预览后台线程 - 解决加载大Excel时界面无响应问题"""
    finished = pyqtSignal(dict)  # (parse_result)
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)  # 读取过程中的进度反馈

    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            self.progress.emit(5, "开始读取文件...")
            from app.services.excel_parser import ExcelParser
            parser = ExcelParser()

            def cb(percent):
                self.progress.emit(int(percent), f"读取文件中... {percent}%")

            parse_result = parser.parse(self.file_path, row_callback=cb)
            self.finished.emit(parse_result)
        except Exception as e:
            self.error.emit(str(e))


class ResultLoadWorker(QThread):
    """结果数据加载后台线程 - 仅加载显示所需数据，避免UI卡死和内存溢出"""
    finished = pyqtSignal(dict)  # {summary: {...}, prepared: [rows], total_rows, display_count}
    error = pyqtSignal(str)

    def __init__(self, record_id, max_display=20000):
        super().__init__()
        self.record_id = record_id
        self.max_display = max_display

    def run(self):
        try:
            from app.models.database import get_session
            from app.models.fee_detail import FeeDetail
            from app.models.fee_record import FeeRecord

            session = get_session()
            try:
                # 优先从 FeeRecord 中直接读取已存储的汇总值（更高效）
                record = session.query(FeeRecord).filter(
                    FeeRecord.id == self.record_id
                ).first()

                if record:
                    total_rows = int(record.total_rows or 0)
                    total_fee = float(getattr(record, "total_fee", 0) or 0)
                    success_count = int(record.success_rows or 0)
                    exception_count = int(record.error_rows or 0)
                else:
                    # 回退：实际查询统计
                    total_rows = session.query(FeeDetail).filter(
                        FeeDetail.record_id == self.record_id
                    ).count()
                    total_fee = 0.0
                    exception_count = 0
                    # 分批统计以避免内存溢出
                    batch_size = 5000
                    offset = 0
                    while True:
                        batch = session.query(FeeDetail).filter(
                            FeeDetail.record_id == self.record_id
                        ).order_by(FeeDetail.id).limit(batch_size).offset(offset).all()
                        if not batch:
                            break
                        for d in batch:
                            total_fee += float(d.calculated_fee or 0)
                            if d.is_exception:
                                exception_count += 1
                        offset += batch_size
                    success_count = total_rows - exception_count

                summary = {
                    "total_rows": total_rows,
                    "total_fee": total_fee,
                    "success_count": success_count,
                    "exception_count": exception_count,
                }

                # 仅取前 max_display 行用于显示
                display_details = session.query(FeeDetail).filter(
                    FeeDetail.record_id == self.record_id
                ).order_by(FeeDetail.row_index).limit(self.max_display).all()
            finally:
                session.close()

            display_count = len(display_details)

            # 预解析为纯Python元组（避免传递SQLAlchemy对象）
            prepared = []
            for d in display_details:
                business_date = ""
                customer_name = ""
                try:
                    if d.original_data:
                        od = d.original_data if isinstance(d.original_data, dict) else json.loads(d.original_data)
                        raw_date = od.get("business_date", "")
                        if raw_date:
                            s = str(raw_date)
                            digits = _re.findall(r"\d+", s)
                            if len(digits) >= 3:
                                business_date = f"{int(digits[0]):04d}/{int(digits[1]):02d}/{int(digits[2]):02d}"
                            else:
                                business_date = s
                        customer_name = str(od.get("customer_name", "") or "")
                except Exception:
                    pass

                prepared.append((
                    str(d.row_index),
                    business_date,
                    d.tracking_no or "",
                    f"{d.station_code or ''} {d.station_name or ''}",
                    d.region_name or "",
                    f"{float(d.weight or 0):.3f}",
                    customer_name,
                    str(d.quantity or 0),
                    f"{float(d.calculated_fee or 0):.2f}",
                    d.rule_name or "",
                    "⚠️ 异常" if d.is_exception else "✓"
                ))

            # 重要：不再传递完整details，仅传递显示所需数据
            # 结算和导出功能会按需独立加载数据
            self.finished.emit({
                "summary": summary,
                "prepared": prepared,
                "total_rows": summary["total_rows"],
                "display_count": display_count,
            })
        except Exception as e:
            self.error.emit(str(e) + "\n" + traceback.format_exc())


class ExportWorker(QThread):
    """导出 Excel 后台线程 - 避免大数据量导出时 UI 未响应"""
    finished = pyqtSignal(str)  # 成功：返回文件路径
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)  # 进度：百分比 + 消息

    def __init__(self, record_id: int, target_file_path: str):
        super().__init__()
        self.record_id = record_id
        self.target_file_path = target_file_path

    def run(self):
        try:
            from app.services.export_service import ExportService
            service = ExportService()
            # 进度回调：发射 progress 信号，由 UI 线程更新浮层
            def on_progress(pct, msg):
                self.progress.emit(pct, msg)

            file_path = service.export_details(
                self.record_id, self.target_file_path, progress_callback=on_progress
            )
            self.finished.emit(file_path)
        except Exception as e:
            self.error.emit(str(e) + "\n" + traceback.format_exc())


class ExportSettlementWorker(QThread):
    """结算单导出后台线程"""
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, details, settlement_type, group_key, target_file_path, record_id=None):
        super().__init__()
        self.details = details
        self.settlement_type = settlement_type
        self.group_key = group_key
        self.target_file_path = target_file_path
        self.record_id = record_id

    def run(self):
        try:
            from openpyxl import Workbook

            type_names = {"station": "网点", "contract": "承包区", "monthly": "月结客户"}
            type_name = type_names.get(self.settlement_type, "结算")

            filtered = []
            for d in self.details:
                if self.settlement_type == "station" and d.station_code == self.group_key:
                    filtered.append(d)
                elif self.settlement_type == "contract":
                    sc = d.station_code or ""
                    if len(sc) >= 3 and sc[:3] == self.group_key:
                        filtered.append(d)
                elif self.settlement_type == "monthly":
                    original = {}
                    try:
                        if d.original_data:
                            original = d.original_data if isinstance(d.original_data, dict) \
                                else json.loads(d.original_data)
                    except Exception:
                        pass
                    cc = original.get("客户编码", original.get("客户代码", ""))
                    if str(cc).strip() == self.group_key:
                        filtered.append(d)

            wb = Workbook(write_only=True)
            ws = wb.create_sheet(type_name)
            ws.append(["行号", "快递单号", "网点编码", "网点名称", "区域",
                       "重量(kg)", "件数", "运费(元)", "应用规则", "是否异常", "备注"])
            for d in filtered:
                ws.append([
                    str(d.row_index), d.tracking_no or "", d.station_code or "",
                    d.station_name or "", d.region_name or "",
                    float(d.weight or 0), int(d.quantity or 0),
                    float(d.calculated_fee or 0), d.rule_name or "",
                    "是" if d.is_exception else "否", d.remark or "",
                ])
            wb.save(self.target_file_path)
            self.finished.emit(self.target_file_path)
        except Exception as e:
            self.error.emit(str(e) + "\n" + traceback.format_exc())


class ExportMultiWorker(QThread):
    """多文件导出后台线程"""
    finished = pyqtSignal(list)  # 成功：返回文件路径列表
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(self, record_ids: list, export_dir: str):
        super().__init__()
        self.record_ids = record_ids
        self.export_dir = export_dir

    def run(self):
        try:
            from app.services.export_service import ExportService
            service = ExportService()
            def on_progress(pct, msg):
                self.progress.emit(pct, msg)
            exported_files = service.export_multiple_records(
                self.record_ids, self.export_dir, progress_callback=on_progress
            )
            self.finished.emit(exported_files)
        except Exception as e:
            self.error.emit(str(e) + "\n" + traceback.format_exc())


class ExportMergedWorker(QThread):
    """合并导出后台线程"""
    finished = pyqtSignal(list)  # 成功：返回文件路径列表
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(self, record_ids: list, export_dir: str):
        super().__init__()
        self.record_ids = record_ids
        self.export_dir = export_dir

    def run(self):
        try:
            from app.services.export_service import ExportService
            service = ExportService()
            def on_progress(pct, msg):
                self.progress.emit(pct, msg)
            exported_files = service.export_merged_records(
                self.record_ids, self.export_dir, progress_callback=on_progress
            )
            self.finished.emit(exported_files)
        except Exception as e:
            self.error.emit(str(e) + "\n" + traceback.format_exc())
