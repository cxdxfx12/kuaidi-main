"""
派费计算服务 - 超高性能版（支持300万行级别）
- 规则字典化 O(1) 命中（客户编码/客户名称/区域关键词直接查字典）
- SQLite WAL + 200MB cache + 1GB mmap
- multiprocessing 多进程并行计算（子进程只负责计算，主进程统一入库）
- python-calamine (Rust引擎) 替代 openpyxl 读取Excel，速度提升5倍
"""
import os
import json
import math
from decimal import Decimal
from datetime import datetime
from typing import Dict, List, Tuple, Optional

from app.core.utils import parse_date

try:
    from python_calamine import CalamineWorkbook
    _HAS_CALAMINE = True
except Exception:
    _HAS_CALAMINE = False
    CalamineWorkbook = None

import pandas as pd

from app.services.excel_parser import ExcelParser
from app.services.rule_service import RuleService, apply_weight_rounding
from app.models.database import get_session, Base
from app.models.fee_record import FeeRecord
from app.models.fee_detail import FeeDetail

# 数据库表名常量
TABLE_FEE_DETAIL = "fee_details"

# 性能参数（调大有利于300万行数据）
BATCH_SIZE = 25000                  # 每批计算/入库行数（之前1万，改2.5万）
CHUNK_SIZE_FOR_MP = 200000          # 多进程时每进程负责的chunk大小（约20万行/进程）
MIN_ROWS_FOR_MULTIPROCESS = 100000  # 小于10万行时不启多进程（启动开销比收益大）


# ============================================
# 全局规则索引：初始化为空字典/列表，防止 None 调用 .get() 崩溃
# ============================================
# 客户级索引（客户名称/客户编码 - 最优先匹配）
_STATION_CODE_MAP = {}              # type: Dict[str, List[Tuple[...]]]
_STATION_NAME_MAP = {}              # type: Dict[str, List[Tuple[...]]]
# 网点级索引（网点名称/网点编码 - 次要匹配）
_OUTLET_CODE_MAP = {}               # type: Dict[str, List[Tuple[...]]]
_OUTLET_NAME_MAP = {}               # type: Dict[str, List[Tuple[...]]]
# 带区域限制的客户专属规则列表
_STATION_WITH_REGION_LIST = []      # type: List[Tuple[...]]
# 区域级规则
_REGION_MAP = {}                    # type: Dict[str, List[Tuple[...]]]
# 全局规则
_GLOBAL_RULES = []                  # type: List[Tuple[...]]
# 活动加价规则
_PROMOTION_RULES = []               # type: List[Dict]
# 活动加价规则预解析缓存（模块级，避免每行都解析日期/拆分字符串）
# 格式: list of (markup_type_int, markup_value_float, region_kws_tuple, promo_name_str, region_note_str)
# markup_type_int: 0=fixed, 1=weight, 2=percent
_PROMOTION_CACHE = []               # type: List[Tuple]
_EMPTY_WEIGHT_FEE = 3.0
_RULES_LOADED = False
# 计泡系数映射（station_code/name -> divisor），默认6000
_计泡系数_MAP = {}             # type: Dict[str, float]
_DEFAULT_计泡系数 = 6000
# 店铺→客户映射（一个客户多个店铺：store_code/store_name → customer_code）
_STORE_CUSTOMER_MAP = {}       # type: Dict[str, str]
_STORE_NAME_MAP = {}           # type: Dict[str, str]  # store_name → customer_code
_CUSTOMER_NAME_MAP = {}        # type: Dict[str, str]  # customer_code → customer_name

# ============================================
# 拉均重模式全局数据
# ============================================
# 拉均重规则索引：list of (customer_codes_set, region_kws_set, avg_weight_limit, 
#                         first_fee, continued_fee, min_fee, continued_unit, weight_rounding, rule_name)
_AVG_WEIGHT_RULES = []          # type: List[Tuple]
# 预计算的均重映射：{(customer_code, region_keyword): (avg_weight, deviation_step, deviation_surcharge)} 
# 模块级，供子进程读取
_AVG_WEIGHT_MAP = {}            # type: Dict[Tuple[str, str], Tuple[float, float, float]]


def _build_rule_indexes(force_reload: bool = False):
    """
    构建全局规则索引（只需要做1次）
    关键设计：规则的 stations 字段同时存入客户级索引和网点级索引，
    这样无论是客户名称"珀莱雅"还是网点名称"浙江杭州集包工厂"都能正确匹配。
    匹配优先级：客户名称 > 客户编码 > 网点名称 > 网点编码 > 区域 > 全局
    """
    global _STATION_CODE_MAP, _STATION_NAME_MAP, _OUTLET_CODE_MAP, _OUTLET_NAME_MAP
    global _STATION_WITH_REGION_LIST, _REGION_MAP, _GLOBAL_RULES, _PROMOTION_RULES
    global _PROMOTION_CACHE, _EMPTY_WEIGHT_FEE, _RULES_LOADED
    global _计泡系数_MAP, _DEFAULT_计泡系数
    import json as _json

    if _RULES_LOADED and not force_reload:
        return

    try:
        from app.services.rule_service import RuleService
        rs = RuleService()
        raw_rules = rs.load_rules()
        _EMPTY_WEIGHT_FEE = rs._load_empty_weight_fee()
        _PROMOTION_RULES = rs.load_promotion_rules()
    except Exception:
        _RULES_LOADED = True
        return

    # 从 default_settings.json 加载计泡系数（优先级1：用户全局默认设置）
    try:
        from app.models.path_config import get_config_file
        _default_path = get_config_file("default_settings.json")
        if os.path.exists(_default_path):
            with open(_default_path, "r", encoding="utf-8") as _f:
                _default_data = _json.load(_f)
            if "vol_divisor" in _default_data:
                _DEFAULT_计泡系数 = float(_default_data.get("vol_divisor", 6000))
    except Exception:
        pass

    # 从 fee_rules.json 加载计泡系数配置（优先级2）
    # 优先读用户数据目录（AppData），同时从打包资源补充缺失字段
    try:
        from app.models.path_config import get_config_file, get_resource_path

        # 从打包资源获取最新默认值（AppData 可能是旧版本）
        _resource_data = None
        _resource_default_vol = None
        _resource_path = get_resource_path("data", "config", "fee_rules.json")
        if os.path.exists(_resource_path):
            with open(_resource_path, "r", encoding="utf-8") as _f:
                _resource_data = _json.load(_f)
            _resource_default_vol = _resource_data.get("default_计泡系数")

        # 从用户数据目录读取（可能有用户自定义规则）
        _user_path = get_config_file("fee_rules.json")
        if os.path.exists(_user_path):
            with open(_user_path, "r", encoding="utf-8") as _f:
                _fee_data = _json.load(_f)
            # 从资源文件补充缺失字段（AppData 可能是旧版本）
            _need_write_back = False
            # 1. 补充 default_计泡系数
            if _fee_data.get("default_计泡系数") is None and _resource_default_vol is not None:
                _fee_data["default_计泡系数"] = _resource_default_vol
                _need_write_back = True
            # 2. 补充 计泡系数_MAP（网点→计泡系数映射，v2.2+ 新增字段）
            if "计泡系数_MAP" not in _fee_data and _resource_data and "计泡系数_MAP" in _resource_data:
                _fee_data["计泡系数_MAP"] = _resource_data["计泡系数_MAP"]
                _need_write_back = True
            # 3. 写回 AppData（下次不再重复补充）
            if _need_write_back:
                try:
                    with open(_user_path, "w", encoding="utf-8") as _f:
                        _json.dump(_fee_data, _f, ensure_ascii=False, indent=2)
                except Exception:
                    pass
        else:
            # AppData 没有，用打包资源的规则
            _fee_data = _resource_data

        # 设置全局默认值
        if _fee_data and _fee_data.get("default_计泡系数") is not None and _DEFAULT_计泡系数 == 6000:
            _DEFAULT_计泡系数 = float(_fee_data["default_计泡系数"])

        # 构建计泡系数 MAP
        # 1. 从计泡系数_MAP字段加载（网点名称→计泡系数映射，如 顺丰→6000）
        _计泡系数_MAP = {}
        if _fee_data:
            _map_from_field = _fee_data.get("计泡系数_MAP", {})
            for _k, _v in _map_from_field.items():
                _计泡系数_MAP[_k] = float(_v)
        # 2. 从rules里读取（每条规则专属的计泡系数，覆盖全局MAP）
        if _fee_data:
            for _rule in _fee_data.get("rules", []):
                _vd = float(_rule.get("计泡系数", _DEFAULT_计泡系数))
                _stations_str = _rule.get("stations", "")
                if _stations_str:
                    for _s in _stations_str.split(","):
                        _s = _s.strip()
                        if _s:
                            _计泡系数_MAP[_s] = _vd  # 规则内的值优先级更高
    except Exception:
        _计泡系数_MAP = {}

    global _AVG_WEIGHT_RULES, _AVG_WEIGHT_MAP

    _STATION_CODE_MAP = {}     # 客户编码索引
    _STATION_NAME_MAP = {}     # 客户名称索引
    _OUTLET_CODE_MAP = {}      # 网点编码索引
    _OUTLET_NAME_MAP = {}      # 网点名称索引
    _STATION_WITH_REGION_LIST = []
    _REGION_MAP = {}
    _GLOBAL_RULES = []
    _AVG_WEIGHT_RULES = []     # 拉均重规则索引
    _AVG_WEIGHT_MAP = {}       # 重置预计算结果

    for r in raw_rules:
        region_kw = []
        if r.regions and r.regions.strip():
            region_kw = [k.strip() for k in r.regions.split(",") if k.strip()]

        # stations字段可能包含：客户名称/客户编码/网点名称/网点编码，全部尝试匹配
        station_values = []
        if r.stations and r.stations.strip():
            for s in r.stations.split(","):
                s = s.strip()
                if s:
                    station_values.append(s)

        rule_core = (
            float(r.min_weight), float(r.max_weight),
            float(r.first_fee), float(r.continued_fee), float(r.min_fee),
            r.name,
            r.continued_unit or "kg",
            r.weight_rounding or "actual",
            getattr(r, 'pricing_mode', 'standard') or 'standard',
            float(getattr(r, 'tier_0_05', 0)),
            float(getattr(r, 'tier_05_1', 0)),
            float(getattr(r, 'tier_1_2', 0)),
            float(getattr(r, 'tier_2_3', 0)),
            float(getattr(r, 'first_fee_30', 0)),
            float(getattr(r, 'continued_fee_30', 0)),
        )

        # 收集拉均重规则（扩展支持阶梯定价 + 偏差加价）
        if getattr(r, 'avg_weight_mode', False) and station_values:
            avg_limit = float(getattr(r, 'avg_weight_limit', 3.0))
            pricing_mode_val = getattr(r, 'pricing_mode', 'standard') or 'standard'
            _AVG_WEIGHT_RULES.append((
                set(station_values),            # 0: 客户编码集合
                set(region_kw) if region_kw else set(),  # 1: 区域关键词集合
                avg_limit,                      # 2: 均重上限(kg)
                float(r.first_fee),             # 3: 首重费
                float(r.continued_fee),         # 4: 续重费
                float(r.min_fee),               # 5: 保底费
                r.continued_unit or "kg",       # 6: 续重单位
                r.weight_rounding or "actual",  # 7: 重量进位
                r.name,                         # 8: 规则名称
                pricing_mode_val,               # 9: 定价模式 (standard/tiered)
                float(getattr(r, 'tier_0_05', 0)),       # 10: 阶梯0-0.5kg
                float(getattr(r, 'tier_05_1', 0)),       # 11: 阶梯0.5-1kg
                float(getattr(r, 'tier_1_2', 0)),        # 12: 阶梯1-2kg
                float(getattr(r, 'tier_2_3', 0)),        # 13: 阶梯2-3kg
                float(getattr(r, 'first_fee_30', 0)),    # 14: 30kg+首重
                float(getattr(r, 'continued_fee_30', 0)),# 15: 30kg+续重
                float(getattr(r, 'avg_weight_deviation_step', 0.1)),   # 16: 偏差步长(kg)
                float(getattr(r, 'avg_weight_deviation_surcharge', 0.0)), # 17: 偏差加价(元)
            ))

        if r.rule_type == "station":
            if not region_kw:
                # 无区域限制的客户/网点规则：
                # stations字段的值同时存入客户级索引和网点级索引
                for v in station_values:
                    # 作为客户编码/客户名称匹配
                    _STATION_CODE_MAP.setdefault(v, []).append((*rule_core, []))
                    _STATION_NAME_MAP.setdefault(v, []).append((*rule_core, []))
                    # 作为网点编码/网点名称匹配
                    _OUTLET_CODE_MAP.setdefault(v, []).append((*rule_core, []))
                    _OUTLET_NAME_MAP.setdefault(v, []).append((*rule_core, []))
            else:
                # 有区域限制的客户专属规则
                _STATION_WITH_REGION_LIST.append(
                    (station_values, station_values, region_kw, *rule_core)
                )

        elif r.rule_type == "region":
            for kw in region_kw:
                _REGION_MAP.setdefault(kw, []).append(rule_core)

        elif r.rule_type == "global":
            _GLOBAL_RULES.append(rule_core)

    # 按最小重量排序，确保小重量优先命中
    for v in _STATION_CODE_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _STATION_NAME_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _OUTLET_CODE_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _OUTLET_NAME_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _REGION_MAP.values():
        v.sort(key=lambda x: x[0])
    _GLOBAL_RULES.sort(key=lambda x: x[0])

    # ========== 活动加价规则预解析 ==========
    # 把日期/区域/类型/值 等都预先解析好，运行期不再重复解析
    try:
        _PROMOTION_CACHE = []
        for pr in _PROMOTION_RULES:
            try:
                start = parse_date(pr.get("start_date", ""))
                end = parse_date(pr.get("end_date", ""))
                if start is None or end is None:
                    continue

                regions_val = str(pr.get("regions", "")).strip()
                # 统一中文逗号"，"→ 英文逗号","，避免用户输入混淆
                if regions_val:
                    regions_val = regions_val.replace("，", ",")
                region_kws = tuple(k.strip() for k in regions_val.split(",") if k.strip()) if regions_val else ()

                markup_type_str = str(pr.get("markup_type", "percent")).strip().lower()
                if markup_type_str == "fixed":
                    markup_type_int = 0
                elif markup_type_str == "weight":
                    markup_type_int = 1
                elif markup_type_str == "percent":
                    markup_type_int = 2
                else:
                    continue

                try:
                    markup_value = float(str(pr.get("markup_value", "0")).strip())
                except (ValueError, TypeError):
                    continue
                if markup_value <= 0:
                    continue

                promo_name = str(pr.get("name", "活动加价"))
                region_note = f"[{regions_val}]" if regions_val else ""

                _PROMOTION_CACHE.append(
                    (start, end, markup_type_int, markup_value, region_kws, promo_name, region_note)
                )
            except Exception:
                continue
    except Exception:
        pass

    _RULES_LOADED = True


def _calc_fee_from_rules(weight: float, rules_list: List[Tuple]) -> Optional[Tuple[float, str]]:
    """
    从规则列表中找到第一条重量匹配的规则，计算费用
    支持阶梯定价模式（pricing_mode == "tiered"）
    :return: (fee, rule_name) 或 None（没命中）
    """
    for rule in rules_list:
        min_w, max_w = rule[0], rule[1]
        if min_w <= weight <= max_w:
            first_f, continued_f, min_f = rule[2], rule[3], rule[4]
            continued_unit = rule[6] if len(rule) > 6 else "kg"
            weight_rounding = rule[7] if len(rule) > 7 else "actual"

            # 检查是否有阶梯定价数据（扩展 tuple 长度 >= 14）
            if len(rule) >= 14:
                pricing_mode = rule[8] if len(rule) > 8 else "standard"
                if pricing_mode == "tiered":
                    tier_0_05 = float(rule[9] if len(rule) > 9 else 0)
                    tier_05_1 = float(rule[10] if len(rule) > 10 else 0)
                    tier_1_2 = float(rule[11] if len(rule) > 11 else 0)
                    tier_2_3 = float(rule[12] if len(rule) > 12 else 0)
                    first_fee_30 = float(rule[13] if len(rule) > 13 else 0)
                    continued_fee_30 = float(rule[14] if len(rule) > 14 else 0)
                    # 阶梯定价不需要进位，直接用实际重量
                    from app.services.rule_service import calc_tiered_fee
                    fee = calc_tiered_fee(
                        weight, tier_0_05, tier_05_1, tier_1_2, tier_2_3,
                        first_f, continued_f, first_fee_30, continued_fee_30,
                        min_f, continued_unit
                    )
                    return fee, rule[5]

            # 标准模式：首重+续重
            rounded_weight = apply_weight_rounding(weight, weight_rounding, None)

            if rounded_weight <= 1.0:
                fee = first_f
            else:
                continued_weight = rounded_weight - 1.0
                if continued_unit == "100g":
                    units = math.ceil(continued_weight / 0.1)
                    fee = first_f + units * continued_f
                else:
                    fee = first_f + continued_weight * continued_f
            fee = round(max(fee, min_f), 2)
            return fee, rule[5]
    return None


def _apply_promotion(base_fee: float, weight: float, region_str: str = "",
                     business_date: Optional[datetime] = None) -> Tuple[float, str]:
    """统一应用活动加价规则（使用预解析缓存，避免每行重复解析日期/字符串）
    - 优先使用 Excel 中的业务日期（快递单日期）判断活动期间
    - 业务日期为空时，fallback 到当前日期
    """
    if not _PROMOTION_CACHE:
        return base_fee, ""

    try:
        if business_date is not None:
            check_date = business_date
        else:
            check_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        region_str = region_str or ""

        for entry in _PROMOTION_CACHE:
            # entry = (start, end, markup_type_int, markup_value, region_kws, promo_name, region_note)
            start = entry[0]
            end = entry[1]
            if not (start <= check_date <= end):
                continue

            # 区域限定检查：region_kws 是 tuple，空表示不限定
            region_kws = entry[4]
            if region_kws:
                if not any(k in region_str for k in region_kws):
                    continue

            markup_type_int = entry[2]
            markup_value = entry[3]

            if markup_type_int == 0:
                promo_amount = markup_value
            elif markup_type_int == 1:
                promo_amount = weight * markup_value
            elif markup_type_int == 2:
                promo_amount = base_fee * (markup_value / 100.0)
            else:
                continue

            promo_amount = round(promo_amount, 2)
            if promo_amount > 0:
                promo_name = entry[5]
                region_note = entry[6]
                return round(base_fee + promo_amount, 2), f"+ {promo_name}{region_note}(+¥{promo_amount})"
    except Exception:
        pass

    return base_fee, ""


def _read_excel_fast(file_path: str, sheet_name: Optional[str] = None,
                    use_columns: Optional[List[str]] = None) -> Tuple[List[str], List[List]]:
    """
    高性能读取Excel：优先 python-calamine(Rust引擎)，不可用时fallback到pandas+openpyxl
    :param use_columns: 指定只读取这些列名（精确匹配），不传则全读
    返回: (columns, data_rows)
      - columns: 列名列表
      - data_rows: 数据行 list[list]，每行保留原类型（None/int/float/str/date/datetime）
    """
    # ========== 方案A：calamine (Rust) ==========
    if _HAS_CALAMINE:
        try:
            wb = CalamineWorkbook.from_path(file_path)
            sheet_names_list = wb.sheet_names

            if sheet_name:
                target_sheets = [sheet_name] if sheet_name in sheet_names_list else [sheet_names_list[0]]
            else:
                target_sheets = sheet_names_list

            columns = None
            data_rows = []

            for idx, sn in enumerate(target_sheets):
                sheet = wb.get_sheet_by_name(sn)
                rows_iter = sheet.iter_rows()
                is_first_sheet = (idx == 0)

                for row_idx, row in enumerate(rows_iter):
                    if row is None:
                        continue

                    if row_idx == 0 and is_first_sheet:
                        columns = ["" if v is None else str(v).strip() for v in row]
                        continue

                    # 数据行：直接使用 tuple（calamine已生成），不做类型转换
                    # 后续 row_tup 构建时按需转换
                    data_rows.append(tuple(row))

            if columns is None:
                columns = []

            # 列过滤：只保留 use_columns 中指定的列
            if use_columns and columns:
                keep_indices = [columns.index(c) for c in use_columns if c in columns]
                if keep_indices:
                    columns = [columns[i] for i in keep_indices]
                    data_rows = [[row[i] for i in keep_indices] for row in data_rows]

            return columns, data_rows
        except Exception:
            pass  # fallback

    # ========== 方案B：pandas + openpyxl (兼容兜底) ==========
    # usecols 让 pandas 在读取时就跳过无关列
    usecols_param = use_columns if use_columns else None
    safe_usecols = None
    if usecols_param:
        try:
            # 先读前几行获取列名，确定 usecols 的有效性
            pd_cols = pd.read_excel(file_path, nrows=0, sheet_name=sheet_name).columns.tolist()
            safe_usecols = [c for c in usecols_param if c in pd_cols]
            if not safe_usecols:
                safe_usecols = None
        except Exception:
            safe_usecols = None
    df = pd.read_excel(file_path, dtype=str, sheet_name=sheet_name, usecols=safe_usecols)
    if isinstance(df, dict):  # 多sheet合并
        columns = []
        data_rows = []
        is_first_sheet = True
        for sheet_df in df.values():
            sheet_df = sheet_df.fillna("")
            sheet_cols = list(sheet_df.columns)
            if is_first_sheet:
                columns = sheet_cols
                is_first_sheet = False
            for row in sheet_df.itertuples(index=False, name=None):
                data_rows.append(tuple(row))
        return columns, data_rows
    else:
        df = df.fillna("")
        columns = list(df.columns)
        data_rows = [tuple(row) for row in df.itertuples(index=False, name=None)]
        return columns, data_rows


def match_rule_fast(weight: float, region: str, station_code: str = "", station_name: str = "",
                     customer_code: str = "", customer_name: str = "",
                     business_date: Optional[datetime] = None) -> Tuple[float, str, bool]:
    """
    规则匹配核心函数（模块级，方便多进程调用）
    匹配优先级（从高到低）：
      1. 客户名称（customer_name）- 如"珀莱雅"
      2. 客户编码（customer_code）
      3. 网点名称（station_name）- 如"浙江杭州集包工厂"
      4. 网点编码（station_code）
      5. 有区域限制的客户/网点专属规则
      6. 区域级规则
      7. 全局兜底规则
    返回: (fee, rule_name, is_exception)
    """
    if not _RULES_LOADED:
        _build_rule_indexes(force_reload=True)

    if weight is None or weight <= 0:
        base_fee = _EMPTY_WEIGHT_FEE
        rule_name = "无重量默认价"
        final_fee, promo_suffix = _apply_promotion(base_fee, weight, region or "", business_date)
        if promo_suffix:
            return final_fee, f"{rule_name} {promo_suffix}", False
        return final_fee, rule_name, False

    region_str = region or ""
    cust_name_str = customer_name or ""
    cust_code_str = customer_code or ""
    outlet_name_str = station_name or ""
    outlet_code_str = station_code or ""

    base_fee = 0.0
    rule_name = ""
    matched = False

    # ==================== 路径1：优先匹配 客户名称（最高优先级） ====================
    if not matched and cust_name_str:
        # 精确匹配客户名称
        cust_rules = _STATION_NAME_MAP.get(cust_name_str)
        if cust_rules:
            result = _calc_fee_from_rules(weight, cust_rules)
            if result is not None:
                base_fee, rule_name = result
                matched = True

    # ==================== 路径2：匹配 客户编码 ====================
    if not matched and cust_code_str:
        cust_rules = _STATION_CODE_MAP.get(cust_code_str)
        if cust_rules:
            result = _calc_fee_from_rules(weight, cust_rules)
            if result is not None:
                base_fee, rule_name = result
                matched = True

    # ==================== 路径3：匹配 网点名称 ====================
    if not matched and outlet_name_str:
        outlet_rules = _OUTLET_NAME_MAP.get(outlet_name_str)
        if outlet_rules:
            result = _calc_fee_from_rules(weight, outlet_rules)
            if result is not None:
                base_fee, rule_name = result
                matched = True

    # ==================== 路径4：匹配 网点编码 ====================
    if not matched and outlet_code_str:
        outlet_rules = _OUTLET_CODE_MAP.get(outlet_code_str)
        if outlet_rules:
            result = _calc_fee_from_rules(weight, outlet_rules)
            if result is not None:
                base_fee, rule_name = result
                matched = True

    # ==================== 路径5：带区域限制的客户/网点专属规则 ====================
    if not matched and _STATION_WITH_REGION_LIST:
        for item in _STATION_WITH_REGION_LIST:
            codes = item[0]
            names = item[1]
            region_kws = item[2]
            min_w = item[3]
            max_w = item[4]
            first_f = item[5]
            continued_f = item[6]
            min_f = item[7]
            r_name = item[8]
            continued_unit = item[9] if len(item) > 9 else "kg"
            weight_rounding = item[10] if len(item) > 10 else "actual"

            # 尝试匹配：客户名称 / 客户编码 / 网点名称 / 网点编码
            name_match = (cust_name_str and cust_name_str in names) or \
                         (cust_code_str and cust_code_str in codes) or \
                         (outlet_name_str and outlet_name_str in names) or \
                         (outlet_code_str and outlet_code_str in codes)
            if not name_match:
                continue
            if region_kws and not any(k in region_str for k in region_kws):
                continue
            if min_w <= weight <= max_w:
                # 检查阶梯定价模式（item[11] = pricing_mode）
                pricing_mode = item[11] if len(item) > 11 else "standard"
                if pricing_mode == "tiered" and len(item) >= 17:
                    tier_0_05 = float(item[12])
                    tier_05_1 = float(item[13])
                    tier_1_2 = float(item[14])
                    tier_2_3 = float(item[15])
                    first_fee_30 = float(item[16])
                    continued_fee_30 = float(item[17] if len(item) > 17 else 0)
                    from app.services.rule_service import calc_tiered_fee
                    fee = calc_tiered_fee(
                        weight, tier_0_05, tier_05_1, tier_1_2, tier_2_3,
                        first_f, continued_f, first_fee_30, continued_fee_30,
                        min_f, continued_unit
                    )
                    base_fee = fee
                else:
                    rounded_weight = apply_weight_rounding(weight, weight_rounding, None)
                    if rounded_weight <= 1.0:
                        fee = first_f
                    else:
                        continued_weight = rounded_weight - 1.0
                        if continued_unit == "100g":
                            units = math.ceil(continued_weight / 0.1)
                            fee = first_f + units * continued_f
                        else:
                            fee = first_f + continued_weight * continued_f
                    base_fee = round(max(fee, min_f), 2)
                rule_name = r_name
                matched = True
                break

    # ==================== 路径6：区域级规则 ====================
    if not matched and region_str:
        for kw in _REGION_MAP.keys():
            if kw in region_str:
                result = _calc_fee_from_rules(weight, _REGION_MAP[kw])
                if result is not None:
                    base_fee, rule_name = result
                    matched = True
                    break

    # ==================== 路径7：全局兜底规则 ====================
    if not matched and _GLOBAL_RULES:
        result = _calc_fee_from_rules(weight, _GLOBAL_RULES)
        if result is not None:
            base_fee, rule_name = result
            matched = True

    # ==================== 没有任何规则命中，返回异常 ====================
    if not matched:
        return 0.0, "无匹配规则", True

    # 最后统一应用活动加价
    final_fee, promo_suffix = _apply_promotion(base_fee, weight, region_str, business_date)
    if promo_suffix:
        return final_fee, f"{rule_name} {promo_suffix}", False
    return final_fee, rule_name, False


def _load_store_customer_map(force_reload: bool = False):
    """加载店铺→客户映射（一个客户多个店铺），从数据库 customer_stores 表读取"""
    global _STORE_CUSTOMER_MAP, _STORE_NAME_MAP, _CUSTOMER_NAME_MAP
    if _STORE_CUSTOMER_MAP and not force_reload:
        return
    try:
        from app.models.database import get_session
        from app.models.customer_store import CustomerStore
        session = get_session()
        try:
            rows = session.query(CustomerStore).all()
            _STORE_CUSTOMER_MAP = {}
            _STORE_NAME_MAP = {}
            _CUSTOMER_NAME_MAP = {}
            for r in rows:
                _STORE_CUSTOMER_MAP[r.store_code] = r.customer_code
                if r.store_name:
                    _STORE_NAME_MAP[r.store_name] = r.customer_code
                if r.customer_code not in _CUSTOMER_NAME_MAP:
                    _CUSTOMER_NAME_MAP[r.customer_code] = r.customer_name
        finally:
            session.close()
    except Exception:
        pass


# ============================================
# 辅助函数：根据快递公司/网点名称模糊匹配计泡系数
# ============================================
def _find_计泡系数(station_code: str, station_name: str,
                 courier_code: str = "", courier_name: str = "") -> float:
    """
    模糊匹配计泡系数：遍历计泡系数_MAP，查找 key 是否出现在相关字段中
    匹配优先级（从高到低）：
      1. courier_code 精确匹配（快递公司编码，如 SF、YTO）
      2. courier_name 模糊匹配（快递公司名称，如 顺丰、圆通速递）
      3. station_code 精确匹配（网点编码，如 SF001）
      4. station_name 模糊匹配（网点名称，如 圆通速递杭州分部）
    - 返回匹配到的计泡系数，否则返回默认的 _DEFAULT_计泡系数
    """
    # ---------- 级别1：快递公司编码精确匹配 ----------
    if courier_code and courier_code in _计泡系数_MAP:
        return _计泡系数_MAP[courier_code]

    # ---------- 级别2：快递公司名称模糊匹配 ----------
    if courier_name:
        for key in sorted(_计泡系数_MAP.keys(), key=len, reverse=True):
            if key in courier_name:
                return _计泡系数_MAP[key]

    # ---------- 级别3：网点编码精确匹配 ----------
    if station_code and station_code in _计泡系数_MAP:
        return _计泡系数_MAP[station_code]

    # ---------- 级别4：网点名称模糊匹配 ----------
    if station_name:
        for key in sorted(_计泡系数_MAP.keys(), key=len, reverse=True):
            if key in station_name:
                return _计泡系数_MAP[key]

    return _DEFAULT_计泡系数


# ============================================
# 拉均重预计算：扫描所有行，按客户+区域分组计算均重
# ============================================
def _precompute_avg_weights(all_row_tuples: List, column_indices: List[int]):
    """
    在正式计算之前，预计算所有拉均重分组的平均重量。
    将结果存入模块全局变量 _AVG_WEIGHT_MAP，供 _process_chunk 使用。
    
    Key: (customer_code, region_keyword) 
    Value: (avg_weight, rule_tuple)
    """
    global _AVG_WEIGHT_MAP
    _AVG_WEIGHT_MAP = {}

    if not _AVG_WEIGHT_RULES or not all_row_tuples:
        return

    idx_customer_code = column_indices[15] if len(column_indices) > 15 else -1
    idx_order_customer = column_indices[17] if len(column_indices) > 17 else -1
    idx_region_name = column_indices[6] if len(column_indices) > 6 else -1
    idx_weight = column_indices[7] if len(column_indices) > 7 else -1

    if idx_customer_code < 0 or idx_weight < 0:
        return

    # 确保店铺→客户映射已加载
    _load_store_customer_map()

    # 第一步：为每个拉均重规则建立 (customer_code, region_keyword) → [weights] 的映射
    # 使用 rule_index + group_key 的方式
    group_weights = {}  # {(customer_code, region_kw, rule_idx): [weights]}

    for row_vals in all_row_tuples:
        raw_customer_code = str(row_vals[idx_customer_code] or "").strip()
        raw_order_customer = str(row_vals[idx_order_customer] or "").strip() if idx_order_customer >= 0 else ""
        region_str = str(row_vals[idx_region_name] or "").strip() if idx_region_name >= 0 else ""
        weight_str = str(row_vals[idx_weight] or "").strip()

        if not raw_customer_code and not raw_order_customer:
            continue

        try:
            weight = float(weight_str) if weight_str else 0.0
        except (ValueError, TypeError):
            continue
        if weight <= 0:
            continue

        # 店铺→客户映射
        resolved_code = raw_customer_code
        if _STORE_CUSTOMER_MAP:
            if raw_customer_code in _STORE_CUSTOMER_MAP:
                resolved_code = _STORE_CUSTOMER_MAP[raw_customer_code]
            elif raw_order_customer and raw_order_customer in _STORE_NAME_MAP:
                resolved_code = _STORE_NAME_MAP[raw_order_customer]
            elif raw_customer_code in _STORE_NAME_MAP:
                resolved_code = _STORE_NAME_MAP[raw_customer_code]

        # 检查是否匹配任何拉均重规则
        for ri, rule_info in enumerate(_AVG_WEIGHT_RULES):
            customer_codes = rule_info[0]      # set of strings
            region_kws = rule_info[1]           # set of strings
            avg_limit = rule_info[2]            # float

            # 客户匹配
            if resolved_code not in customer_codes and raw_customer_code not in customer_codes:
                continue

            # 区域匹配：rule.regions 关键词在 row.region 中
            if region_kws and region_str:
                matched_kw = None
                for kw in region_kws:
                    if kw in region_str:
                        matched_kw = kw
                        break
                if matched_kw is None:
                    continue
            elif region_kws and not region_str:
                continue
            else:
                matched_kw = ""  # 无区域限制

            # 重量在均重范围内（< limit）
            if weight >= avg_limit:
                continue

            group_key = (resolved_code, matched_kw, ri)
            if group_key not in group_weights:
                group_weights[group_key] = []
            group_weights[group_key].append(weight)

    # 第二步：计算每组均重（支持阶梯定价模式）
    for (customer_code, region_kw, ri), weights in group_weights.items():
        if not weights:
            continue
        avg_weight = sum(weights) / len(weights)
        avg_weight = round(avg_weight, 3)  # 保留3位小数
        rule_info = _AVG_WEIGHT_RULES[ri]

        # 提取规则参数（兼容新老格式）
        first_fee = rule_info[3]
        continued_fee = rule_info[4]
        min_fee = rule_info[5]
        continued_unit = rule_info[6]
        weight_rounding = rule_info[7]
        rule_name = rule_info[8]
        pricing_mode = rule_info[9] if len(rule_info) > 9 else "standard"
        tier_0_05 = float(rule_info[10]) if len(rule_info) > 10 else 0.0
        tier_05_1 = float(rule_info[11]) if len(rule_info) > 11 else 0.0
        tier_1_2 = float(rule_info[12]) if len(rule_info) > 12 else 0.0
        tier_2_3 = float(rule_info[13]) if len(rule_info) > 13 else 0.0
        first_fee_30 = float(rule_info[14]) if len(rule_info) > 14 else 0.0
        continued_fee_30 = float(rule_info[15]) if len(rule_info) > 15 else 0.0
        deviation_step = float(rule_info[16]) if len(rule_info) > 16 else 0.1
        deviation_surcharge = float(rule_info[17]) if len(rule_info) > 17 else 0.0

        # 按定价模式计算均重对应的每票费用
        if pricing_mode == "tiered":
            from app.services.rule_service import calc_tiered_fee
            fee_per_ticket = calc_tiered_fee(
                avg_weight, tier_0_05, tier_05_1, tier_1_2, tier_2_3,
                first_fee, continued_fee, first_fee_30, continued_fee_30,
                min_fee, continued_unit
            )
        else:
            rounded_weight = apply_weight_rounding(avg_weight, weight_rounding, None)
            if rounded_weight <= 1.0:
                fee_per_ticket = first_fee
            else:
                continued_weight = rounded_weight - 1.0
                if continued_unit == "100g":
                    units = math.ceil(continued_weight / 0.1)
                    fee_per_ticket = first_fee + units * continued_fee
                else:
                    fee_per_ticket = first_fee + continued_weight * continued_fee
            fee_per_ticket = round(max(fee_per_ticket, min_fee), 2)

        map_key = (customer_code, region_kw)
        _AVG_WEIGHT_MAP[map_key] = (avg_weight, deviation_step, deviation_surcharge)


# ============================================
# 多进程 worker：只负责计算（纯函数，无IO，无GUI）
# ============================================
def _process_chunk(args):
    """
    子进程工作函数：接收一块数据（行列表），计算后返回结果列表
    规则索引在子进程启动时通过 initializer 重建一次
    去重由主进程在分chunk前完成
    """
    chunk_rows, idx_list = args

    try:
        # 恢复拉均重预计算数据
        global _AVG_WEIGHT_MAP, _AVG_WEIGHT_RULES
        _saved_avg_weight_map = dict(_AVG_WEIGHT_MAP) if _AVG_WEIGHT_MAP else {}
        _saved_avg_weight_rules = list(_AVG_WEIGHT_RULES)

        # 确保规则索引已构建（子进程第一次调用时需要）
        if not _RULES_LOADED:
            _build_rule_indexes(force_reload=True)
        # 确保店铺→客户映射已加载
        _load_store_customer_map()

        # 恢复拉均重预计算数据
        if _saved_avg_weight_map:
            _AVG_WEIGHT_MAP.update(_saved_avg_weight_map)
        if _saved_avg_weight_rules and not _AVG_WEIGHT_RULES:
            _AVG_WEIGHT_RULES.extend(_saved_avg_weight_rules)

        results = []

        for row_vals in chunk_rows:
            tracking_no = row_vals[0] or ""
            station_code = row_vals[1] or ""
            station_name = row_vals[2] or ""
            courier_code = row_vals[3] or ""
            courier_name = row_vals[4] or ""
            region_code = row_vals[5] or ""
            region_name = row_vals[6] or ""
            weight_str = row_vals[7] or ""
            length_str = row_vals[8] or ""
            width_str = row_vals[9] or ""
            height_str = row_vals[10] or ""
            vol_weight_str = row_vals[11] or ""
            quantity_str = row_vals[12] or ""
            service_type = row_vals[13] or ""
            raw_date = row_vals[14] or ""
            raw_customer_code = row_vals[15] or ""
            raw_customer = row_vals[16] or ""
            raw_order_customer = row_vals[17] or ""  # 订单客户（面单上的店铺名称）
            remark = row_vals[18] or ""
            excel_row_index = row_vals[19]

            # ==== 店铺→客户映射（一个客户多个店铺的核心逻辑）====
            # 将店铺编码/名称解析为父客户编码，保证使用客户规则计算
            resolved_customer_code = raw_customer_code
            if _STORE_CUSTOMER_MAP:
                # 优先按店铺编码匹配
                if raw_customer_code and raw_customer_code in _STORE_CUSTOMER_MAP:
                    resolved_customer_code = _STORE_CUSTOMER_MAP[raw_customer_code]
                # 其次按"订单客户"（店铺名称）匹配
                elif raw_order_customer and raw_order_customer in _STORE_NAME_MAP:
                    resolved_customer_code = _STORE_NAME_MAP[raw_order_customer]
                # 再用原始客户编码尝试店铺名称匹配
                elif raw_customer_code and raw_customer_code in _STORE_NAME_MAP:
                    resolved_customer_code = _STORE_NAME_MAP[raw_customer_code]
                if resolved_customer_code != raw_customer_code:
                    # 同时用父客户名称替换（如果映射中有）
                    parent_name = _CUSTOMER_NAME_MAP.get(resolved_customer_code, "")
                    if parent_name and not raw_customer:
                        raw_customer = parent_name
            # ====================================================

            try:
                weight = float(weight_str) if weight_str else 0.0
            except (ValueError, TypeError):
                weight = 0.0

            # ========== 体积重计算 ==========
            # 体积重 = 长 × 宽 × 高 ÷ 抛货系数
            # 计费重量 = max(实重, 体积重)
            billing_weight = weight
            try:
                length = float(length_str) if length_str else 0.0
                width = float(width_str) if width_str else 0.0
                height = float(height_str) if height_str else 0.0
                if length > 0 and width > 0 and height > 0:
                    # 优先用文件中预存的体积重，否则自己计算
                    if vol_weight_str:
                        vol_weight = float(vol_weight_str)
                    else:
                        # 通过模糊匹配查找计泡系数
                        divisor = _find_计泡系数(station_code, station_name,
                                               courier_code, courier_name)
                        vol_weight = (length * width * height) / divisor
                    # 取较大值为计费重量
                    if vol_weight > billing_weight:
                        billing_weight = vol_weight
            except (ValueError, TypeError, ZeroDivisionError):
                pass

            try:
                quantity = int(float(quantity_str)) if quantity_str else 1
            except (ValueError, TypeError):
                quantity = 1

            # 解析业务日期（用于活动加价期间判断）
            biz_date = parse_date(raw_date) if raw_date else None

            # ========== 拉均重模式：超重包裹用实际重量，均重内包裹用组均重 ==========
            actual_weight = billing_weight  # 保留实际计费重量（用于备注）
            avg_weight_deviation_info = None  # (deviation_step, deviation_surcharge)，用于偏差加价
            if _AVG_WEIGHT_MAP and billing_weight > 0 and resolved_customer_code:
                for ri, rule_info in enumerate(_AVG_WEIGHT_RULES):
                    # 客户匹配
                    if resolved_customer_code not in rule_info[0]:
                        # 也尝试用原始编码匹配
                        if not raw_customer_code or raw_customer_code not in rule_info[0]:
                            continue
                    region_kws = rule_info[1]
                    limit = rule_info[2]
                    # 重量超过上限 → 不参与均重，按实际重量计算
                    if billing_weight >= limit:
                        continue
                    # 区域匹配
                    if region_kws:
                        matched_kw = None
                        for kw in region_kws:
                            if kw in (region_name or ""):
                                matched_kw = kw
                                break
                        if matched_kw is None:
                            continue
                    else:
                        matched_kw = ""
                    # 查找预计算均重
                    search_code = resolved_customer_code
                    # 也要尝试用原始编码查找（兼容映射前的编码）
                    if (search_code, matched_kw) not in _AVG_WEIGHT_MAP and raw_customer_code:
                        search_code = raw_customer_code
                    map_key = (search_code, matched_kw)
                    if map_key in _AVG_WEIGHT_MAP:
                        avg_weight_val, dev_step, dev_surcharge = _AVG_WEIGHT_MAP[map_key]
                        billing_weight = avg_weight_val  # 用组均重替换计费重量
                        avg_weight_deviation_info = (dev_step, dev_surcharge)
                        break

            fee, rule_name, is_exc = match_rule_fast(billing_weight, region_name,
                                                      station_code, station_name,
                                                      resolved_customer_code, raw_customer,
                                                      biz_date)

            # ========== 均重偏差加价：实际重量超过组均重时，按步长加价 ==========
            if not is_exc and avg_weight_deviation_info is not None:
                dev_step, dev_surcharge = avg_weight_deviation_info
                if dev_surcharge > 0 and dev_step > 0 and actual_weight > billing_weight:
                    excess = actual_weight - billing_weight
                    # 使用 round 消除浮点精度问题（如 0.4-0.3=0.10000000000000003）
                    steps = math.ceil(round(excess / dev_step, 9))
                    surcharge = round(steps * dev_surcharge, 2)
                    if surcharge > 0:
                        fee = round(fee + surcharge, 2)
                        rule_name = f"{rule_name} [偏差+¥{surcharge}]"

            extra_data = None
            if raw_date or raw_customer_code or raw_customer or raw_order_customer:
                extra_data = json.dumps({
                    "business_date": raw_date,
                    "customer_code": raw_customer_code,
                    "customer_name": raw_customer,
                    "order_customer": raw_order_customer
                }, ensure_ascii=False)

            results.append((
                excel_row_index, tracking_no, station_code, station_name,
                courier_code, courier_name, region_code, region_name,
                billing_weight, quantity, service_type, extra_data,
                fee, rule_name, 1 if is_exc else 0,
                "invalid_data" if is_exc else None,
                remark or (f"无效重量:{weight_str}" if is_exc and weight <= 0 else "")
            ))

        return results
    except Exception:
        # 单个 chunk 处理失败时返回空列表，不影响其他 chunk
        return []


class CalculateService:
    """计算服务 - 支持300万行大数据量"""

    def __init__(self):
        self.parser = ExcelParser()
        self.rule_service = RuleService()
        # 每次新建 CalculateService 都强制重建索引，确保获取最新规则（包括活动加价规则）
        _build_rule_indexes(force_reload=True)
        self._empty_weight_fee = _EMPTY_WEIGHT_FEE

    def _load_empty_weight_fee(self) -> float:
        """保留兼容（实际值来自 _build_rule_indexes）"""
        return self._empty_weight_fee

    def _match_rule(self, weight: float, region: str, station_code: str = "", station_name: str = "") -> Tuple[float, str, bool]:
        """兼容方法：底层调用模块级的 match_rule_fast"""
        return match_rule_fast(weight, region, station_code, station_name)

    def _load_rules_list(self) -> List[Tuple]:
        """保留兼容（实际上规则已字典化，此处返回空列表即可，仅占位）"""
        return []

    def import_and_calculate(self, file_path: str, sheet_name=None,
                             progress_callback=None) -> Dict:
        """
        一键导入并计算（支持300万行超大数据量）
        - 单进程模式：< 10万行
        - 多进程模式：>= 10万行，自动启动 multiprocessing 并行计算
        - 支持多Sheet：如果未指定sheet_name，自动合并所有Sheet的数据
        """
        def report(percent, stage):
            if progress_callback:
                try:
                    progress_callback(int(percent), str(stage))
                except Exception:
                    pass

        # ============ 阶段0：验证规则索引 ============
        # 确保规则索引已正确加载（对于多进程也很重要：验证主进程能正确读取规则）
        _build_rule_indexes(force_reload=True)
        customer_rule_count = sum(len(v) for v in _STATION_NAME_MAP.values()) + \
                              sum(len(v) for v in _STATION_CODE_MAP.values())
        outlet_rule_count = sum(len(v) for v in _OUTLET_NAME_MAP.values()) + \
                            sum(len(v) for v in _OUTLET_CODE_MAP.values())
        report(1, f"规则校验完成（客户级规则 {customer_rule_count} 条，"
                f"网点级规则 {outlet_rule_count} 条，"
                f"区域规则 {sum(len(v) for v in _REGION_MAP.values())} 条，"
                f"全局规则 {len(_GLOBAL_RULES)} 条，"
                f"活动规则 {len(_PROMOTION_RULES)} 条）")

        # ============ 阶段1：读取Excel（优先Rust引擎，省掉pandas中间层） ============
        # 精确列名列表：只读取这8列，大幅减少内存和后续处理开销
        _EXACT_COL_NAMES = ["业务时间", "运单号", "结算重量", "目的省份",
                            "体积重", "订单/面单网点", "订单客户", "客户"]
        if file_path.endswith(".csv"):
            report(3, "正在读取CSV...")
            df = pd.read_csv(file_path, dtype=str, usecols=lambda c: c in _EXACT_COL_NAMES)
            columns = list(df.columns)
            data_rows = []
            for row in df.itertuples(index=False, name=None):
                data_rows.append(["" if v is None else str(v) for v in row])
            del df
            import gc as _gc0
            _gc0.collect()
        else:
            report(5, "正在读取Excel (Rust引擎)...")
            columns, data_rows = _read_excel_fast(file_path, sheet_name,
                                                  use_columns=_EXACT_COL_NAMES)
            report(15, f"读取完成，共 {len(data_rows):,} 行")

        row_count = len(data_rows)

        # 精确列名映射：只匹配这8列的精确名称，不做模糊匹配
        # 实际列名 → 标准字段名
        _PRECISE_COLUMNS = [
            ("业务时间", "business_date"),
            ("运单号", "tracking_no"),
            ("结算重量", "weight"),
            ("目的省份", "region_name"),
            ("体积重", "volume_weight"),
            ("订单/面单网点", "station_name"),
            ("订单客户", "order_customer"),
            ("客户", "customer_name"),
        ]
        col_map = {}
        unmatched = []
        for actual_name, std_name in _PRECISE_COLUMNS:
            if actual_name in columns:
                col_map[std_name] = actual_name
            else:
                unmatched.append(actual_name)

        if unmatched:
            report(15, f"注意：以下列未在文件中找到：{', '.join(unmatched)}")

        report(16, f"文件读取完成，共 {row_count:,} 行数据，正在初始化...")

        # 预计算列索引（基于已修剪的 columns）
        def _get_col_idx(std_name: str) -> int:
            actual = col_map.get(std_name)
            if actual and actual in columns:
                return columns.index(actual)
            return -1

        idx_tracking = _get_col_idx("tracking_no")
        idx_station_code = _get_col_idx("station_code")
        idx_station_name = _get_col_idx("station_name")
        idx_courier_code = _get_col_idx("courier_code")
        idx_courier_name = _get_col_idx("courier_name")
        idx_region_code = _get_col_idx("region_code")
        idx_region_name = _get_col_idx("region_name")
        idx_weight = _get_col_idx("weight")
        idx_length = _get_col_idx("length")
        idx_width = _get_col_idx("width")
        idx_height = _get_col_idx("height")
        idx_volume_weight = _get_col_idx("volume_weight")
        idx_quantity = _get_col_idx("quantity")
        idx_service = _get_col_idx("service_type")
        idx_date = _get_col_idx("business_date")
        idx_customer_code = _get_col_idx("customer_code")
        idx_customer = _get_col_idx("customer_name")
        idx_order_customer = _get_col_idx("order_customer")
        idx_remark = _get_col_idx("remark")

        # 把列索引按固定顺序打包（传给多进程 worker 用）
        column_indices = [
            idx_tracking, idx_station_code, idx_station_name,
            idx_courier_code, idx_courier_name, idx_region_code, idx_region_name,
            idx_weight, idx_length, idx_width, idx_height, idx_volume_weight,
            idx_quantity, idx_service,
            idx_date, idx_customer_code, idx_customer, idx_order_customer, idx_remark
        ]

        # ============ 阶段2：创建任务记录 ============
        record = FeeRecord(
            file_name=os.path.basename(file_path),
            file_path=file_path,
            file_size=os.path.getsize(file_path),
            total_rows=row_count,
            status="processing"
        )

        session = get_session()
        try:
            session.add(record)
            session.commit()
            record_id = record.id

            # ============ 阶段3：计算并入库 ============
            # 关键修复：先删除该record_id可能存在的旧数据（防止程序崩溃重启后重复）
            try:
                conn_pre = session.connection().connection
                cur_pre = conn_pre.cursor()
                cur_pre.execute(f"DELETE FROM {TABLE_FEE_DETAIL} WHERE record_id = ?", (record_id,))
                conn_pre.commit()
                cur_pre.close()
            except Exception:
                pass

            # ============ 阶段3-前置：SQLite性能调优（对百万级数据关键） ============
            # WAL模式：写入并发 + 写速度提升
            # synchronous=OFF：允许OS缓存写入，崩溃可能丢最近一批数据，但程序崩溃极少见
            # cache_size=-800000：分配800MB页面缓存（260万行×~60字节≈156MB，够存）
            # temp_store=MEMORY：临时表放内存不写盘
            try:
                perf_conn = session.connection().connection
                perf_cur = perf_conn.cursor()
                perf_cur.execute("PRAGMA journal_mode = WAL")
                perf_cur.execute("PRAGMA synchronous = OFF")
                perf_cur.execute("PRAGMA cache_size = -800000")
                perf_cur.execute("PRAGMA temp_store = MEMORY")
                perf_cur.execute("PRAGMA mmap_size = 2000000000")  # 2GB内存映射
                perf_conn.commit()
                perf_cur.close()
            except Exception:
                pass

            # 累计写入行数（用于控制大事务提交频率）
            rows_since_commit = 0
            # 每 500K 行 commit 一次：避免事务过大，又把2.5万行一次的commit从104次降到5次
            COMMIT_EVERY_ROWS = 500000

            def _lazy_commit_if_needed(current_conn, rows_count, force_now=False):
                nonlocal rows_since_commit
                rows_since_commit += rows_count
                if force_now or rows_since_commit >= COMMIT_EVERY_ROWS:
                    current_conn.commit()
                    rows_since_commit = 0

            # 决策：10万行以下走单进程（省启动开销），10万行以上自动多进程
            use_multiprocess = row_count >= MIN_ROWS_FOR_MULTIPROCESS

            # ============ 阶段3-a：直接从 data_rows 构建行数据 + 全局去重 ============
            # data_rows 已经是 list[list[str]]，用 column_indices 直接取对应列
            all_row_tuples = []
            seen_global = set()
            duplicate_count = 0
            empty_tracking_count = 0

            import gc
            n_cols = len(columns)
            for pos in range(row_count):
                excel_row = pos + 2
                raw = data_rows[pos]
                row_tup = []
                for i in column_indices:
                    if 0 <= i < len(raw):
                        v = raw[i]
                        if v is None:
                            row_tup.append("")
                        elif isinstance(v, float) and v != v:  # NaN
                            row_tup.append("")
                        else:
                            row_tup.append(str(v))
                    else:
                        row_tup.append("")
                row_tup.append(excel_row)

                tracking_no = row_tup[0] or ""
                if not tracking_no:
                    empty_tracking_count += 1
                    all_row_tuples.append(row_tup)
                    continue
                if tracking_no in seen_global:
                    duplicate_count += 1
                    continue
                seen_global.add(tracking_no)
                all_row_tuples.append(row_tup)

            # 更新去重后的行数
            row_count_original = row_count
            row_count = len(all_row_tuples)

            report(19,
                   f"数据构建完成。原始 {row_count_original:,} 行，"
                   f"去重后 {row_count:,} 行 "
                   f"（跳过重复 {duplicate_count:,} 行，空单号 {empty_tracking_count:,} 行）")

            # 释放 data_rows 内存
            del data_rows

            # ========== 拉均重预计算（在正式计算前，扫描所有行计算组均重） ==========
            if _AVG_WEIGHT_RULES:
                report(18, f"检测到 {len(_AVG_WEIGHT_RULES)} 条拉均重规则，正在预计算组均重...")
                _precompute_avg_weights(all_row_tuples, column_indices)
                if _AVG_WEIGHT_MAP:
                    report(18.5, f"拉均重组均重计算完成，共 {len(_AVG_WEIGHT_MAP)} 个分组")
                else:
                    report(18.5, "拉均重预计算未生成任何分组（可能无数据命中规则）")

            if use_multiprocess:
                import multiprocessing as mp
                cpu_count = mp.cpu_count()
                pool_size = max(1, min(cpu_count - 1, 8))

                # 保存备份引用（以防 multiprocessing 失败时回退到单进程）
                backup_rows = all_row_tuples
                mp_failed = False

                report(20, f"大数据模式：启动 {pool_size} 个进程并行计算 "
                           f"（CPU {cpu_count}核）")

                # 方案E改进：均匀分配chunk，确保每个进程任务量接近
                # pool_size个进程，分成 pool_size*2 个chunk，让快的进程可以处理更多
                # 用"整除+余数"的方式分配：前 extra_count 个chunk多1行，保证最大差不超过1行
                num_chunks = pool_size * 2
                base_chunk = row_count // num_chunks
                extra_count = row_count % num_chunks
                chunks = []
                pos = 0
                for idx in range(num_chunks):
                    cur_size = base_chunk + (1 if idx < extra_count else 0)
                    end = pos + cur_size
                    chunks.append((all_row_tuples[pos:end], None))
                    pos = end

                del all_row_tuples
                gc.collect()

                report(22, f"已分 {len(chunks)} 块，每块约 {base_chunk:,} 行，启动并行计算...")

                processed = 0
                total_to_process = row_count
                success_count = 0
                exception_count = 0
                total_fee = Decimal("0")

                # 应用层去重集合：确保同一批次不会重复写入
                inserted_tracking = set()

                try:
                    # 使用 concurrent.futures.ProcessPoolExecutor（比 multiprocessing.Pool 更稳定）
                    from concurrent.futures import ProcessPoolExecutor, as_completed
                    with ProcessPoolExecutor(max_workers=pool_size) as executor:
                        # 提交所有任务，用 as_completed 实现类似 imap_unordered 的无序迭代
                        future_to_chunk = {executor.submit(_process_chunk, chunk): idx
                                          for idx, chunk in enumerate(chunks)}
                        for future in as_completed(future_to_chunk):
                            try:
                                chunk_results = future.result()
                            except Exception as chunk_err:
                                # 单个 chunk 失败不影响其他，继续处理
                                continue

                            for b_start in range(0, len(chunk_results), BATCH_SIZE):
                                b_end = min(b_start + BATCH_SIZE, len(chunk_results))
                                batch_to_insert = []
                                for r in chunk_results[b_start:b_end]:
                                    is_exc = r[14]
                                    t_no = r[1] or ""
                                    if t_no and t_no in inserted_tracking:
                                        continue
                                    if t_no:
                                        inserted_tracking.add(t_no)
                                    if is_exc:
                                        exception_count += 1
                                    else:
                                        success_count += 1
                                        total_fee += Decimal(str(r[12]))
                                    batch_to_insert.append((
                                        record_id,
                                        r[0], r[1], r[2], r[3], r[4],
                                        r[5], r[6], r[7], r[8], r[9],
                                        r[10], r[11], r[12], r[13],
                                        is_exc, r[15], r[16],
                                    ))
                                self._bulk_insert_details(session, batch_to_insert)
                                mp_conn = session.connection().connection
                                _lazy_commit_if_needed(mp_conn, len(batch_to_insert))

                            processed += len(chunk_results)
                            progress = 22 + min(processed / total_to_process, 1.0) * 68
                            report(progress,
                                   f"并行计算中... {processed:,}/{total_to_process:,} 行 "
                                   f"({int(progress)}%)  成功 {success_count:,}，异常 {exception_count:,}")

                except Exception as mp_err:
                    # multiprocessing 启动失败，回退到单进程模式
                    mp_failed = True
                    report(25, f"多进程模式启动失败，回退到单进程模式... 原因: {str(mp_err)[:100]}")

                if mp_failed:
                    # 回退到单进程
                    report(25, f"单进程模式：开始计算 {row_count:,} 行...")
                    # 重置计数器
                    success_count = 0
                    exception_count = 0
                    total_fee = Decimal("0")
                    inserted_tracking = set()

                    results = _process_chunk((backup_rows, None))

                    for b_start in range(0, len(results), BATCH_SIZE):
                        b_end = min(b_start + BATCH_SIZE, len(results))
                        batch_to_insert = []
                        for r in results[b_start:b_end]:
                            is_exc = r[14]
                            t_no = r[1] or ""
                            if t_no and t_no in inserted_tracking:
                                continue
                            if t_no:
                                inserted_tracking.add(t_no)
                            if is_exc:
                                exception_count += 1
                            else:
                                success_count += 1
                                total_fee += Decimal(str(r[12]))
                            batch_to_insert.append((
                                record_id,
                                r[0], r[1], r[2], r[3], r[4],
                                r[5], r[6], r[7], r[8], r[9],
                                r[10], r[11], r[12], r[13],
                                is_exc, r[15], r[16],
                            ))
                        self._bulk_insert_details(session, batch_to_insert)
                        _lazy_commit_if_needed(session.connection().connection, len(batch_to_insert))

                    report(90, f"单进程完成。成功 {success_count:,}，异常 {exception_count:,}")

                # 多进程结束：最后一批强制提交
                _lazy_commit_if_needed(session.connection().connection, 0, force_now=True)
                del inserted_tracking

            else:
                # ============ 单进程模式（< 10万行） ============
                report(22, f"单进程模式：开始计算 {row_count:,} 行...")

                results = _process_chunk((all_row_tuples, None))

                # 应用层去重 + 分批写入
                success_count = 0
                exception_count = 0
                total_fee = Decimal("0")
                inserted_tracking = set()

                for b_start in range(0, len(results), BATCH_SIZE):
                    b_end = min(b_start + BATCH_SIZE, len(results))
                    batch_to_insert = []
                    for r in results[b_start:b_end]:
                        is_exc = r[14]
                        t_no = r[1] or ""
                        if t_no and t_no in inserted_tracking:
                            continue
                        if t_no:
                            inserted_tracking.add(t_no)
                        if is_exc:
                            exception_count += 1
                        else:
                            success_count += 1
                            total_fee += Decimal(str(r[12]))
                        batch_to_insert.append((
                            record_id,
                            r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7],
                            r[8], r[9], r[10], r[11], r[12], r[13],
                            is_exc, r[15], r[16]
                        ))
                    self._bulk_insert_details(session, batch_to_insert)
                    # 懒提交：累积到50万行才真正flush磁盘
                    sp_conn = session.connection().connection
                    _lazy_commit_if_needed(sp_conn, len(batch_to_insert))
                    progress = 22 + min(b_end, row_count) / row_count * 68
                    report(progress,
                           f"计算中... {b_end:,}/{row_count:,} 行 ({int(progress)}%)"
                           f"  成功 {success_count:,}，异常 {exception_count:,}")

                # 单进程结束：最后一批强制提交
                _lazy_commit_if_needed(session.connection().connection, 0, force_now=True)
                del results, inserted_tracking, all_row_tuples, seen_global

            # ============ 阶段4：更新记录状态 (90%-100%) ============
            # 关键修复：从数据库读取实际插入的行数，替换可能不准确的内存计数
            actual_rows = success_count + exception_count
            actual_fee = total_fee
            try:
                verify_conn = session.connection().connection
                verify_cur = verify_conn.cursor()
                verify_cur.execute(
                    f"SELECT COUNT(*) FROM {TABLE_FEE_DETAIL} WHERE record_id = ?",
                    (record_id,)
                )
                actual_rows = verify_cur.fetchone()[0]
                verify_cur.execute(
                    f"SELECT COUNT(*) FROM {TABLE_FEE_DETAIL} WHERE record_id = ? AND is_exception = 1",
                    (record_id,)
                )
                actual_exc = verify_cur.fetchone()[0]
                verify_cur.execute(
                    f"SELECT COALESCE(SUM(calculated_fee), 0) FROM {TABLE_FEE_DETAIL} WHERE record_id = ?",
                    (record_id,)
                )
                actual_fee = Decimal(str(round(float(verify_cur.fetchone()[0]), 2)))
                verify_cur.close()
                report(92, f"数据库验证：实际 {actual_rows:,} 行（预期 {success_count + exception_count:,}），"
                           f"成功 {actual_rows - actual_exc:,}，异常 {actual_exc:,}")
                success_count = actual_rows - actual_exc
                exception_count = actual_exc
            except Exception:
                pass

            record.total_rows = actual_rows
            record.success_rows = success_count
            record.error_rows = exception_count
            record.total_fee = actual_fee
            record.status = "success"
            record.completed_at = datetime.now()
            session.commit()

            report(100, f"✅ 全部完成！总计 {actual_rows:,} 行，运费 ¥{float(actual_fee):.2f}")

            return {
                "record_id": record_id,
                "total_fee": float(actual_fee),
                "success_count": success_count,
                "exception_count": exception_count,
                "total_rows": actual_rows
            }

        except Exception as e:
            session.rollback()
            record.status = "failed"
            record.error_message = str(e)
            session.commit()
            raise e
        finally:
            session.close()

    def _bulk_insert_details(self, session, values: List[Tuple]):
        """
        批量写入运费明细 - 使用原生SQLite连接 executemany，比ORM快10-20倍
        """
        if not values:
            return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        params = []
        for v in values:
            params.append((
                v[0], v[1], v[2], v[3], v[4],
                v[5], v[6], v[7], v[8], v[9],
                v[10], v[11], v[12], v[13], v[14],
                v[15], v[16], v[17], now
            ))

        # 使用原生 SQLite 连接做最快的批量写入
        # INSERT OR IGNORE：如果 (record_id, tracking_no) 已存在则跳过，防止重复插入
        conn = session.connection().connection  # 拿到原生 sqlite3.Connection
        cursor = conn.cursor()
        try:
            cursor.executemany(
                f"""
                INSERT OR IGNORE INTO {TABLE_FEE_DETAIL} (
                    record_id, row_index, tracking_no, station_code, station_name,
                    courier_code, courier_name, region_code, region_name, weight,
                    quantity, service_type, original_data, calculated_fee, rule_name,
                    is_exception, exception_type, remark, created_at
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?, ?, ?, ?
                )
                """,
                params
            )
            # 注意：不再在此处 commit()，由调用方 _lazy_commit_if_needed 控制提交频率
            # 百万级数据下：每2.5万行一次 commit = 104次fsync → 每50万行一次 commit = 5次fsync
            # 配合 PRAGMA synchronous=OFF，整体写库性能提升3-5倍
        finally:
            cursor.close()
