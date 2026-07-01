"""
硬件绑定授权模块
────────────────
方案：机器码（MAC + 硬盘序列号 + 主板序列号）→ SHA256 指纹
激活码：到期天数偏移(2B) + HMAC-SHA256(机器码|到期日期, 密钥)[:13B] = 15B
      → BASE32 → 24字符 → 补齐25字符 → XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
存储：加密的 license.dat（用户数据目录下）

安全特性：
- 机器码不可逆（SHA256）
- 激活码绑定机器 + 到期日期（换机器或过期均失效）
- 到期日期嵌入激活码，验证 O(1) 复杂度
- 密钥仅内置在应用和管理工具中
- 防时间篡改：记录上次启动时间，发现时间倒退则锁定
"""
import os
import sys
import json
import hmac
import hashlib
import base64
import struct
import subprocess
import uuid
from datetime import datetime, date
from typing import Tuple, Optional

from app.models.path_config import get_data_dir

# ===================== 常量 =====================

# 授权密钥（与 keygen 工具共用，不要泄露）
_SECRET_KEY = b"SFee_Express_Billing_2024_HMAC_K3y!@#"

# 机器码显示格式
MACHINE_CODE_GROUPS = 5
MACHINE_CODE_CHARS_PER_GROUP = 4

# 激活码格式（便于用户输入）
LICENSE_KEY_GROUPS = 5
LICENSE_KEY_CHARS_PER_GROUP = 5

# 日期基准：2024-01-01（用于天数偏移编码）
_EPOCH_DATE = date(2024, 1, 1)

# 存储文件名
LICENSE_FILE_NAME = "license.dat"

# 消息
MSG_NOT_ACTIVATED = "软件尚未激活，请先激活"
MSG_EXPIRED = "软件授权已到期，请联系管理员续费"
MSG_MACHINE_CHANGED = "机器环境已变更，请重新激活"
MSG_TIME_TAMPERED = "检测到系统时间异常，软件已锁定"


# ===================== 机器码生成 =====================

def _run_cmd(cmd: list) -> str:
    """执行命令并返回 stdout（静默失败）"""
    try:
        # Windows 下隐藏窗口
        si = None
        if sys.platform == "win32":
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=5,
            startupinfo=si
        )
        return result.stdout.strip()
    except Exception:
        return ""


def _get_mac_addresses() -> str:
    """获取所有物理网卡 MAC 地址（排除虚拟/回环）"""
    macs = []
    try:
        # 方法1：uuid 获取（跨平台）
        node = uuid.getnode()
        mac = ':'.join(f'{(node >> elements) & 0xff:02x}'
                       for elements in range(0, 2*6, 2))[::-1]
        if mac != "00:00:00:00:00:00":
            macs.append(mac.replace(":", "").upper())
    except Exception:
        pass

    # 方法2：wmic（Windows 精确获取物理网卡）
    if sys.platform == "win32":
        output = _run_cmd([
            "wmic", "nic", "where", "PhysicalAdapter=TRUE",
            "get", "MACAddress", "/format:csv"
        ])
        for line in output.split("\n"):
            line = line.strip()
            if ":" in line and len(line) > 10:
                # 提取 MAC 地址
                parts = line.split(",")
                for p in parts:
                    p = p.strip().replace(":", "").upper()
                    if len(p) == 12 and all(c in "0123456789ABCDEF" for c in p):
                        if p not in macs:
                            macs.append(p)

    # 去重排序（保证一致性）
    macs = sorted(set(macs))
    return ",".join(macs[:3])  # 最多取3个，防止过多


def _get_disk_serial() -> str:
    """获取系统盘序列号"""
    if sys.platform == "win32":
        serial = _run_cmd(["wmic", "diskdrive", "get", "SerialNumber", "/format:csv"])
        if serial:
            for line in serial.split("\n"):
                line = line.strip()
                if line and not line.startswith("Node"):
                    # 提取 SerialNumber
                    parts = line.split(",")
                    for p in parts:
                        p = p.strip()
                        if p and len(p) >= 4:
                            return p
    else:
        # macOS / Linux
        if sys.platform == "darwin":
            serial = _run_cmd(["system_profiler", "SPNVMeDataType"])
            if not serial:
                serial = _run_cmd(["system_profiler", "SPSerialATADataType"])
            # 提取序列号（简化处理）
            for line in serial.split("\n"):
                if "Serial" in line:
                    return line.strip()
        else:
            serial = _run_cmd(["lsblk", "-o", "SERIAL", "-n", "-d"])
            if serial:
                return serial.strip().split("\n")[0]
    return ""


def _get_motherboard_serial() -> str:
    """获取主板序列号"""
    if sys.platform == "win32":
        serial = _run_cmd(["wmic", "baseboard", "get", "SerialNumber", "/format:csv"])
        if serial:
            for line in serial.split("\n"):
                line = line.strip()
                if line and not line.startswith("Node"):
                    parts = line.split(",")
                    for p in parts:
                        p = p.strip()
                        if p and len(p) >= 4:
                            return p
    else:
        if sys.platform == "darwin":
            serial = _run_cmd(["ioreg", "-l", "-w0", "|", "grep", "IOPlatformSerialNumber"])
            if serial:
                return serial.strip()
        else:
            serial = _run_cmd(["dmidecode", "-s", "baseboard-serial-number"])
            if serial:
                return serial.strip()
    return ""


def get_machine_code_raw() -> str:
    """获取原始硬件指纹（内部使用）"""
    parts = [
        _get_mac_addresses(),
        _get_disk_serial(),
        _get_motherboard_serial(),
    ]
    return "|".join(parts)


def get_machine_code_display() -> str:
    """
    获取用户可读的机器码（20字符，用于展示和输入）
    格式：XXXX-XXXX-XXXX-XXXX-XXXX
    """
    raw = get_machine_code_raw()
    h = hashlib.sha256(raw.encode("utf-8")).hexdigest().upper()
    # 取前20位
    code = h[:MACHINE_CODE_GROUPS * MACHINE_CODE_CHARS_PER_GROUP]
    # 格式化为 XXXX-XXXX-XXXX-XXXX-XXXX
    groups = [
        code[i:i + MACHINE_CODE_CHARS_PER_GROUP]
        for i in range(0, len(code), MACHINE_CODE_CHARS_PER_GROUP)
    ]
    return "-".join(groups)


def get_machine_fingerprint() -> str:
    """
    获取机器指纹（SHA256，用于签名验证，不可逆）
    """
    raw = get_machine_code_raw()
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


# ===================== 激活码生成与验证 =====================

def generate_activation_key(machine_code: str, expire_date_str: str) -> str:
    """
    生成激活码（管理员工具用）
    
    方案：到期天数偏移(2B) + HMAC-SHA256[:13B] = 15B → BASE32 → 24字符
    补齐到 25 字符以适配 5×5 显示格式。
    
    参数：
        machine_code: 用户机器码（20字符，含横线）
        expire_date_str: 到期日期，格式 YYYY-MM-DD
    
    返回：
        25字符激活码（格式 XXXXX-XXXXX-XXXXX-XXXXX-XXXXX）
    """
    clean_mc = machine_code.replace("-", "").upper()
    
    # 到期日期 → 天数偏移（2字节大端序，最大到 2203年）
    expire_dt = datetime.strptime(expire_date_str, "%Y-%m-%d").date()
    days_offset = (expire_dt - _EPOCH_DATE).days
    if days_offset < 0 or days_offset > 65535:
        raise ValueError(f"到期日期超出范围")
    date_bytes = struct.pack(">H", days_offset)
    
    # HMAC 签名（取13字节）
    payload = f"{clean_mc}|{expire_date_str}"
    sig = hmac.new(_SECRET_KEY, payload.encode("utf-8"), hashlib.sha256).digest()[:13]
    
    # 拼接 2+13=15 字节 → BASE32 → 24 字符
    key_bytes = date_bytes + sig
    b32 = base64.b32encode(key_bytes).decode("utf-8")  # 24 chars
    # 补齐到 25 字符（5组×5），用 'A' 填充（BASE32 不含 'A' = 0）
    # 'A' 填充位在解码时会被自动忽略
    code = b32.ljust(25, 'A')
    
    groups = [
        code[i:i + LICENSE_KEY_CHARS_PER_GROUP]
        for i in range(0, LICENSE_KEY_GROUPS * LICENSE_KEY_CHARS_PER_GROUP, LICENSE_KEY_CHARS_PER_GROUP)
    ]
    return "-".join(groups[:LICENSE_KEY_GROUPS])


def verify_activation_key(machine_code: str, key: str) -> Tuple[bool, str, Optional[str]]:
    """
    验证激活码（O(1) 复杂度）
    
    处理逻辑：
    - 去掉填充的 'A' 字符，还原 24 字符 BASE32
    - BASE32 解码出 15 字节（2字节日期 + 13字节HMAC）
    - 验证签名匹配
    
    返回：
        (是否有效, 消息, 到期日期字符串或None)
    """
    clean_mc = machine_code.replace("-", "").upper()
    clean_key = key.replace("-", "").upper().strip()
    
    if len(clean_key) < LICENSE_KEY_GROUPS * LICENSE_KEY_CHARS_PER_GROUP:
        return False, "激活码格式不正确（需要25位），请检查后重新输入", None
    
    # 去掉末尾填充的 'A' 字符（BASE32 中 'A' = 0）
    b32_clean = clean_key.rstrip('A')
    
    # 还原为 24 字符（用于 BASE32 解码）
    if len(b32_clean) > 24:
        b32_clean = b32_clean[:24]
    elif len(b32_clean) < 24:
        # 如果去掉 A 后不足 24，补齐到 24（用 A = 0）
        b32_clean = b32_clean.ljust(24, 'A')
    
    # BASE32 解码（24字符 → 15字节）
    padded = b32_clean + "=" * ((8 - len(b32_clean) % 8) % 8)
    try:
        key_bytes = base64.b32decode(padded)
    except Exception:
        return False, "激活码格式不正确，请检查后重新输入", None
    
    if len(key_bytes) < 15:
        return False, "激活码不完整", None
    
    try:
        days_offset = struct.unpack(">H", key_bytes[:2])[0]
        expire_date = _EPOCH_DATE
        from datetime import timedelta
        expire_date = expire_date + timedelta(days=days_offset)
    except (struct.error, ValueError, OverflowError):
        return False, "激活码已损坏", None
    
    expire_str = expire_date.strftime("%Y-%m-%d")
    
    # 验证 HMAC 签名
    payload = f"{clean_mc}|{expire_str}"
    expected_sig = hmac.new(_SECRET_KEY, payload.encode("utf-8"), hashlib.sha256).digest()[:13]
    actual_sig = key_bytes[2:15]
    
    if not hmac.compare_digest(expected_sig, actual_sig):
        return False, "激活码无效，请检查机器码和激活码是否正确", None
    
    today = date.today()
    remaining = (expire_date - today).days
    if remaining < 0:
        return False, f"该激活码已到期（{expire_str}）", None
    
    return True, f"激活成功，有效期至 {expire_str}（剩余 {remaining} 天）", expire_str


# ===================== 授权文件管理 =====================

def _get_license_path() -> str:
    """获取授权文件路径"""
    return os.path.join(get_data_dir(), LICENSE_FILE_NAME)


def _encrypt_license_data(data: dict) -> str:
    """加密授权数据（简单 XOR + BASE64，防直接查看）"""
    json_str = json.dumps(data, sort_keys=True)
    key_bytes = _SECRET_KEY
    encrypted = bytes([
        ord(c) ^ key_bytes[i % len(key_bytes)]
        for i, c in enumerate(json_str)
    ])
    return base64.b64encode(encrypted).decode("utf-8")


def _decrypt_license_data(encrypted_str: str) -> Optional[dict]:
    """解密授权数据"""
    try:
        encrypted = base64.b64decode(encrypted_str)
        key_bytes = _SECRET_KEY
        decrypted = bytes([
            b ^ key_bytes[i % len(key_bytes)]
            for i, b in enumerate(encrypted)
        ])
        return json.loads(decrypted.decode("utf-8"))
    except Exception:
        return None


def save_license(machine_code: str, key: str, expire_date: str) -> Tuple[bool, str]:
    """保存授权信息到文件
    
    返回: (是否成功, 错误信息)
    """
    try:
        data = {
            "machine_fingerprint": get_machine_fingerprint(),
            "activation_key_hash": hashlib.sha256(key.encode()).hexdigest(),
            "expire_date": expire_date,
            "activated_date": date.today().strftime("%Y-%m-%d"),
            "last_check_date": date.today().strftime("%Y-%m-%d"),
        }
        encrypted = _encrypt_license_data(data)
        path = _get_license_path()
        dir_path = os.path.dirname(path)
        os.makedirs(dir_path, exist_ok=True)
        
        # 如果文件已存在且被设为隐藏/只读，先取消属性
        if sys.platform == "win32" and os.path.exists(path):
            try:
                import ctypes
                # FILE_ATTRIBUTE_NORMAL = 0x80，取消隐藏和只读
                ctypes.windll.kernel32.SetFileAttributesW(path, 0x80)
            except Exception:
                pass
        
        with open(path, "w", encoding="utf-8") as f:
            f.write(encrypted)
        
        # 设置隐藏属性（Windows）
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.kernel32.SetFileAttributesW(path, 2)
            except Exception:
                pass
        return True, ""
    except Exception as e:
        import traceback
        err_msg = f"{type(e).__name__}: {e}"
        traceback.print_exc()
        return False, err_msg


def load_license() -> Optional[dict]:
    """读取授权信息"""
    path = _get_license_path()
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            encrypted = f.read().strip()
        return _decrypt_license_data(encrypted)
    except Exception:
        return None


def delete_license() -> bool:
    """删除授权文件（用于反激活或重置）"""
    path = _get_license_path()
    try:
        if os.path.exists(path):
            os.remove(path)
        return True
    except Exception:
        return False


# ===================== 启动时授权校验 =====================

def check_license_on_startup() -> Tuple[bool, str]:
    """
    软件启动时调用，检查授权状态
    
    返回 (是否通过, 消息)
    - 已激活且未过期 → (True, "授权有效")
    - 未激活 → (False, "软件尚未激活，请先激活")
    - 已过期 → (False, "软件授权已到期...")
    - 机器变更 → (False, "机器环境已变更...")
    - 时间异常 → (False, "检测到系统时间异常...")
    """
    lic = load_license()
    
    if lic is None:
        return False, MSG_NOT_ACTIVATED
    
    # 1) 检查机器指纹
    current_fp = get_machine_fingerprint()
    stored_fp = lic.get("machine_fingerprint", "")
    if not hmac.compare_digest(current_fp, stored_fp):
        return False, MSG_MACHINE_CHANGED
    
    # 2) 防时间篡改检查
    today_str = date.today().strftime("%Y-%m-%d")
    last_check = lic.get("last_check_date", "")
    
    if last_check:
        try:
            last_date = datetime.strptime(last_check, "%Y-%m-%d").date()
            today = date.today()
            # 如果系统时间比上次启动还早（倒退），判定为时间篡改
            if today < last_date:
                # 锁定：删除授权文件
                delete_license()
                return False, MSG_TIME_TAMPERED
        except ValueError:
            pass
    
    # 更新上次检查时间
    try:
        lic["last_check_date"] = today_str
        encrypted = _encrypt_license_data(lic)
        path = _get_license_path()
        # 写入前先取消隐藏属性
        if sys.platform == "win32" and os.path.exists(path):
            try:
                import ctypes
                ctypes.windll.kernel32.SetFileAttributesW(path, 0x80)
            except Exception:
                pass
        with open(path, "w", encoding="utf-8") as f:
            f.write(encrypted)
        # 恢复隐藏属性
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.kernel32.SetFileAttributesW(path, 2)
            except Exception:
                pass
    except Exception:
        pass
    
    # 3) 检查到期日期
    expire_str = lic.get("expire_date", "")
    try:
        expire_date = datetime.strptime(expire_str, "%Y-%m-%d").date()
        remaining = (expire_date - date.today()).days
        
        if remaining < 0:
            return False, MSG_EXPIRED
        
        if remaining <= 30:
            return True, f"授权即将到期（剩余 {remaining} 天）"
        
        return True, "授权有效"
    except ValueError:
        return False, "授权数据异常"


def get_license_info() -> dict:
    """获取授权详细信息（用于界面展示）"""
    lic = load_license()
    if not lic:
        return {
            "activated": False,
            "machine_code": get_machine_code_display(),
            "expire_date": "",
            "activated_date": "",
            "remaining_days": -1,
            "message": MSG_NOT_ACTIVATED,
        }
    
    expire_str = lic.get("expire_date", "")
    try:
        expire_date = datetime.strptime(expire_str, "%Y-%m-%d").date()
        remaining = (expire_date - date.today()).days
    except ValueError:
        remaining = -1
    
    return {
        "activated": True,
        "machine_code": get_machine_code_display(),
        "expire_date": expire_str,
        "activated_date": lic.get("activated_date", ""),
        "remaining_days": remaining,
        "message": "已激活" if remaining > 0 else "已过期",
    }
