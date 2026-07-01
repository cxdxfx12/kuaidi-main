"""
激活码生成工具（管理员专用）
─────────────────────────────
使用说明：
    python keygen.py <机器码> <到期日期>

示例：
    python keygen.py "A1B2-C3D4-E5F6-G7H8-I9J0" "2026-12-31"

输出：25位激活码（格式 XXXXX-XXXXX-XXXXX-XXXXX-XXXXX）

注意：
    1. 机器码由用户从激活窗口复制获得（20字符，含横线）
    2. 到期日期格式为 YYYY-MM-DD
    3. 此工具不得随软件分发给最终用户
"""
import sys
# 添加项目根目录到Python路径
sys.path.insert(0, '.')
from app.core.license_manager import generate_activation_key


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    machine_code = sys.argv[1]
    expire_date = sys.argv[2]

    # 简单验证
    clean_mc = machine_code.replace("-", "").upper()
    if len(clean_mc) != 20:
        print(f"❌ 机器码格式错误：期望20位字符，实际{len(clean_mc)}位")
        print(f"   输入: {machine_code}")
        sys.exit(1)

    # 验证日期格式
    from datetime import datetime
    try:
        parsed = datetime.strptime(expire_date, "%Y-%m-%d")
        if parsed.date() < datetime.now().date():
            print("⚠️  警告：到期日期为过去日期，生成的激活码将立即失效")
    except ValueError:
        print(f"❌ 日期格式错误：请使用 YYYY-MM-DD 格式")
        sys.exit(1)

    # 生成激活码
    key = generate_activation_key(machine_code, expire_date)

    print()
    print("=" * 50)
    print("  激活码生成结果")
    print("=" * 50)
    print(f"  机器码  : {machine_code}")
    print(f"  到期日期: {expire_date}")
    print(f"  激活码  : {key}")
    print("=" * 50)
    print()
    print("请将以上激活码发送给用户完成激活。")
    print()


if __name__ == "__main__":
    main()
