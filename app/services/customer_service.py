"""
客户-店铺管理服务
提供客户与店铺的关联查询、增删等功能
"""
from typing import Dict, List, Optional, Set
from app.models.database import get_session
from app.models.customer_store import CustomerStore


class CustomerService:
    """客户店铺关联管理"""

    def __init__(self):
        self._cache: Optional[Dict[str, str]] = None  # store_code → customer_code
        self._customer_names: Optional[Dict[str, str]] = None  # customer_code → customer_name

    def _refresh_cache(self):
        """刷新内存缓存"""
        session = get_session()
        try:
            rows = session.query(CustomerStore).all()
            self._cache = {}
            self._customer_names = {}
            for r in rows:
                self._cache[r.store_code] = r.customer_code
                if r.customer_code not in self._customer_names:
                    self._customer_names[r.customer_code] = r.customer_name
        finally:
            session.close()

    def ensure_cache(self):
        """确保缓存已初始化"""
        if self._cache is None:
            self._refresh_cache()

    def get_parent_customer(self, store_code: str) -> str:
        """根据店铺编码获取父客户编码（未找到返回原值）"""
        self.ensure_cache()
        sc = (store_code or "").strip()
        if not sc:
            return sc
        return self._cache.get(sc, sc)

    def get_customer_name(self, customer_code: str) -> str:
        """获取客户名称"""
        self.ensure_cache()
        return self._customer_names.get(customer_code, "")

    def get_all_stores_for_customer(self, customer_code: str) -> List[Dict]:
        """获取某客户下的所有店铺"""
        session = get_session()
        try:
            rows = session.query(CustomerStore).filter(
                CustomerStore.customer_code == customer_code
            ).order_by(CustomerStore.store_code).all()
            return [
                {
                    "id": r.id,
                    "store_code": r.store_code,
                    "store_name": r.store_name or r.store_code,
                    "remark": r.remark or "",
                }
                for r in rows
            ]
        finally:
            session.close()

    def get_all_customers(self) -> List[Dict]:
        """获取所有客户（去重，含店铺数量）"""
        session = get_session()
        try:
            rows = session.query(CustomerStore).order_by(
                CustomerStore.customer_code
            ).all()

            customer_map: Dict[str, Dict] = {}
            for r in rows:
                cc = r.customer_code
                if cc not in customer_map:
                    customer_map[cc] = {
                        "customer_code": cc,
                        "customer_name": r.customer_name,
                        "stores": [],
                    }
                customer_map[cc]["stores"].append({
                    "store_code": r.store_code,
                    "store_name": r.store_name or r.store_code,
                })

            return list(customer_map.values())
        finally:
            session.close()

    def add_store(self, customer_code: str, customer_name: str,
                  store_code: str, store_name: str = "", remark: str = "") -> bool:
        """为某客户添加一个店铺"""
        if not customer_code.strip() or not store_code.strip():
            return False
        session = get_session()
        try:
            # 检查是否已存在
            existing = session.query(CustomerStore).filter(
                CustomerStore.customer_code == customer_code.strip(),
                CustomerStore.store_code == store_code.strip(),
            ).first()
            if existing:
                # 更新名称
                existing.customer_name = customer_name.strip()
                existing.store_name = store_name.strip()
                existing.remark = remark.strip()
                session.commit()
                self._cache = None  # 使缓存失效
                return True

            cs = CustomerStore(
                customer_code=customer_code.strip(),
                customer_name=customer_name.strip(),
                store_code=store_code.strip(),
                store_name=store_name.strip() or store_code.strip(),
                remark=remark.strip(),
            )
            session.add(cs)
            session.commit()
            self._cache = None
            return True
        except Exception:
            session.rollback()
            return False
        finally:
            session.close()

    def delete_store(self, store_id: int) -> bool:
        """删除一条店铺关联"""
        session = get_session()
        try:
            cs = session.query(CustomerStore).filter(CustomerStore.id == store_id).first()
            if cs:
                session.delete(cs)
                session.commit()
                self._cache = None
                return True
            return False
        except Exception:
            session.rollback()
            return False
        finally:
            session.close()

    def delete_stores_by_code(self, customer_code: str, store_code: str) -> bool:
        """按编码删除店铺关联"""
        session = get_session()
        try:
            session.query(CustomerStore).filter(
                CustomerStore.customer_code == customer_code,
                CustomerStore.store_code == store_code,
            ).delete()
            session.commit()
            self._cache = None
            return True
        except Exception:
            session.rollback()
            return False
        finally:
            session.close()

    def batch_add_stores(self, customer_code: str, customer_name: str,
                         stores: List[Dict]) -> Dict:
        """批量添加店铺
        :param stores: [{"store_code": "...", "store_name": "...", "remark": "..."}, ...]
        :return: {"success": bool, "inserted": int, "updated": int}
        """
        if not customer_code.strip() or not stores:
            return {"success": False, "inserted": 0, "updated": 0}
        session = get_session()
        inserted = 0
        updated = 0
        try:
            for s in stores:
                sc = str(s.get("store_code", "")).strip()
                if not sc:
                    continue
                existing = session.query(CustomerStore).filter(
                    CustomerStore.customer_code == customer_code.strip(),
                    CustomerStore.store_code == sc,
                ).first()
                if existing:
                    existing.store_name = str(s.get("store_name", "")).strip() or sc
                    existing.remark = str(s.get("remark", "")).strip()
                    updated += 1
                else:
                    cs = CustomerStore(
                        customer_code=customer_code.strip(),
                        customer_name=customer_name.strip(),
                        store_code=sc,
                        store_name=str(s.get("store_name", "")).strip() or sc,
                        remark=str(s.get("remark", "")).strip(),
                    )
                    session.add(cs)
                    inserted += 1
            session.commit()
            self._cache = None
            return {"success": True, "inserted": inserted, "updated": updated}
        except Exception:
            session.rollback()
            return {"success": False, "inserted": inserted, "updated": updated}
        finally:
            session.close()
