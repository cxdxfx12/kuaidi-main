"""
客户-店铺关联表
支持一个客户拥有多个店铺，按客户规则统一结算
"""
from sqlalchemy import Column, Integer, String, DateTime, Text
from datetime import datetime
from app.models.database import Base


class CustomerStore(Base):
    """客户店铺关联表"""
    __tablename__ = "customer_stores"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    customer_code = Column(String(50), nullable=False, index=True, comment="客户编码（主客户）")
    customer_name = Column(String(100), nullable=False, comment="客户名称")
    store_code = Column(String(50), nullable=False, index=True, comment="店铺编码（快递面单上的客户编码）")
    store_name = Column(String(100), nullable=True, comment="店铺名称")
    remark = Column(Text, nullable=True, comment="备注")
    created_at = Column(DateTime, default=datetime.now, comment="创建时间")
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now, comment="更新时间")

    def __repr__(self):
        return f"<CustomerStore(customer={self.customer_code}, store={self.store_code})>"
