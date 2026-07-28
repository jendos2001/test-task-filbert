from sqlalchemy import Column, Integer, Float, String, ForeignKey
from sqlalchemy.orm import declarative_base, relationship


Base = declarative_base()


class ProductOrder(Base):
    __tablename__ = 'Product_Order'
    id = Column(Integer, primary_key=True)
    productId = Column(Integer, ForeignKey('Product.id'))
    orderId = Column(Integer, ForeignKey('Order.id'))


class Product(Base):
    __tablename__ = 'Product'
    id = Column(Integer, primary_key=True)
    title = Column(String)
    description = Column(String)
    price = Column(Float)
    orders = relationship(
        'Order', secondary=ProductOrder.__tablename__, overlaps='products')


class Order(Base):
    __tablename__ = 'Order'
    id = Column(Integer, primary_key=True)
    first_name = Column(String)
    last_name = Column(String)
    postal_code = Column(String)
    payment_information = Column(String)
    shipping_information = Column(String)
    tax = Column(Float)
    price_on_cite = Column(Float)
    products = relationship(
        'Product', secondary=ProductOrder.__tablename__, overlaps='orders')
