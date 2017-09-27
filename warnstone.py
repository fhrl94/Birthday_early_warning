# 导入:
import sqlite3

import sys
from sqlalchemy import Column, create_engine, Date, String, Integer, Boolean
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base

cx = sqlite3.connect(sys.path[0]+"/emp.sqlite3")
engine = create_engine("sqlite:///"+sys.path[0]+"/emp.sqlite3", echo=True)
# 创建DBSession类型:
DBSession = sessionmaker(bind=engine)

# 创建对象的基类:
Base = declarative_base()
# 连接数据库
session = DBSession()

#  定义表（职员表、职位关系表）详见active中的 create_table 函数
# 定义User对象:
class EmployeeInfo(Base):
    # 表的名字:
    __tablename__ = 'EmployeeInfo'
    #  职员表 姓名、工号、出生日期、职位唯一ID（涉及到sqlite支不支持UID），需要处理 职务的获取（存在兼职的情况）
    # 表的结构:
    id = Column(Integer(), primary_key=True)
    name = Column(String(20))
    code = Column(String(10))
    birthDate = Column(Date())
    positionID = Column(String(36))
    job = Column(Integer())
    departmentname=Column(String(50))
    IsPrimary = Column(Boolean())

    def __str__(self):
        return self.name


#  上下级关系表
class Relation(Base):
    # 表的名字:
    __tablename__ = 'Relation'
    #  职位ID、上级职位ID
    # 表的结构:
    id = Column(Integer(), primary_key=True)
    positionID = Column(String(36))
    parentID = Column(String(36))

    def __str__(self):
        return self.name


#  转存表（员工姓名、工号、出生日期、部门名称、职位ID、职务（主管、主管1、经理、经理1、总监、部门第一负责人、总经理、董事长）），分周转存表、月转存表
class WeekMapping(Base):
    # 表的名字:
    __tablename__ = 'WeekMapping'
    #  职员表 姓名、工号、出生日期、部门名称、职位唯一ID（涉及到sqlite支不支持UID），需要处理 职务的获取（存在兼职的情况）
    # 表的结构:
    id = Column(Integer(), primary_key=True)
    name = Column(String(20))
    code = Column(String(10))
    birthDate = Column(Date())
    departmentname=Column(String(50))
    positionID = Column(String(36))
    job = Column(Integer())
    date = Column(Date())
    count = Column(Integer())
    director = Column(String(20))
    director1 = Column(String(20))
    manager = Column(String(20))
    manager1 = Column(String(20))
    majordomo = Column(String(20))
    principal = Column(String(20))
    general_manager = Column(String(20))

    def __str__(self):
        return self.name


class MonthMapping(Base):
    # 表的名字:
    __tablename__ = 'MonthMapping'
    #  职员表 姓名、工号、出生日期、部门名称、职位唯一ID（涉及到sqlite支不支持UID），需要处理 职务的获取（存在兼职的情况）
    # 表的结构:
    id = Column(Integer(), primary_key=True)
    name = Column(String(20))
    code = Column(String(10))
    birthDate = Column(Date())
    departmentname=Column(String(50))
    positionID = Column(String(36))
    job = Column(Integer())
    date = Column(Date())
    count = Column(Integer())
    director = Column(String(20))
    director1 = Column(String(20))
    manager = Column(String(20))
    manager1 = Column(String(20))
    majordomo = Column(String(20))
    principal = Column(String(20))
    general_manager = Column(String(20))

    def __str__(self):
        return self.name


# 如果没有创建表，则创建
Base.metadata.create_all(engine)

def stoneobject():
    return session
