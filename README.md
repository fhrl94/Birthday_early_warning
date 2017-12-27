生日预警思路

1、数据获取

从金蝶中获取人员数据存储至 EmployeeInfo ，并标识兼职人员 IsPrimary=0

将上下级关系存储至 Relation 等待后期寻找上级使用

2、使用 unloading() 调用 afterday 之后的天数 number 的生日人员名单 存储到相应表中（ DaysMapping、MonthMapping ）
并根据 Relation 寻找所有上级

3、根据 DaysMapping、MonthMapping中的数据 ，获取**去重**后的上级人员名单，根据名单查询对应人员，从而构建
邮件正文主题。发送邮件。

代码重构，[详见项目Britdhay_warning](https://github.com/fhrl94/Britdhay_warning.git)
