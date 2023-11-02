# SQL每日一题(20230814)

## SQL每日一题(20230814)

### 题目

有如下两张表G0227A（客户表）

|Id|Name|
|---|---|
|1|曹操|
|2|关于|
|3|刘备|
|4|张飞|

G0227B（订单表）

|Id|CustomerId|
|---|---|
|1|3|
|2|1|

查询G0227B表（订单表）中找出从来没有买过商品的用户。

预计结果如下：

|Id|Name|
|---|---|
|2|关于|

### 要求：用至少四种方法求解。

```sql
create table G0814A  
(  
Id int,  
Name varchar(20)  
)  
  
insert into G0814A values (1,'曹操');  
insert into G0814A values (2,'关羽');  
insert into G0814A values (3,'刘备');  
insert into G0814A values (4,'张飞');  
  
create table G0814B(  
Id int,  
CustomerId int  
)  
  
insert into G0814B values (1,3);  
insert into G0814B values (2,1);
```

### 参考答案

```sql
-- 方法一: 关联查询
SELECT a.* FROM G0814A aLEFT JOIN G0814B b ON b.customerid=a.idWHERE b.customerid IS NULL;

-- 方法二:不存在 not exists
select *from G0814A awhere not exists(select * from G0814B bwhere a.id = b.customerid;

-- 方法三:不包含 not in
select￥from G0814A awhere a.id not inselect b.customerid from G0814B b);

-- 方法四: 差集
select * from G0814Awhere id inselect a.id from G0814A aexceptselect b.customerid from G0814B b
);
```

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/sql/sql-daily-question-20230814/  

