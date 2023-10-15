---
title: "SQL每日一题F1025，复杂逻辑处理"
categories: ["编程"]
tags: ["SQL"]
date: 2023-08-01
---
## SQL每日一题F1025，复杂逻辑处理

```sql
create table F1025
(
id int,
num int
)

insert into F1025 values(1,5);
insert into F1025 values(2,11);
insert into F1025 values(3,0);
insert into F1025 values(4,-2);
insert into F1025 values(5,2);
insert into F1025 values(6,9);
insert into F1025 values(7,1);
insert into F1025 values(8,-4);
insert into F1025 values(9,-7);
```


-- Q：要求当Num中的数据同时大于上下两行数据，返回值为“是”，
--        当Num中的数据小于上下两行数据中的任何一行，返回值为“否”
-- 例如：11大于5,11大于0，所以11那行返回值为“是”；5小于11，所以5那行返回值为“否”

```sql
SELECT 
    a.id
    ,a.Num
    ,CASE WHEN a.Num>ISNULL((
        SELECT Num FROM F1025 b WHERE b.id=a.id-1 
    ),a.Num-1)
    AND a.Num>ISNULL((
        SELECT Num FROM F1025 c WHERE c.id=a.id+1 
    ),a.Num-1) THEN '是' ELSE '否' END AS Result
FROM F1025 AS a
```