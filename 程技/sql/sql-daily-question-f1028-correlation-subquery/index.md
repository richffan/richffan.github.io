# SQL每日一题F1028，关联子查询

## SQL每日一题F1028，关联子查询
create table F1028A
(aid varchar(20),
bid varchar(20)
)
insert into F1028A values ('跑步','张三');
insert into F1028A values ('游泳','张三');
insert into F1028A values ('跳远','李四');
insert into F1028A values ('跳高','王五');

create table F1028B
(aid varchar(20),
bid varchar(20),
cid varchar(20)
)
insert into F1028B values ('跑步','张三','胜利');
insert into F1028B values ('游泳','张三','胜利');
insert into F1028B values ('跳高','王五','胜利');

-- Q：anum表示每个人参加的项目数，bnum表示每个人在各自项目中胜利的次数，该如何写这个查询？

--解法一：
SELECT
ISNULL(t1.bid, t2.bid) AS bid,
ISNULL(t1.anum, 0) anum,
ISNULL(t2.bnum, 0) bnum
FROM
(
SELECT bid, COUNT(1) AS anum
FROM F1028A GROUP BY bid
) t1
FULL JOIN
(
SELECT bid,
SUM(CASE WHEN cid = '胜利' THEN 1 ELSE 0 END
) bnum
FROM F1028B GROUP BY bid
) t2 ON t2.bid = t1.bid;

--解法二：
SELECT bid
,COUNT(1) AS anum
,ISNULL(
(
SELECT SUM(CASE WHEN cid='胜利' THEN 1 ELSE 0 END)
FROM F1028B b
WHERE a.bid=b.bid
),0) AS bnum
FROM F1028A a GROUP BY bid

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/sql/sql-daily-question-f1028-correlation-subquery/  

