# SQL每日一题F1028，关联子查询

## SQL每日一题F1028，关联子查询
create table F1028A
(aid varchar(20),
bid varchar(20)
)
insert into F1028A values (&#39;跑步&#39;,&#39;张三&#39;);
insert into F1028A values (&#39;游泳&#39;,&#39;张三&#39;);
insert into F1028A values (&#39;跳远&#39;,&#39;李四&#39;);
insert into F1028A values (&#39;跳高&#39;,&#39;王五&#39;);

create table F1028B
(aid varchar(20),
bid varchar(20),
cid varchar(20)
)
insert into F1028B values (&#39;跑步&#39;,&#39;张三&#39;,&#39;胜利&#39;);
insert into F1028B values (&#39;游泳&#39;,&#39;张三&#39;,&#39;胜利&#39;);
insert into F1028B values (&#39;跳高&#39;,&#39;王五&#39;,&#39;胜利&#39;);

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
SUM(CASE WHEN cid = &#39;胜利&#39; THEN 1 ELSE 0 END
) bnum
FROM F1028B GROUP BY bid
) t2 ON t2.bid = t1.bid;

--解法二：
SELECT bid
,COUNT(1) AS anum
,ISNULL(
(
SELECT SUM(CASE WHEN cid=&#39;胜利&#39; THEN 1 ELSE 0 END)
FROM F1028B b
WHERE a.bid=b.bid
),0) AS bnum
FROM F1028A a GROUP BY bid

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/sql/sql-daily-question-f1028-correlation-subquery/  

