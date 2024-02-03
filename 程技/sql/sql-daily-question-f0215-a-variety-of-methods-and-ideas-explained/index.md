# SQL每日一题F0215，多种方法及思路讲解

## SQL每日一题F0215，多种方法及思路讲解
```sql
CREATE TABLE F0215

(

StuID INT,

CID VARCHAR(10),

Course INT

)

  

INSERT INTO F0215 VALUES

(1,&#39;001&#39;,67),

(1,&#39;002&#39;,89),

(1,&#39;003&#39;,94),

(2,&#39;001&#39;,95),

(2,&#39;002&#39;,88),

(2,&#39;004&#39;,78),

(3,&#39;001&#39;,94),

(3,&#39;002&#39;,77),

(3,&#39;003&#39;,90)

select * from f0215

/*

查询出既学过&#39;001&#39;课程，也学过&#39;003&#39;号课程的学生ID

*/

  

--错误写法

SELECT * FROM F0215

WHERE CID=&#39;001&#39; AND CID=&#39;003&#39;

  

SELECT * FROM F0215

WHERE CID=&#39;001&#39; OR CID=&#39;003&#39;

  

--思路一，取自连接符合条件的学生

SELECT T1.STUID FROM

( SELECT STUID FROM F0215 WHERE CID=&#39;001&#39; ) T1

INNER JOIN

(SELECT STUID FROM F0215 WHERE CID=&#39;003&#39; )  T2  

ON T2.STUID=T1.STUID

  
  

SELECT A.STUID

FROM F0215 A,F0215 B

WHERE A.CID = &#39;001&#39;

AND B.CID = &#39;003&#39;

AND A.STUID = B.STUID

  

--思路二，使用交集取出同时满足条件的学生

SELECT SC.STUID

FROM F0215 SC

WHERE  SC.CID=&#39;001&#39;

INTERSECT

SELECT SC.STUID

FROM F0215 SC

WHERE  SC.CID=&#39;003&#39;

  

--思路三

  

SELECT StuID FROM F0215

WHERE CID IN (&#39;001&#39;,&#39;003&#39;)

GROUP BY StuID

HAVING COUNT(StuID)=2

  

--思路四（思路三的变体）

SELECT STUID FROM

(

SELECT STUID FROM F0215 WHERE CID = &#39;001&#39;

UNION ALL

SELECT STUID FROM F0215 WHERE CID = &#39;003&#39;

) A

GROUP BY STUID HAVING COUNT(STUID) = 2
```

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/sql/sql-daily-question-f0215-a-variety-of-methods-and-ideas-explained/  

