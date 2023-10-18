# SQL每日一题F1021，while循环操作

## SQL每日一题F1021，while循环操作
/*
写一个查询语句要求
求出整数1到100之间排除55后的和
预计结果是4995
该如何写这个查询？

要求：使用while循环
*/

DECLARE @i INT
DECLARE @sum INT
SET @i=0
SET @sum=0
WHILE @i<100
BEGIN
SET @i=@i+1
IF @i=55
CONTINUE
ELSE
SET @sum=@sum+@i
END
PRINT @sum

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/2023/08/sql-daily-question-f1021-while-loop-operation/  

