# live_max_online

```sql
-- step0 建表
-- drop table cm.tb_live_logs;
create table cm.tb_live_logs
(
    live_id     int,
    user_id     int,
    ts  datetime,
    type varchar(10)
);

-- step1 插入数据
insert into cm.tb_live_logs
(live_id, user_id, ts, type) values
(901,1001,'2022-10-01 12:00:00','IN'),
(901,1002,'2022-10-01 12:01:00','IN'),
(901,1003,'2022-10-01 12:01:00','IN'),
(901,1004,'2022-10-01 12:02:00','IN'),
(901,1005,'2022-10-01 12:02:00','IN'),
(901,1006,'2022-10-01 12:03:00','IN'),
(901,1007,'2022-10-01 12:03:00','IN'),
(901,1008,'2022-10-01 12:05:00','IN'),
(901,1009,'2022-10-01 12:05:00','IN'),
(901,1010,'2022-10-01 12:06:03','IN'),
(902,1101,'2022-10-01 12:00:00','IN'),
(902,1102,'2022-10-01 12:01:00','IN'),
(902,1103,'2022-10-01 12:01:00','IN'),
(902,1104,'2022-10-01 12:02:00','IN'),
(902,1105,'2022-10-01 12:29:00','IN'),
(902,1106,'2022-10-01 12:30:00','IN'),
(902,1107,'2022-10-01 12:31:00','IN'),
(902,1108,'2022-10-01 12:32:00','IN'),
(902,1109,'2022-10-01 12:39:00','IN'),
(902,1110,'2022-10-01 12:06:03','IN'),
(901,1001,'2022-10-01 12:03:03','OUT'),
(901,1002,'2022-10-01 12:01:00','OUT'),
(901,1003,'2022-10-01 12:03:03','OUT'),
(901,1004,'2022-10-01 12:05:03','OUT'),
(901,1005,'2022-10-01 12:10:03','OUT'),
(901,1006,'2022-10-01 12:03:01','OUT'),
(901,1007,'2022-10-01 12:03:03','OUT'),
(901,1008,'2022-10-01 12:06:12','OUT'),
(901,1009,'2022-10-01 12:06:03','OUT'),
(901,1010,'2022-10-01 12:10:03','OUT'),
(902,1101,'2022-10-01 12:03:03','OUT'),
(902,1102,'2022-10-01 12:03:03','OUT'),
(902,1103,'2022-10-01 12:03:03','OUT'),
(902,1104,'2022-10-01 12:05:03','OUT'),
(902,1105,'2022-10-01 12:30:03','OUT'),
(902,1106,'2022-10-01 12:30:01','OUT'),
(902,1107,'2022-10-01 12:40:03','OUT'),
(902,1108,'2022-10-01 12:44:12','OUT'),
(902,1109,'2022-10-01 12:42:03','OUT'),
(902,1110,'2022-10-01 12:10:03','OUT');

-- step2 IN/OUT转数字
select
    live_id,
    user_id,
    ts ,
    IF(type = 'IN', 1, -1) as contribution
from cm.tb_live_logs;

-- step3 按照时间进行累加，求每个时间点的delta
select
    live_id,
    ts,
    sum(contribution) as new_number
from (select
        live_id,
        user_id,
        ts,
        IF(type = 'IN', 1, -1) as contribution
      from cm.tb_live_logs
      ) t1
group by live_id, ts;

-- step4 按照时间顺序累加
select
    *,
    sum(new_number) over (partition by live_id order by ts) as online_number
from (select
        live_id,
        ts,
        sum(contribution) as new_number
      from (select
                live_id,
                user_id,
                ts,
                IF(type = 'IN', 1, -1) as contribution
            from cm.tb_live_logs
            ) t1
      group by live_id, ts
      ) t2;

-- step5 求累加过程中，最大的值，就是同一时刻最大在线用户
select
    live_id,
    max(online_number) as max_online_number
from (select
        *,
        sum(new_number) over (partition by live_id order by ts) as online_number
        from (select
                live_id,
                ts,
                sum(contribution) as new_number
              from (select
                        live_id,
                        user_id,
                        ts,
                        IF(type = 'IN', 1, -1) as contribution
                    from cm.tb_live_logs
                    ) t1
              group by live_id, ts
              ) t2
        ) t3
group by live_id;
```

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/sql/live_max_online/  

