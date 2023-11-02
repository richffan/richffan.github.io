# 给定用户登录表，求表中每一天的3天留存率

```sql
-- 给定用户登录表，求表中每一天的"3天留存率"

-- step 0 表准备，用户ID，登录时间
create table cm.tb_user_logs(
    user_id int,
    visit_time datetime
);

insert into cm.tb_user_logs
(user_id, visit_time)
values
(1001, '2022-10-01 08:01:23'),
(1001, '2022-10-01 08:11:15'),
(1002, '2022-10-01 08:22:19'),
(1003, '2022-10-01 09:00:53'),
(1002, '2022-10-02 18:00:13'),
(1004, '2022-10-02 13:30:43'),
(1004, '2022-10-02 15:06:22'),
(1005, '2022-10-02 08:00:39'),
(1002, '2022-10-03 08:00:13'),
(1003, '2022-10-03 18:00:13'),
(1004, '2022-10-03 21:00:13'),
(1006, '2022-10-03 22:00:13'),
(1001, '2022-10-04 11:10:13'),
(1002, '2022-10-04 12:00:13'),
(1002, '2022-10-04 09:00:13'),
(1004, '2022-10-04 08:00:13'),
(1006, '2022-10-04 08:00:13'),
(1004, '2022-10-05 08:00:13'),
(1005, '2022-10-05 08:00:13'),
(1002, '2022-10-05 08:00:43'),
(1003, '2022-10-05 12:00:13'),
(1004, '2022-10-05 10:00:43'),
(1006, '2022-10-05 08:00:11'),
(1001, '2022-10-06 09:00:47'),
(1001, '2022-10-06 07:00:15'),
(1002, '2022-10-06 18:00:43'),
(1003, '2022-10-07 20:00:19'),
(1002, '2022-10-07 21:00:23'),
(1004, '2022-10-07 22:00:43');

-- step 1 创建视图，进行用户ID和时间去重
create view A as
select distinct
       user_id,
       date(visit_time) as dt
from cm.tb_user_logs;

-- step2. 计算每天的"3天活跃用户数"，即对于每一天而言，在当天的活跃用户中，3天后还活跃的那些用户---分子
select
    t1.dt      as dt,
    count(t1.user_id) as 3day_active_cnt
from A as t1
         join A t2 on t1.dt = t2.dt - 3
where 1 = 1
  and t1.user_id = t2.user_id
group by t1.dt;

-- step3. 计算每天的活跃用户数---分母
select
    dt,
    count(user_id) as active_cnt
from A
group by dt;

-- step4. 将上述两个步骤结果，按照同一天的日期，进行关联，求比率即可
select t3.dt,
       t3.3days_active_cnt,
       t4.active_cnt,
       t3.3days_active_cnt * 1.0 / t4.active_cnt as 3day_alive_ratio
from (select t1.dt,
             count(t1.user_id) as 3days_active_cnt
            from A as t1
                     join A t2 on t1.dt = t2.dt - 3
            where 1=1
                and t1.user_id = t2.user_id
            group by dt) as t3
         join
         (select
              dt,
              count(user_id) as active_cnt
          from A
          group by dt
          ) as t4
         on t3.dt = t4.dt
order by t3.dt;

select distinct (user_id)
from cm.tb_user_logs
where date(visit_time)='2022-10-01';

select distinct (user_id)
from cm.tb_user_logs
where date(visit_time)='2022-10-04';
```


---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/sql/3days_alive_users/  

