# top10_sql_skills

```sql
-- 01
-- bad
create table cm.tb_user(
    id int,
    name varchar(255),
    age int,
    gender varchar(16)
);
-- good
create table cm.tb_user(
    id int primary key,
    name varchar(255),
    age int,
    gender varchar(16),
    store_time datetime default now()
);

-- 02
-- bad
create table cm.tb_user(
    id int primary key ,
    name varchar(255) not null default &#39;NULL&#39;,
    age int,
    gender varchar(16),
    create_time datetime default now(),
    update_time datetime default now()
);
-- good
create table cm.tb_user(
    id int primary key ,
    name varchar(255) not null default &#39;-&#39;,
    age int,
    gender varchar(16),
    store_time datetime default now(),
    update_time datetime default now()
);

-- 03
-- bad
select min(age), max(age) from cm.tb_user where age &lt; 40 group by gender;
-- good
select
    min(age),
    max(age)
from cm.tb_user
where age &lt; 40
group by gender;

-- 04
-- bad
select id,
       name
from cm.tb_user
where age &gt; 10
  and id&lt;=10010;
-- good
select id,
       name
from cm.tb_user
where 1=1
  and age &gt; 10
  and id&lt;=10010;

-- 05
-- bad
insert into cm.tb_user
values
    (1001, &#39;Mike&#39;, 20, &#39;M&#39;);
-- good
insert into cm.tb_user
(id, name, age, gender)
values
    (1001, &#39;Mike&#39;, 20, &#39;M&#39;);

-- 06
-- good
explain
select *
from cm.tb_user
where id=1001 and age = 20;

-- 07
-- bad
select *
from cm.tb_user;
-- good
select id,
       name
from cm.tb_user;

-- 08
-- 分批次&#43;limit进行delete或update
delete
from cm.tb_user
where age &gt; 35
limit 5;

-- 09
-- bad
delete from cm.tb_user
where age &lt;= 20;
update cm.tb_user
set age = age &#43; 1;
-- good
begin;
delete from cm.tb_user
where age &lt;= 20;
update cm.tb_user
set age = age &#43; 1;
commit;

-- 10
-- data sample
insert into cm.data_count
(dt, count)
values
    (&#39;2023-06-20&#39;, 101),
    (&#39;2023-06-21&#39;, 231),
    (&#39;2023-06-22&#39;, 170),
    (&#39;2023-06-23&#39;, 146),
    (&#39;2023-06-24&#39;, 187),
    (&#39;2023-06-25&#39;, 123),
    (&#39;2023-06-26&#39;, 221),
    (&#39;2023-06-27&#39;, 101),
    (&#39;2023-06-28&#39;, 103),
    (&#39;2023-06-29&#39;, 122),
    (&#39;2023-06-30&#39;, 144);
-- bad
select dt, count
from cm.data_count t1;
-- good
select
    t1.dt,
    t1.count * 1.0 / t2.count as ratio
from cm.data_count t1
    join cm.data_count t2
on datediff(t1.dt, t2.dt)=7;
```

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/sql/top10_sql_skills/  

