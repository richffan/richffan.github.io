# 求占据前90%销售额的商品类型

```sql
-- 求占据前90%销售额的商品类型
-- step0. 准备数据
create table cm.tb_sale_amount(
    good_category int comment &#39;商品类型ID&#39;,
    sale_date date comment &#39;销售日期&#39;,
    amount int comment &#39;销售额&#39;,
    primary key(good_category, sale_date)
);


truncate cm.tb_sale_amount;

insert into cm.tb_sale_amount
(good_category, sale_date, amount) values
(1003, &#39;2022-01-10&#39;, 264),
(1001, &#39;2022-06-01&#39;, 21),
(1005, &#39;2022-06-01&#39;, 73),
(1002, &#39;2022-06-27&#39;, 44),
(1006, &#39;2022-06-27&#39;, 405),
(1003, &#39;2022-09-10&#39;, 16),
(1005, &#39;2022-09-13&#39;, 72),
(1004, &#39;2022-10-01&#39;, 29),
(1005, &#39;2022-10-03&#39;, 332),
(1001, &#39;2022-10-29&#39;, 10),
(1006, &#39;2022-10-29&#39;, 137),
(1002, &#39;2022-12-02&#39;, 23),
(1007, &#39;2022-12-02&#39;, 19),
(1003, &#39;2022-12-02&#39;, 30),
(1008, &#39;2022-12-03&#39;, 3),
(1009, &#39;2022-12-04&#39;, 1),
(1010, &#39;2022-12-05&#39;, 9),
(1003, &#39;2022-12-30&#39;, 121);

-- step1. 计算每种商品的总销售额，并降序排序
select
    good_category,
    sum(amount) as good_amount
from cm.tb_sale_amount
group by good_category
order by good_amount desc;

-- step2. 求全部商品的总销售额，为了step3求各种商品的占比，需要先求和。注意：求总和时，窗口值既不排序也不进行分组
select
    *,
    sum(good_amount) over () as all_amount
from
    (select
        good_category,
        sum(amount) as good_amount
    from cm.tb_sale_amount
    group by good_category
    order by good_amount desc
    ) t1;

-- step3. 求占比
select
    *,
    good_amount * 1.0 / all_amount as ratio
from (select
          *,
          sum(good_amount) over () as all_amount
      from (select
                good_category,
                sum(amount) as good_amount
            from cm.tb_sale_amount
            group by good_category
            order by good_amount desc
            ) as t1
      ) t2;

-- step4. 求累计占比，注意：求累计值时，一定要进行排序
select
    *,
    sum(ratio) over (order by ratio desc) as acc_ratio
from(
    select
        *,
        good_amount * 1.0 / all_amount as ratio
    from (select
              *,
              sum(good_amount) over () as all_amount
            from (select
                    good_category,
                    sum(amount) as good_amount
                from cm.tb_sale_amount
                group by good_category
                order by good_amount desc
            ) as t1
        ) t2
    ) t3;

-- step5. 求前一行的累计占比
select
    *,
    lag(acc_ratio) over() as pre_acc_ratio
from(select
         *,
       sum(ratio) over (order by ratio desc) as acc_ratio
    from(
        select
            *,
            good_amount * 1.0 / all_amount as ratio
        from (select
                  *,
                  sum(good_amount) over () as all_amount
              from (select
                        good_category,
                        sum(amount) as good_amount
                    from cm.tb_sale_amount
                    group by good_category
                    order by good_amount desc
                ) t1
            ) t2
        ) t3
    ) t4;

-- step6. 过滤
select *
from (
    select
    *,
    lag(acc_ratio) over() as pre_acc_ratio
from(select
         *,
       sum(ratio) over (order by ratio desc) as acc_ratio
     from(
        select
            *,
            good_amount * 1.0 / all_amount as ratio
        from (select
                  *,
                  sum(good_amount) over () as all_amount
              from (select
                        good_category,
                        sum(amount) as good_amount
                    from cm.tb_sale_amount
                    group by good_category
                    order by good_amount desc
                )  t1
              ) t2
        ) t3
    ) t4
) t5
where pre_acc_ratio IS NULL or pre_acc_ratio &lt; 0.90;
```

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/sql/90_percent_contribution_goods/  

