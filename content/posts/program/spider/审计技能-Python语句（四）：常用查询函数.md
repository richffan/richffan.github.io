---
title: 审计技能|Python语句（四）：常用查询函数
date: 2023-08-15 17:56:00
categories: ['审计']
tags: ['python']
---
## 1.分组聚合函数（groupby）

类似SQL的groupby那样的数据透视功能。可以使用pandas库的groupby函数来实现分组聚合操作。

举例：

有一个A表，分别有甲项目和乙项目，甲项目连续两年的销售为50万元和54万元，乙项目连续两年的销售额为74万元和56万元。

现在求甲项目和乙项目的各自销售额之和，具体如下：
代码：
```python
import pandas as pd

## 创建A表
data = {'项目名称': ['甲项目', '甲项目', '乙项目', '乙项目'], '销售额': [50, 54, 74, 56]}
df = pd.DataFrame(data)

## 按照“项目名称”进行聚合统计
grouped = df.groupby('项目名称')
result = grouped.sum()

## 打印结果
print(result)
```

运行结果：

|销售额|项目名称|
|---|---|
|乙项目|130|
|甲项目|104|



## 2.时间转换函数(to_datetime)

可以使用pandas库的to_datetime函数来实现时间转换操作。

举例：将分秒的时间转换为日期

| Date | Sales |
|---|---|
| 2020-01-01 10:30:00 | 1000 |
| 2020-01-02 15:20:00 | 1500 |
| 2020-01-03 18:40:00 | 2000 |

代码：
```python
import pandas as pd

data = pd.DataFrame({'Date': ['2020-01-01 10:30:00', '2020-01-02 15:20:00', '2020-01-03 18:40:00'],
'Sales': [1000, 1500, 2000]})

data['Date'] = pd.to_datetime(data['Date']).dt.date
print(data)
```

运行结果：

| Date | Sales |
|---|---|
|2020-01-01|1000|
|2020-01-02|1500|
|2020-01-03|2000|



## 3.分列函数（str.split）

在Python中，可以使用pandas库的str.split函数来实现分列操作。

举例：将地址列按照逗号分隔符进行拆分，并将拆分后的结果存储到城市列和省份列中。

数据示例：

| Customer | Address |
|---|---|
| A | Beijing,China |
| B | Shanghai,China |
| C | New York,United States |

代码：
```python
import pandas as pd
data = pd.DataFrame({
    'Customer': ['A', 'B', 'C'],
    'Address': ['Beijing,China', 'Shanghai,China', 'New York,United States']})
data[['City', 'Country/Region']] = data['Address'].str.split(',', expand=True)

print(data[['City', 'Country/Region']])
```

运行结果：

|City|Country/Region|
|---|---|
|0| Beijing/China|
|1|Shanghai/China|
|2|New York/United States|