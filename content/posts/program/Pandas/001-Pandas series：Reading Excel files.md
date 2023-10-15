---
title: "Pandas系列之一：读取Excel文件"
categories: ["编程"]
tags: ["Pandas"]
date: 2023-09-01
---
在使用Python的Pandas包进行数据处理或分析时，读取数据源文件是第一步。本文结合笔者经验对Pandas读取xls或xlsx文件的方法进行了详细总结。

Pandas中的read_excel函数参数如下：
```python
def read_excel(
	io,
	sheet_name=0,
	header=0,
	names=None,
	index_col=None,
	usecols=None,
	squeeze=False,
	dtype=None,
	engine=None,
	converters=None,
	true_values=None,
	false_values=None,
	skiprows=None,
	nrows=None,
	na_values=None,
	keep_default_na=True,
	verbose=False,
	parse_dates=False,
	date_parser=None,
	thousands=None,
	comment=None,
	skipfooter=0,
	convert_float=True,
	mangle_dupe_cols=True,
	**kwds,
)
```

主要参数说明和示例如下：

### io

可传入文件路径字符串或ExcelFile/xlrd.Book对象。例如：

> In[8]: 
```python
## 文件路径字符串
df1 = pd.read_excel('1.xlsx', sheet_name='sheet1')

df1.head(1)
```

> Out[8]:

|序号|证券代码|证券简称|交易日期|日收盘价|
|---|---|---|---|---|
|0|688001|华兴源创|2020-01-02|46.05|

> In[9]: 
```python
## ExcelFile文件对象
xlsx = pd.ExcelFile('1.xlsx')
df2 = pd.read_excel(xlsx, sheet_name='sheet1')

df2.head(1)
```

> Out[9]:

|序号|证券代码|证券简称|交易日期|日收盘价|
|---|---|---|---|---|
|0|688001|华兴源创|2020-01-02|46.05|


### sheet_name

指定要读取的工作表，可接受str、int、list或None参数类型，默认为0，标识第一个Sheet。str为工作表名称；int为工作表索引，0标识第一个工作表；list用来获取多个工作表，返回的是字典，key为工作表名称或索引（取决于list中的元素类型），value为工作表数据；None标识获取所有工作表。例如：

> In[10]: 
```python
## 获取第一个工作表
df3 = pd.read_excel('1.xlsx', sheet_name=0)
## df3 = pd.read_excel('1.xlsx', sheet_name='sheet1')

df3.head(1)
```

> Out[10]:

|序号|证券代码|证券简称|交易日期|日收盘价|
|---|---|---|---|---|
|0|688001|华兴源创|2020-01-02|46.05|

> In[13]: 
```python
## 获取前三个工作表
df4 = pd.read_excel('1.xlsx', sheet_name=[0, 1, 2])
## df4 = pd.read_excel('1.xlsx', sheet_name=['sheet1', 'sheet2', 'sheet3'])

df4.head(1)
```

> Out[13]:

|序号|证券代码|证券简称|交易日期|日收盘价|
|---|---|---|---|---|
|0|688001|华兴源创|2020-01-02|46.05|

上述sheet_name参数也可以结合名称和索引一起使用，如[0,1,‘Sheet3’]。

### header

指定表头，可接受int、list of int类型。如果传入的为list，则所有传入的表头行构成MultiIndex。例如：

> In[20]: 
```python
df5 = pd.read_excel('1.xlsx', sheet_name='sheet2', header=[0 ,1])

df5.head(1)
```

> Out[20]:

||证券信息||日期|收盘价|
|---|---|---|---|---|
|序号|证券代码|证券简称|交易日期|日收盘价|
|0|688001|华兴源创|2020-01-02|46.05|

> In[24]: 
```python
df5['证券信息'].head(1)
```

> Out[24]:

|序号|证券代码|证券简称|
|---|---|---|
|0|688001|华兴源创|


上述代码指定第一行和第二行为标题行，构成MultiIndex。

### names

列表格式的列名，如果文件没有表头行，则需要显示传递header=None。

### index_col

index列，可接受int, list ofint或str类型参数，默认为None。在金融数据分析中，常常使用日期列作为index_col，从而进行时间序列处理，如index_col = ‘Date’，Date为字段名。如果传入的为list，则构成MultiIndex。

### usecols

需要解析的列名，接受int，str，list类型参数。如果不写该参数或为None，则默认解析所有列。如果为str，则可传入类似“A:E”,“A,C,E:F”这样格式的多列。常用的做法就是直接使用含有列名的列表。此外，该参数还接受callable类型值，解析返回为True的列。示例如下：

> In[22]: 
```python
## 仅解析含“证券”关键字的列
df6 = pd.read_excel('1.xlsx', usecols=lambda x:'证券' in x)

df6.head(1)
```

> Out[22]:

|序号|证券代码|证券简称|
|---|---|---|
|0|688001|华兴源创|
|1|688002|睿创微纳|
|2|688003|天准科技|
|3|688005|睿百科技|
|4|688006|杭百科技|

### squeeze

bool类型，默认为False。如果传入的数据仅有一列，则解析为Series。

### dtype

指定读取的列类型，接受字典格式，如：{'a': np.float64, 'b': np.int32}

> In[30]: 
```python
## 指定列的格式
df7 = pd.read_excel('1.xlsx', dtype={'证券代码': np.str, '日收盘价': np.float32})

df7.head(1)
```

> Out[30]:

|序号|证券代码|证券简称|交易日期|日收盘价|
|---|---|---|---|---|
|0|688001|华兴源创|2020-01-02|46.049999|
|1|688002|睿创微纳|2020-01-02|38.970001|
|2|688003|天准科技|2020-01-02|30.129999|
|3|688005|睿百科技|2020-01-02|33.680000|

### engine、converters、true_values, false_values

不常用的参数。

### skiprows

忽略的行数，int或list类型，序号从0开始。如：忽略第二行和第四行，则skiprows=[1,3]。

### nrows

解析的行数，int类型。当原始数据文件非常大时，可使用该参数读取部分行。

### na_values、keep_default_na、na_filter、verbose

非常用参数

### parse_dates

可接受bool、list或dict类型参数值，默认为False。当参数值为True时，解析index为日期格式；如果为list时，如[1,2,3]，则逐个解析第1,2,3列为日期格式。

> In[34]: 
```python
## 解析日期格式
df2 = pd.read_excel('1.xlsx', parse_dates=['交易日期'])

df1.head(1)
```

> Out[34]:

|序号|证券代码|证券简称|交易日期|日收盘价|
|---|---|---|---|---|
|0|688001|华兴源创|2020-01-02|46.05|
|1|688002|睿创微纳|2020-01-02|38.97|
|2|688003|天准科技|2020-01-02|30.13|

以上就是Pandas中read_excel函数的主要用法，掌握了基本规则后，剩下的就是各种灵活应用和处理了。