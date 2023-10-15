---
title: "Pandas系列之三：数据合并"
categories: ["编程"]
tags: ["Pandas"]
date: 2023-09-01
---
在使用Pandas处理并分析数据时，我们可能会载入多份数据集并保存在不同的变量中。而对这些不同的数据框或表进行连接合并则是常见的操作之一。在Pandas中，我们主要使用四种方式来进行数据的拼接：append、concat、merge和join。下面，笔者对这三种方法进行相应的总结。

### 一、append方法

该函数用于将另一个序列或数据框的行添加到现有的数据框中，函数定义及参数说明如下：

```python
def append(
	self, other, ignore_index=False, verify_integrity=False, sort=False
) -> "DateFrame":
```

**参数说明：**

other：需要追加的数据，为DataFrame或Series对象或由他们构成的列表。  
ignore_index：bool类型，默认为False。如果为True，则不使用索引标签。  
verify_integrity：bool类型，默认为False。如果为True，则当创建重复索引时抛出ValueError异常。  

sort：bool类型，默认为False。当被添加的数据列与目标列不同时进行排序。

**示例：**  

> In[24]: 

```python
df = pd.DateFrame([[1, 2], [3, 4]],columns=list('AB'))
df2 = pd.DateFrame([[5, 6], [7, 8]],columns=list('AB'))

## 保留索引
df.append(df2)
```

> Out[24]:

|序号|证券代码|证券简称|
|:---:|:---:|:---:|
|0|1|2|
|1|3|4|
|0|5|6|
|1|7|8|

#### 二、concat方法

concat（英文concatenate的缩写）是将两个表拼接在一起，因此需要指定拼接的轴，即按行拼接（axis=0）或按列拼接（axis=1），函数定义及参数说明如下：

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44nWaqbQKbBK5wMy0cBm0DsBTd4ibxOLo7ibn60fqAsicYsPMObzQFtsSjZQ/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

**参数说明：**

objs：Series或DataFrame对象的序列或映射，如[df1, df2, …]。

axis：拼接的方向，默认为0，即按行拼接（纵向/垂直拼接）。

join：如何处理其他轴上的索引，取值inner或outer，默认为outer。

ignore_index：bool类型，默认为False。如果为True，则结果轴将被标记为0,1,…,n-1。

keys：序列，默认为None。使用传递的键作为最外层构建层次索引。

levels：序列列表，默认值为None。用于构建多级索引的特定级别（唯一值），否则将从键推断。

names：列表，默认为None。结果层次索引中的级别的名称。

**示例：**  

（1）拼接序列

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44n0vzQgL2WN77wfN7o3CUjKeuE8ia19exJuQWCWoxHAkaEAzXTsPnVATg/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

（2）拼接数据框  

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44nwDIRibd2aXQA0z78gRnFJxdK6UdI5WAXxbEboVWz81v0KxzqkMTvLdA/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44n0klrpYcuzIS5TMrtiaIpmZlApVXwe71uYFwNzonibguhlfFCD81hdgTQ/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44ncLq20KLeEc1WxDV03xYWcU2tTF3nCelbhvyFESl159KO8nowjGhjKQ/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

#### 三、merge方法

如果读者熟悉SQL语言，那么使用merge函数将得心应手。merge操作类似于数据库语言中的join，可以将数据框或命名的序列对象按指定的key进行连接。函数定义及参数说明如下：

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44nncxCHJA875P21ptESj7KlJXKzqTm17Dvr3L4icuM33u044wp7vxVjcA/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

**参数说明：**

left：左侧DataFrame或Series对象。

right：右侧DataFrame或Series对象。

how：连接的方式，可选inner、outer、left、right、cross，默认为inner。

on：连接的列或索引级别名称，需要在左侧和右侧对象中找到。

left_on：左侧中的列或索引级别。

right_on：右侧中的列或索引级别。

left_index：如果为True，则使用左侧DataFrame中的索引（行标签）作为其连接键，默认为False。

right_index：与left_index含义相同。

sort：按字典顺序通过连接键对结果排序，默认为False。

suffixes：用于重叠列的字符串后缀元组。

copy：从传递的DataFrame对象复制数据，默认为True。

indicator：在输出的DataFrame中增加一列，名为_merge，包每行源的信息。

validate：字符串类型，默认为None，用于检查合并类型，包括one_to_one、one_to_many、one_to_many、one_to_many。

**示例：**

（1）内连接  

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44nHDErzicwZ2gNny26puz2vXxicIiaYL5gZU2wd2QNVzLLgTwadYtNNMRQg/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

（2）外连接

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44nhuCGRqroD1UbJSxEdAyyyRLeOicichvl6J44yPqDwnzX8rxicFuIcxP9w/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

#### 四、join方法

与merge类似，Pandas中的join函数也可以对两张表进行连接操作。函数定义及参数说明如下：

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44nZibjGdhLP51lZarSsZQZVPFRuZcAnnx8qTANv2yvr0wJ0vqSVEBFjHg/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

**参数说明：**

other：DataFrame、Series或DataFrame列表。

on：str或str列表类型，表示连接的键。注意：这里设置的是对象本身的列，连接的是other中的index，否则为index-on-index。

how：连接的类型，共有left、right、inner、outer四种，默认为left

lsuffix：str类型，默认为空字符串，左侧重复列的后缀。

rsuffix：与lsuffix相同，右侧重复列的后缀。

sort：bool类型，根据连接的键对结果数据框进行排序，默认为False。

**  
示例：**  

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44njNMD5dRSrXj6bk07UF9CrwEqh44lgV39p4eLPpNZ61T5ic0dwG0jXcw/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44n6NycomntGGn5oC4ZQibbiaayPrGmibv0icOqXc9BicnLAqEFVzSBmVCFzKQ/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

![图片](https://mmbiz.qpic.cn/mmbiz_png/6icvyDTK5FzB0XFkp0O9nW6apqwQyz44nWPeDa267jVjRChafO4sbbjSUhPXk2icicFh01cDicEMIibZnplDbUFHghA/640?wx_fmt=png&wxfrom=5&wx_lazy=1&wx_co=1)

上面介绍了append、concat、merge和join四种用于数据拼接的方法，它们的主要区别如下：

（1）append方法主要用于在DataFrame的末尾添加一行或多行数据，笔者常在循环体中动态增加结果集的行，此时使用append较多。

（2）concat方法通过指定axis参数实现水平（按列）或垂直（按行）方向的拼接。

（3）merge方法通过指定的列或行索引进行合并，为类方法，结果数据集的行或列均可能增加。

（4）join方法将相同行索引的数据合并在一起，为对象方法，结果数据集的行不会增加，列可能增加。