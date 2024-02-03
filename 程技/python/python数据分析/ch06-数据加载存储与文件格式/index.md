# 第六章 数据加载、存储与文件格式


访问数据是使用本书所介绍的这些工具的第一步。我会着重介绍pandas的数据输入与输出，虽然别的库中也有不少以此为目的的工具。

输入输出通常可以划分为几个大类：读取文本文件和其他更高效的磁盘存储格式，加载数据库中的数据，利用Web API操作网络资源。

## 6.1 读写文本格式的数据
pandas提供了一些用于将表格型数据读取为DataFrame对象的函数。表6-1对它们进行了总结，其中read_csv和read_table可能会是你今后用得最多的。

![表6-1 pandas中的解析函数](https://gitee.com/wugenqiang/images/raw/master/02/1240-20201023230549415.png)

我将大致介绍一下这些函数在将文本数据转换为DataFrame时所用到的一些技术。这些函数的选项可以划分为以下几个大类：

- 索引：将一个或多个列当做返回的DataFrame处理，以及是否从文件、用户获取列名。
- 类型推断和数据转换：包括用户定义值的转换、和自定义的缺失值标记列表等。
- 日期解析：包括组合功能，比如将分散在多个列中的日期时间信息组合成结果中的单个列。
- 迭代：支持对大文件进行逐块迭代。
- 不规整数据问题：跳过一些行、页脚、注释或其他一些不重要的东西（比如由成千上万个逗号隔开的数值数据）。

因为工作中实际碰到的数据可能十分混乱，一些数据加载函数（尤其是read_csv）的选项逐渐变得复杂起来。面对不同的参数，感到头痛很正常（read_csv有超过50个参数）。pandas文档有这些参数的例子，如果你感到阅读某个文件很难，可以通过相似的足够多的例子找到正确的参数。

其中一些函数，比如pandas.read_csv，有类型推断功能，因为列数据的类型不属于数据类型。也就是说，你不需要指定列的类型到底是数值、整数、布尔值，还是字符串。其它的数据格式，如HDF5、Feather和msgpack，会在格式中存储数据类型。

日期和其他自定义类型的处理需要多花点工夫才行。首先我们来看一个以逗号分隔的（CSV）文本文件：
```python
In [8]: !cat examples/ex1.csv
a,b,c,d,message
1,2,3,4,hello
5,6,7,8,world
9,10,11,12,foo
```

&gt;笔记：这里，我用的是Unix的cat shell命令将文件的原始内容打印到屏幕上。如果你用的是Windows，你可以使用type达到同样的效果。

由于该文件以逗号分隔，所以我们可以使用read_csv将其读入一个DataFrame：
```python
In [9]: df = pd.read_csv(&#39;examples/ex1.csv&#39;)

In [10]: df
Out[10]: 
   a   b   c   d message
0  1   2   3   4   hello
1  5   6   7   8   world
2  9  10  11  12     foo
```

我们还可以使用read_table，并指定分隔符：
```python
In [11]: pd.read_table(&#39;examples/ex1.csv&#39;, sep=&#39;,&#39;)
Out[11]: 
   a   b   c   d message
0  1   2   3   4   hello
1  5   6   7   8   world
2  9  10  11  12     foo
```

并不是所有文件都有标题行。看看下面这个文件：
```python
In [12]: !cat examples/ex2.csv
1,2,3,4,hello
5,6,7,8,world
9,10,11,12,foo
```

读入该文件的办法有两个。你可以让pandas为其分配默认的列名，也可以自己定义列名：
```python
In [13]: pd.read_csv(&#39;examples/ex2.csv&#39;, header=None)
Out[13]: 
   0   1   2   3      4
0  1   2   3   4  hello
1  5   6   7   8  world
2  9  10  11  12    foo

In [14]: pd.read_csv(&#39;examples/ex2.csv&#39;, names=[&#39;a&#39;, &#39;b&#39;, &#39;c&#39;, &#39;d&#39;, &#39;message&#39;])
Out[14]: 
   a   b   c   d message
0  1   2   3   4   hello
1  5   6   7   8   world
2  9  10  11  12     foo
```

假设你希望将message列做成DataFrame的索引。你可以明确表示要将该列放到索引4的位置上，也可以通过index_col参数指定&#34;message&#34;：
```python
In [15]: names = [&#39;a&#39;, &#39;b&#39;, &#39;c&#39;, &#39;d&#39;, &#39;message&#39;]

In [16]: pd.read_csv(&#39;examples/ex2.csv&#39;, names=names, index_col=&#39;message&#39;)
Out[16]: 
         a   b   c   d
message               
hello    1   2   3   4
world    5   6   7   8
foo      9  10  11  12
```

如果希望将多个列做成一个层次化索引，只需传入由列编号或列名组成的列表即可：
```python
In [17]: !cat examples/csv_mindex.csv
key1,key2,value1,value2
one,a,1,2
one,b,3,4
one,c,5,6
one,d,7,8
two,a,9,10
two,b,11,12
two,c,13,14
two,d,15,16

In [18]: parsed = pd.read_csv(&#39;examples/csv_mindex.csv&#39;,
   ....:                      index_col=[&#39;key1&#39;, &#39;key2&#39;])

In [19]: parsed
Out[19]: 
           value1  value2
key1 key2                
one  a          1       2
     b          3       4
     c          5       6
     d          7       8
two  a          9      10
     b         11      12
     c         13      14
     d         15      16
```

有些情况下，有些表格可能不是用固定的分隔符去分隔字段的（比如空白符或其它模式）。看看下面这个文本文件：
```python
In [20]: list(open(&#39;examples/ex3.txt&#39;))
Out[20]: 
[&#39;            A         B         C\n&#39;,
 &#39;aaa -0.264438 -1.026059 -0.619500\n&#39;,
 &#39;bbb  0.927272  0.302904 -0.032399\n&#39;,
 &#39;ccc -0.264273 -0.386314 -0.217601\n&#39;,
 &#39;ddd -0.871858 -0.348382  1.100491\n&#39;]
```

虽然可以手动对数据进行规整，这里的字段是被数量不同的空白字符间隔开的。这种情况下，你可以传递一个正则表达式作为read_table的分隔符。可以用正则表达式表达为\s&#43;，于是有：
```python
In [21]: result = pd.read_table(&#39;examples/ex3.txt&#39;, sep=&#39;\s&#43;&#39;)

In [22]: result
Out[22]: 
            A         B         C
aaa -0.264438 -1.026059 -0.619500
bbb  0.927272  0.302904 -0.032399
ccc -0.264273 -0.386314 -0.217601
ddd -0.871858 -0.348382  1.100491
```

这里，由于列名比数据行的数量少，所以read_table推断第一列应该是DataFrame的索引。

这些解析器函数还有许多参数可以帮助你处理各种各样的异形文件格式（表6-2列出了一些）。比如说，你可以用skiprows跳过文件的第一行、第三行和第四行：
```python
In [23]: !cat examples/ex4.csv
# hey!
a,b,c,d,message
# just wanted to make things more difficult for you
# who reads CSV files with computers, anyway?
1,2,3,4,hello
5,6,7,8,world
9,10,11,12,foo
In [24]: pd.read_csv(&#39;examples/ex4.csv&#39;, skiprows=[0, 2, 3])
Out[24]: 
   a   b   c   d message
0  1   2   3   4   hello
1  5   6   7   8   world
2  9  10  11  12     foo
```

缺失值处理是文件解析任务中的一个重要组成部分。缺失数据经常是要么没有（空字符串），要么用某个标记值表示。默认情况下，pandas会用一组经常出现的标记值进行识别，比如NA及NULL：
```python
In [25]: !cat examples/ex5.csv
something,a,b,c,d,message
one,1,2,3,4,NA
two,5,6,,8,world
three,9,10,11,12,foo
In [26]: result = pd.read_csv(&#39;examples/ex5.csv&#39;)

In [27]: result
Out[27]: 
  something  a   b     c   d message
0       one  1   2   3.0   4     NaN
1       two  5   6   NaN   8   world
2     three  9  10  11.0  12     foo

In [28]: pd.isnull(result)
Out[28]: 
   something      a      b      c      d  message
0      False  False  False  False  False     True
1      False  False  False   True  False    False
2      False  False  False  False  False    False
```

na_values可以用一个列表或集合的字符串表示缺失值：
```python
In [29]: result = pd.read_csv(&#39;examples/ex5.csv&#39;, na_values=[&#39;NULL&#39;])

In [30]: result
Out[30]: 
  something  a   b     c   d message
0       one  1   2   3.0   4     NaN
1       two  5   6   NaN   8   world
2     three  9  10  11.0  12     foo
```

字典的各列可以使用不同的NA标记值：
```python
In [31]: sentinels = {&#39;message&#39;: [&#39;foo&#39;, &#39;NA&#39;], &#39;something&#39;: [&#39;two&#39;]}

In [32]: pd.read_csv(&#39;examples/ex5.csv&#39;, na_values=sentinels)
Out[32]:
something  a   b     c   d message
0       one  1   2   3.0   4     NaN
1       NaN  5   6   NaN   8   world
2     three  9  10  11.0  12     NaN
```

表6-2列出了pandas.read_csv和pandas.read_table常用的选项。

![](https://gitee.com/wugenqiang/images/raw/master/02/1240-20201024084538674.png)

![](https://gitee.com/wugenqiang/images/raw/master/02/1240-20201024084606416.png)

![](https://gitee.com/wugenqiang/images/raw/master/02/1240-20201024084613844.png)

### 6.1.1 逐块读取文本文件
在处理很大的文件时，或找出大文件中的参数集以便于后续处理时，你可能只想读取文件的一小部分或逐块对文件进行迭代。

在看大文件之前，我们先设置pandas显示地更紧些：
```python
In [33]: pd.options.display.max_rows = 10
```

然后有：
```python
In [34]: result = pd.read_csv(&#39;examples/ex6.csv&#39;)

In [35]: result
Out[35]: 
           one       two     three      four key
0     0.467976 -0.038649 -0.295344 -1.824726   L
1    -0.358893  1.404453  0.704965 -0.200638   B
2    -0.501840  0.659254 -0.421691 -0.057688   G
3     0.204886  1.074134  1.388361 -0.982404   R
4     0.354628 -0.133116  0.283763 -0.837063   Q
...        ...       ...       ...       ...  ..
9995  2.311896 -0.417070 -1.409599 -0.515821   L
9996 -0.479893 -0.650419  0.745152 -0.646038   E
9997  0.523331  0.787112  0.486066  1.093156   K
9998 -0.362559  0.598894 -1.843201  0.887292   G
9999 -0.096376 -1.012999 -0.657431 -0.573315   0
[10000 rows x 5 columns]
If you want to only read a small
```

如果只想读取几行（避免读取整个文件），通过nrows进行指定即可：
```python
In [36]: pd.read_csv(&#39;examples/ex6.csv&#39;, nrows=5)
Out[36]: 
        one       two     three      four key
0  0.467976 -0.038649 -0.295344 -1.824726   L
1 -0.358893  1.404453  0.704965 -0.200638   B
2 -0.501840  0.659254 -0.421691 -0.057688   G
3  0.204886  1.074134  1.388361 -0.982404   R
4  0.354628 -0.133116  0.283763 -0.837063   Q
```

要逐块读取文件，可以指定chunksize（行数）：
```python
In [874]: chunker = pd.read_csv(&#39;ch06/ex6.csv&#39;, chunksize=1000)

In [875]: chunker
Out[875]: &lt;pandas.io.parsers.TextParser at 0x8398150&gt;
```

read_csv所返回的这个TextParser对象使你可以根据chunksize对文件进行逐块迭代。比如说，我们可以迭代处理ex6.csv，将值计数聚合到&#34;key&#34;列中，如下所示：
```python
chunker = pd.read_csv(&#39;examples/ex6.csv&#39;, chunksize=1000)

tot = pd.Series([])
for piece in chunker:
    tot = tot.add(piece[&#39;key&#39;].value_counts(), fill_value=0)

tot = tot.sort_values(ascending=False)
```

然后有：
```python
In [40]: tot[:10]
Out[40]: 
E    368.0
X    364.0
L    346.0
O    343.0
Q    340.0
M    338.0
J    337.0
F    335.0
K    334.0
H    330.0
dtype: float64
```

TextParser还有一个get_chunk方法，它使你可以读取任意大小的块。

### 6.1.2 将数据写出到文本格式
数据也可以被输出为`分隔符`格式的文本。我们再来看看之前读过的一个CSV文件：
```python
In [41]: data = pd.read_csv(&#39;examples/ex5.csv&#39;)

In [42]: data
Out[42]: 
  something  a   b     c   d message
0       one  1   2   3.0   4     NaN
1       two  5   6   NaN   8   world
2     three  9  10  11.0  12     foo
```

利用DataFrame的`to_csv`方法，我们可以将数据写到一个以`逗号`分隔的文件中：
```python
In [43]: data.to_csv(&#39;examples/out.csv&#39;)

In [44]: !cat examples/out.csv
,something,a,b,c,d,message
0,one,1,2,3.0,4,
1,two,5,6,,8,world
2,three,9,10,11.0,12,foo
```

当然，还可以使用其他分隔符（由于这里直接写出到sys.stdout，所以仅仅是打印出文本结果而已）：
```python
In [45]: import sys

In [46]: data.to_csv(sys.stdout, sep=&#39;|&#39;)
|something|a|b|c|d|message
0|one|1|2|3.0|4|
1|two|5|6||8|world
2|three|9|10|11.0|12|foo
```

缺失值在输出结果中会被表示为空字符串。你可能希望将其表示为别的标记值：
```python
In [47]: data.to_csv(sys.stdout, na_rep=&#39;NULL&#39;)
,something,a,b,c,d,message
0,one,1,2,3.0,4,NULL
1,two,5,6,NULL,8,world
2,three,9,10,11.0,12,foo
```

如果没有设置其他选项，则会写出行和列的标签。当然，它们也都可以被禁用：
```python
In [48]: data.to_csv(sys.stdout, index=False, header=False)
one,1,2,3.0,4,
two,5,6,,8,world
three,9,10,11.0,12,foo
```

此外，你还可以只写出一部分的列，并以你指定的顺序排列：
```python
In [49]: data.to_csv(sys.stdout, index=False, columns=[&#39;a&#39;, &#39;b&#39;, &#39;c&#39;])
a,b,c
1,2,3.0
5,6,
9,10,11.0
```

Series也有一个to_csv方法：
```python
In [50]: dates = pd.date_range(&#39;1/1/2000&#39;, periods=7)

In [51]: ts = pd.Series(np.arange(7), index=dates)

In [52]: ts.to_csv(&#39;examples/tseries.csv&#39;)

In [53]: !cat examples/tseries.csv
2000-01-01,0
2000-01-02,1
2000-01-03,2
2000-01-04,3
2000-01-05,4
2000-01-06,5
2000-01-07,6
```

### 6.1.3 处理分隔符格式
大部分存储在磁盘上的表格型数据都能用pandas.read_table进行加载。然而，有时还是需要做一些手工处理。由于接收到含有畸形行的文件而使read_table出毛病的情况并不少见。为了说明这些基本工具，看看下面这个简单的CSV文件：
```python
In [54]: !cat examples/ex7.csv
&#34;a&#34;,&#34;b&#34;,&#34;c&#34;
&#34;1&#34;,&#34;2&#34;,&#34;3&#34;
&#34;1&#34;,&#34;2&#34;,&#34;3&#34;
```

对于**任何单字符分隔符文件**，可以直接使用Python内置的`csv`模块。将任意已打开的文件或文件型的对象传给csv.reader：
```python
import csv
f = open(&#39;examples/ex7.csv&#39;)

reader = csv.reader(f)
```

对这个reader进行迭代将会为每行产生一个元组（并移除了所有的引号）：对这个reader进行迭代将会为每行产生一个元组（并移除了所有的引号）：
```python
In [56]: for line in reader:
   ....:     print(line)
[&#39;a&#39;, &#39;b&#39;, &#39;c&#39;]
[&#39;1&#39;, &#39;2&#39;, &#39;3&#39;]
[&#39;1&#39;, &#39;2&#39;, &#39;3&#39;]
```

现在，为了使数据格式合乎要求，你需要对其做一些整理工作。我们一步一步来做。首先，读取文件到一个多行的列表中：
```python
In [57]: with open(&#39;examples/ex7.csv&#39;) as f:
   ....:     lines = list(csv.reader(f))
```

然后，我们将这些行分为`标题行`和`数据行`：
```python
In [58]: header, values = lines[0], lines[1:]
```

然后，我们可以用字典构造式和zip(*values)，后者将行转置为列，创建数据列的字典：
```python
In [59]: data_dict = {h: v for h, v in zip(header, zip(*values))}

In [60]: data_dict
Out[60]: {&#39;a&#39;: (&#39;1&#39;, &#39;1&#39;), &#39;b&#39;: (&#39;2&#39;, &#39;2&#39;), &#39;c&#39;: (&#39;3&#39;, &#39;3&#39;)}
```

CSV文件的形式有很多。只需定义csv.Dialect的一个子类即可定义出新格式（如专门的分隔符、字符串引用约定、行结束符等）：
```python
class my_dialect(csv.Dialect):
    lineterminator = &#39;\n&#39;
    delimiter = &#39;;&#39;
    quotechar = &#39;&#34;&#39;
    quoting = csv.QUOTE_MINIMAL
reader = csv.reader(f, dialect=my_dialect)
```

各个CSV语支的参数也可以用关键字的形式提供给csv.reader，而无需定义子类：
```python
reader = csv.reader(f, delimiter=&#39;|&#39;)
```

可用的选项（csv.Dialect的属性）及其功能如表6-3所示。

![](https://gitee.com/wugenqiang/images/raw/master/02/1240-20201024092934346.png)

&gt;笔记：对于那些使用复杂分隔符或多字符分隔符的文件，csv模块就无能为力了。这种情况下，你就只能使用字符串的split方法或正则表达式方法re.split进行行拆分和其他整理工作了。

要手工输出分隔符文件，你可以使用csv.writer。它接受一个已打开且可写的文件对象以及跟csv.reader相同的那些语支和格式化选项：
```python
with open(&#39;mydata.csv&#39;, &#39;w&#39;) as f:
    writer = csv.writer(f, dialect=my_dialect)
    writer.writerow((&#39;one&#39;, &#39;two&#39;, &#39;three&#39;))
    writer.writerow((&#39;1&#39;, &#39;2&#39;, &#39;3&#39;))
    writer.writerow((&#39;4&#39;, &#39;5&#39;, &#39;6&#39;))
    writer.writerow((&#39;7&#39;, &#39;8&#39;, &#39;9&#39;))
```

### 6.1.4 JSON数据
JSON（JavaScript Object Notation的简称）已经成为通过HTTP请求在Web浏览器和其他应用程序之间发送数据的标准格式之一。它是一种比表格型文本格式（如CSV）灵活得多的数据格式。下面是一个例子：
```python
obj = &#34;&#34;&#34;
{&#34;name&#34;: &#34;Wes&#34;,
 &#34;places_lived&#34;: [&#34;United States&#34;, &#34;Spain&#34;, &#34;Germany&#34;],
 &#34;pet&#34;: null,
 &#34;siblings&#34;: [{&#34;name&#34;: &#34;Scott&#34;, &#34;age&#34;: 30, &#34;pets&#34;: [&#34;Zeus&#34;, &#34;Zuko&#34;]},
              {&#34;name&#34;: &#34;Katie&#34;, &#34;age&#34;: 38,
               &#34;pets&#34;: [&#34;Sixes&#34;, &#34;Stache&#34;, &#34;Cisco&#34;]}]
}
&#34;&#34;&#34;
```
除其空值null和一些其他的细微差别（如列表末尾不允许存在多余的逗号）之外，JSON非常接近于有效的Python代码。基本类型有对象（字典）、数组（列表）、字符串、数值、布尔值以及null。对象中所有的键都必须是字符串。许多Python库都可以读写JSON数据。我将使用json，因为它是构建于Python标准库中的。通过`json.loads`即可将JSON字符串转换成Python形式：
```python
In [62]: import json

In [63]: result = json.loads(obj)

In [64]: result
Out[64]: 
{&#39;name&#39;: &#39;Wes&#39;,
 &#39;pet&#39;: None,
 &#39;places_lived&#39;: [&#39;United States&#39;, &#39;Spain&#39;, &#39;Germany&#39;],
 &#39;siblings&#39;: [{&#39;age&#39;: 30, &#39;name&#39;: &#39;Scott&#39;, &#39;pets&#39;: [&#39;Zeus&#39;, &#39;Zuko&#39;]},
  {&#39;age&#39;: 38, &#39;name&#39;: &#39;Katie&#39;, &#39;pets&#39;: [&#39;Sixes&#39;, &#39;Stache&#39;, &#39;Cisco&#39;]}]}
```

json.dumps则将Python对象转换成JSON格式：
```python
In [65]: asjson = json.dumps(result)
```

如何将（一个或一组）JSON对象转换为DataFrame或其他便于分析的数据结构就由你决定了。最简单方便的方式是：向DataFrame构造器传入一个字典的列表（就是原先的JSON对象），并选取数据字段的子集：
```python
In [66]: siblings = pd.DataFrame(result[&#39;siblings&#39;], columns=[&#39;name&#39;, &#39;age&#39;])

In [67]: siblings
Out[67]: 
    name  age
0  Scott   30
1  Katie   38
```

pandas.read_json可以自动将特别格式的JSON数据集转换为Series或DataFrame。例如：
```python
In [68]: !cat examples/example.json
[{&#34;a&#34;: 1, &#34;b&#34;: 2, &#34;c&#34;: 3},
 {&#34;a&#34;: 4, &#34;b&#34;: 5, &#34;c&#34;: 6},
 {&#34;a&#34;: 7, &#34;b&#34;: 8, &#34;c&#34;: 9}]
```

pandas.read_json的默认选项假设JSON数组中的每个对象是表格中的一行：
```python
In [69]: data = pd.read_json(&#39;examples/example.json&#39;)

In [70]: data
Out[70]: 
   a  b  c
0  1  2  3
1  4  5  6
2  7  8  9
```

第7章中关于USDA Food Database的那个例子进一步讲解了JSON数据的读取和处理（包括嵌套记录）。

如果你需要将数据从pandas输出到JSON，可以使用to_json方法：
```python
In [71]: print(data.to_json())
{&#34;a&#34;:{&#34;0&#34;:1,&#34;1&#34;:4,&#34;2&#34;:7},&#34;b&#34;:{&#34;0&#34;:2,&#34;1&#34;:5,&#34;2&#34;:8},&#34;c&#34;:{&#34;0&#34;:3,&#34;1&#34;:6,&#34;2&#34;:9}}

In [72]: print(data.to_json(orient=&#39;records&#39;))
[{&#34;a&#34;:1,&#34;b&#34;:2,&#34;c&#34;:3},{&#34;a&#34;:4,&#34;b&#34;:5,&#34;c&#34;:6},{&#34;a&#34;:7,&#34;b&#34;:8,&#34;c&#34;:9}]
```

### 6.1.5 XML和HTML：Web信息收集

Python有许多可以读写常见的HTML和XML格式数据的库，包括lxml、Beautiful Soup和html5lib。lxml的速度比较快，但其它的库处理有误的HTML或XML文件更好。

pandas有一个内置的功能，read_html，它可以使用lxml和Beautiful Soup自动将HTML文件中的表格解析为DataFrame对象。为了进行展示，我从美国联邦存款保险公司下载了一个HTML文件（pandas文档中也使用过），它记录了银行倒闭的情况。首先，你需要安装read_html用到的库：
```
conda install lxml
pip install beautifulsoup4 html5lib
```

如果你用的不是conda，可以使用``pip install lxml``。

pandas.read_html有一些选项，默认条件下，它会搜索、尝试解析&lt;table&gt;标签内的的表格数据。结果是一个列表的DataFrame对象：
```python
In [73]: tables = pd.read_html(&#39;examples/fdic_failed_bank_list.html&#39;)

In [74]: len(tables)
Out[74]: 1

In [75]: failures = tables[0]

In [76]: failures.head()
Out[76]: 
                      Bank Name             City  ST   CERT  \
0                   Allied Bank         Mulberry  AR     91   
1  The Woodbury Banking Company         Woodbury  GA  11297   
2        First CornerStone Bank  King of Prussia  PA  35312   
3            Trust Company Bank          Memphis  TN   9956   
4    North Milwaukee State Bank        Milwaukee  WI  20364   
                 Acquiring Institution        Closing Date       Updated Date  
0                         Today&#39;s Bank  September 23, 2016  November 17, 2016  
1                          United Bank     August 19, 2016  November 17, 2016  
2  First-Citizens Bank &amp; Trust Company         May 6, 2016  September 6, 2016  
3           The Bank of Fayette County      April 29, 2016  September 6, 2016  
4  First-Citizens Bank &amp; Trust Company      March 11, 2016      June 16, 2016
```

因为failures有许多列，pandas插入了一个换行符\。

这里，我们可以做一些数据清洗和分析（后面章节会进一步讲解），比如计算按年份计算倒闭的银行数：
```python
In [77]: close_timestamps = pd.to_datetime(failures[&#39;Closing Date&#39;])

In [78]: close_timestamps.dt.year.value_counts()
Out[78]: 
2010    157
2009    140
2011     92
2012     51
2008     25
       ... 
2004      4
2001      4
2007      3
2003      3
2000      2
Name: Closing Date, Length: 15, dtype: int64
```

### 6.1.6 利用lxml.objectify解析XML
XML（Extensible Markup Language）是另一种常见的支持分层、嵌套数据以及元数据的结构化数据格式。本书所使用的这些文件实际上来自于一个很大的XML文档。

前面，我介绍了pandas.read_html函数，它可以使用lxml或Beautiful Soup从HTML解析数据。XML和HTML的结构很相似，但XML更为通用。这里，我会用一个例子演示如何利用lxml从XML格式解析数据。

纽约大都会运输署发布了一些有关其公交和列车服务的数据资料（http://www.mta.info/developers/download.html）。这里，我们将看看包含在一组XML文件中的运行情况数据。每项列车或公交服务都有各自的文件（如Metro-North Railroad的文件是Performance_MNR.xml），其中每条XML记录就是一条月度数据，如下所示：
```xml
&lt;INDICATOR&gt;
  &lt;INDICATOR_SEQ&gt;373889&lt;/INDICATOR_SEQ&gt;
  &lt;PARENT_SEQ&gt;&lt;/PARENT_SEQ&gt;
  &lt;AGENCY_NAME&gt;Metro-North Railroad&lt;/AGENCY_NAME&gt;
  &lt;INDICATOR_NAME&gt;Escalator Availability&lt;/INDICATOR_NAME&gt;
  &lt;DESCRIPTION&gt;Percent of the time that escalators are operational
  systemwide. The availability rate is based on physical observations performed
  the morning of regular business days only. This is a new indicator the agency
  began reporting in 2009.&lt;/DESCRIPTION&gt;
  &lt;PERIOD_YEAR&gt;2011&lt;/PERIOD_YEAR&gt;
  &lt;PERIOD_MONTH&gt;12&lt;/PERIOD_MONTH&gt;
  &lt;CATEGORY&gt;Service Indicators&lt;/CATEGORY&gt;
  &lt;FREQUENCY&gt;M&lt;/FREQUENCY&gt;
  &lt;DESIRED_CHANGE&gt;U&lt;/DESIRED_CHANGE&gt;
  &lt;INDICATOR_UNIT&gt;%&lt;/INDICATOR_UNIT&gt;
  &lt;DECIMAL_PLACES&gt;1&lt;/DECIMAL_PLACES&gt;
  &lt;YTD_TARGET&gt;97.00&lt;/YTD_TARGET&gt;
  &lt;YTD_ACTUAL&gt;&lt;/YTD_ACTUAL&gt;
  &lt;MONTHLY_TARGET&gt;97.00&lt;/MONTHLY_TARGET&gt;
  &lt;MONTHLY_ACTUAL&gt;&lt;/MONTHLY_ACTUAL&gt;
&lt;/INDICATOR&gt;
```

我们先用lxml.objectify解析该文件，然后通过getroot得到该XML文件的根节点的引用：
```python
from lxml import objectify

path = &#39;datasets/mta_perf/Performance_MNR.xml&#39;
parsed = objectify.parse(open(path))
root = parsed.getroot()
```

root.INDICATOR返回一个用于产生各个&lt;INDICATOR&gt;XML元素的生成器。对于每条记录，我们可以用标记名（如YTD_ACTUAL）和数据值填充一个字典（排除几个标记）：
```python
data = []

skip_fields = [&#39;PARENT_SEQ&#39;, &#39;INDICATOR_SEQ&#39;,
               &#39;DESIRED_CHANGE&#39;, &#39;DECIMAL_PLACES&#39;]

for elt in root.INDICATOR:
    el_data = {}
    for child in elt.getchildren():
        if child.tag in skip_fields:
            continue
        el_data[child.tag] = child.pyval
    data.append(el_data)
```

最后，将这组字典转换为一个DataFrame：
```python
In [81]: perf = pd.DataFrame(data)

In [82]: perf.head()
Out[82]:
Empty DataFrame
Columns: []
Index: []
```

XML数据可以比本例复杂得多。每个标记都可以有元数据。看看下面这个HTML的链接标签（它也算是一段有效的XML）：
```python
from io import StringIO
tag = &#39;&lt;a href=&#34;http://www.google.com&#34;&gt;Google&lt;/a&gt;&#39;
root = objectify.parse(StringIO(tag)).getroot()
```

现在就可以访问标签或链接文本中的任何字段了（如href）：
```python
In [84]: root
Out[84]: &lt;Element a at 0x7f6b15817748&gt;

In [85]: root.get(&#39;href&#39;)
Out[85]: &#39;http://www.google.com&#39;

In [86]: root.text
Out[86]: &#39;Google&#39;
```

## 6.2 二进制数据格式

实现数据的高效二进制格式存储最简单的办法之一是使用Python内置的pickle序列化。pandas对象都有一个用于将数据以pickle格式保存到磁盘上的to_pickle方法：
```python
In [87]: frame = pd.read_csv(&#39;examples/ex1.csv&#39;)

In [88]: frame
Out[88]: 
   a   b   c   d message
0  1   2   3   4   hello
1  5   6   7   8   world
2  9  10  11  12     foo

In [89]: frame.to_pickle(&#39;examples/frame_pickle&#39;)
```

你可以通过pickle直接读取被pickle化的数据，或是使用更为方便的pandas.read_pickle：
```python
In [90]: pd.read_pickle(&#39;examples/frame_pickle&#39;)
Out[90]: 
   a   b   c   d message
0  1   2   3   4   hello
1  5   6   7   8   world
2  9  10  11  12     foo
```

&gt;注意：pickle仅建议用于短期存储格式。其原因是很难保证该格式永远是稳定的；今天pickle的对象可能无法被后续版本的库unpickle出来。虽然我尽力保证这种事情不会发生在pandas中，但是今后的某个时候说不定还是得“打破”该pickle格式。

pandas内置支持两个二进制数据格式：HDF5和MessagePack。下一节，我会给出几个HDF5的例子，但我建议你尝试下不同的文件格式，看看它们的速度以及是否适合你的分析工作。pandas或NumPy数据的其它存储格式有：

- bcolz：一种可压缩的列存储二进制格式，基于Blosc压缩库。
- Feather：我与R语言社区的Hadley Wickham设计的一种跨语言的列存储文件格式。Feather使用了Apache Arrow的列式内存格式。

### 6.2.1 使用HDF5格式

HDF5是一种存储大规模科学数组数据的非常好的文件格式。它可以被作为C标准库，带有许多语言的接口，如Java、Python和MATLAB等。HDF5中的HDF指的是层次型数据格式（hierarchical data format）。每个HDF5文件都含有一个文件系统式的节点结构，它使你能够存储多个数据集并支持元数据。与其他简单格式相比，HDF5支持多种压缩器的即时压缩，还能更高效地存储重复模式数据。对于那些非常大的无法直接放入内存的数据集，HDF5就是不错的选择，因为它可以高效地分块读写。

虽然可以用PyTables或h5py库直接访问HDF5文件，pandas提供了更为高级的接口，可以简化存储Series和DataFrame对象。HDFStore类可以像字典一样，处理低级的细节：
```python
In [92]: frame = pd.DataFrame({&#39;a&#39;: np.random.randn(100)})

In [93]: store = pd.HDFStore(&#39;mydata.h5&#39;)

In [94]: store[&#39;obj1&#39;] = frame

In [95]: store[&#39;obj1_col&#39;] = frame[&#39;a&#39;]

In [96]: store
Out[96]: 
&lt;class &#39;pandas.io.pytables.HDFStore&#39;&gt;
File path: mydata.h5
/obj1                frame        (shape-&gt;[100,1])                               
        
/obj1_col            series       (shape-&gt;[100])                                 
        
/obj2                frame_table  (typ-&gt;appendable,nrows-&gt;100,ncols-&gt;1,indexers-&gt;
[index])
/obj3                frame_table  (typ-&gt;appendable,nrows-&gt;100,ncols-&gt;1,indexers-&gt;
[index])
```

HDF5文件中的对象可以通过与字典一样的API进行获取：
```python
In [97]: store[&#39;obj1&#39;]
Out[97]: 
           a
0  -0.204708
1   0.478943
2  -0.519439
3  -0.555730
4   1.965781
..       ...
95  0.795253
96  0.118110
97 -0.748532
98  0.584970
99  0.152677
[100 rows x 1 columns]
```

HDFStore支持两种存储模式，&#39;fixed&#39;和&#39;table&#39;。后者通常会更慢，但是支持使用特殊语法进行查询操作：
```python
In [98]: store.put(&#39;obj2&#39;, frame, format=&#39;table&#39;)

In [99]: store.select(&#39;obj2&#39;, where=[&#39;index &gt;= 10 and index &lt;= 15&#39;])
Out[99]: 
           a
10  1.007189
11 -1.296221
12  0.274992
13  0.228913
14  1.352917
15  0.886429

In [100]: store.close()
```

put是store[&#39;obj2&#39;] = frame方法的显示版本，允许我们设置其它的选项，比如格式。

pandas.read_hdf函数可以快捷使用这些工具：
```python
In [101]: frame.to_hdf(&#39;mydata.h5&#39;, &#39;obj3&#39;, format=&#39;table&#39;)

In [102]: pd.read_hdf(&#39;mydata.h5&#39;, &#39;obj3&#39;, where=[&#39;index &lt; 5&#39;])
Out[102]: 
          a
0 -0.204708
1  0.478943
2 -0.519439
3 -0.555730
4  1.965781
```

&gt;笔记：如果你要处理的数据位于远程服务器，比如Amazon S3或HDFS，使用专门为分布式存储（比如Apache Parquet）的二进制格式也许更加合适。Python的Parquet和其它存储格式还在不断的发展之中，所以这本书中没有涉及。

如果需要本地处理海量数据，我建议你好好研究一下PyTables和h5py，看看它们能满足你的哪些需求。。由于许多数据分析问题都是IO密集型（而不是CPU密集型），利用HDF5这样的工具能显著提升应用程序的效率。

&gt;注意：HDF5不是数据库。它最适合用作“一次写多次读”的数据集。虽然数据可以在任何时候被添加到文件中，但如果同时发生多个写操作，文件就可能会被破坏。

### 6.2.2 读取Microsoft Excel文件

pandas的ExcelFile类或pandas.read_excel函数支持读取存储在Excel 2003（或更高版本）中的表格型数据。这两个工具分别使用扩展包xlrd和openpyxl读取XLS和XLSX文件。你可以用pip或conda安装它们。

要使用ExcelFile，通过传递xls或xlsx路径创建一个实例：
```python
In [104]: xlsx = pd.ExcelFile(&#39;examples/ex1.xlsx&#39;)
```

存储在表单中的数据可以read_excel读取到DataFrame（原书这里写的是用parse解析，但代码中用的是read_excel，是个笔误：只换了代码，没有改文字）：
```python
In [105]: pd.read_excel(xlsx, &#39;Sheet1&#39;)
Out[105]: 
   a   b   c   d message
0  1   2   3   4   hello
1  5   6   7   8   world
2  9  10  11  12     foo
```

如果要读取一个文件中的多个表单，创建ExcelFile会更快，但你也可以将文件名传递到pandas.read_excel：
```python
In [106]: frame = pd.read_excel(&#39;examples/ex1.xlsx&#39;, &#39;Sheet1&#39;)

In [107]: frame
Out[107]: 
   a   b   c   d message
0  1   2   3   4   hello
1  5   6   7   8   world
2  9  10  11  12     foo
```

如果要将pandas数据写入为Excel格式，你必须首先创建一个ExcelWriter，然后使用pandas对象的to_excel方法将数据写入到其中：
```python
In [108]: writer = pd.ExcelWriter(&#39;examples/ex2.xlsx&#39;)

In [109]: frame.to_excel(writer, &#39;Sheet1&#39;)

In [110]: writer.save()
```

你还可以不使用ExcelWriter，而是传递文件的路径到to_excel：
```python
In [111]: frame.to_excel(&#39;examples/ex2.xlsx&#39;)
```

## 6.3 Web APIs交互
许多网站都有一些通过JSON或其他格式提供数据的公共API。通过Python访问这些API的办法有不少。一个简单易用的办法（推荐）是requests包（http://docs.python-requests.org）。

为了搜索最新的30个GitHub上的pandas主题，我们可以发一个HTTP GET请求，使用requests扩展库：
```python
In [113]: import requests

In [114]: url = &#39;https://api.github.com/repos/pandas-dev/pandas/issues&#39;

In [115]: resp = requests.get(url)

In [116]: resp
Out[116]: &lt;Response [200]&gt;
```

响应对象的json方法会返回一个包含被解析过的JSON字典，加载到一个Python对象中：
```python
In [117]: data = resp.json()

In [118]: data[0][&#39;title&#39;]
Out[118]: &#39;Period does not round down for frequencies less that 1 hour&#39;
```

data中的每个元素都是一个包含所有GitHub主题页数据（不包含评论）的字典。我们可以直接传递数据到DataFrame，并提取感兴趣的字段：
```python
In [119]: issues = pd.DataFrame(data, columns=[&#39;number&#39;, &#39;title&#39;,
   .....:                                      &#39;labels&#39;, &#39;state&#39;])

In [120]: issues
Out[120]:
    number                                              title  \
0    17666  Period does not round down for frequencies les...   
1    17665           DOC: improve docstring of function where   
2    17664               COMPAT: skip 32-bit test on int repr   
3    17662                          implement Delegator class
4    17654  BUG: Fix series rename called with str alterin...   
..     ...                                                ...   
25   17603  BUG: Correctly localize naive datetime strings...   
26   17599                     core.dtypes.generic --&gt; cython   
27   17596   Merge cdate_range functionality into bdate_range   
28   17587  Time Grouper bug fix when applied for list gro...   
29   17583  BUG: fix tz-aware DatetimeIndex &#43; TimedeltaInd...   
                                               labels state  
0                                                  []  open  
1   [{&#39;id&#39;: 134699, &#39;url&#39;: &#39;https://api.github.com...  open  
2   [{&#39;id&#39;: 563047854, &#39;url&#39;: &#39;https://api.github....  open  
3                                                  []  open  
4   [{&#39;id&#39;: 76811, &#39;url&#39;: &#39;https://api.github.com/...  open  
..                                                ...   ...  
25  [{&#39;id&#39;: 76811, &#39;url&#39;: &#39;https://api.github.com/...  open  
26  [{&#39;id&#39;: 49094459, &#39;url&#39;: &#39;https://api.github.c...  open  
27  [{&#39;id&#39;: 35818298, &#39;url&#39;: &#39;https://api.github.c...  open  
28  [{&#39;id&#39;: 233160, &#39;url&#39;: &#39;https://api.github.com...  open  
29  [{&#39;id&#39;: 76811, &#39;url&#39;: &#39;https://api.github.com/...  open  
[30 rows x 4 columns]
```

花费一些精力，你就可以创建一些更高级的常见的Web API的接口，返回DataFrame对象，方便进行分析。

## 6.4 数据库交互

在商业场景下，大多数数据可能不是存储在文本或Excel文件中。基于SQL的关系型数据库（如SQL Server、PostgreSQL和MySQL等）使用非常广泛，其它一些数据库也很流行。**数据库的选择通常取决于性能、数据完整性以及应用程序的伸缩性需求。**

将数据从SQL加载到DataFrame的过程很简单，此外pandas还有一些能够简化该过程的函数。例如，我将使用SQLite数据库（通过Python内置的sqlite3驱动器）：
```python
In [121]: import sqlite3

In [122]: query = &#34;&#34;&#34;
   .....: CREATE TABLE test
   .....: (a VARCHAR(20), b VARCHAR(20),
   .....:  c REAL,        d INTEGER
   .....: );&#34;&#34;&#34;

In [123]: con = sqlite3.connect(&#39;mydata.sqlite&#39;)

In [124]: con.execute(query)
Out[124]: &lt;sqlite3.Cursor at 0x7f6b12a50f10&gt;

In [125]: con.commit()
```

然后插入几行数据：
```python
In [126]: data = [(&#39;Atlanta&#39;, &#39;Georgia&#39;, 1.25, 6),
   .....:         (&#39;Tallahassee&#39;, &#39;Florida&#39;, 2.6, 3),
   .....:         (&#39;Sacramento&#39;, &#39;California&#39;, 1.7, 5)]

In [127]: stmt = &#34;INSERT INTO test VALUES(?, ?, ?, ?)&#34;

In [128]: con.executemany(stmt, data)
Out[128]: &lt;sqlite3.Cursor at 0x7f6b15c66ce0&gt;
```

从表中选取数据时，大部分Python SQL驱动器（PyODBC、psycopg2、MySQLdb、pymssql等）都会返回一个元组列表：
```python
In [130]: cursor = con.execute(&#39;select * from test&#39;)

In [131]: rows = cursor.fetchall()

In [132]: rows
Out[132]: 
[(&#39;Atlanta&#39;, &#39;Georgia&#39;, 1.25, 6),
 (&#39;Tallahassee&#39;, &#39;Florida&#39;, 2.6, 3),
 (&#39;Sacramento&#39;, &#39;California&#39;, 1.7, 5)]
```

你可以将这个元组列表传给DataFrame构造器，但还需要列名（位于光标的description属性中）：
```python
In [133]: cursor.description
Out[133]: 
((&#39;a&#39;, None, None, None, None, None, None),
 (&#39;b&#39;, None, None, None, None, None, None),
 (&#39;c&#39;, None, None, None, None, None, None),
 (&#39;d&#39;, None, None, None, None, None, None))

In [134]: pd.DataFrame(rows, columns=[x[0] for x in cursor.description])
Out[134]: 
             a           b     c  d
0      Atlanta     Georgia  1.25  6
1  Tallahassee     Florida  2.60  3
2   Sacramento  California  1.70  5
```

这种数据规整操作相当多，你肯定不想每查一次数据库就重写一次。[SQLAlchemy项目](http://www.sqlalchemy.org/)是一个流行的Python SQL工具，它抽象出了SQL数据库中的许多常见差异。pandas有一个read_sql函数，可以让你轻松的从SQLAlchemy连接读取数据。这里，我们用SQLAlchemy连接SQLite数据库，并从之前创建的表读取数据：
```python
In [135]: import sqlalchemy as sqla

In [136]: db = sqla.create_engine(&#39;sqlite:///mydata.sqlite&#39;)

In [137]: pd.read_sql(&#39;select * from test&#39;, db)
Out[137]: 
             a           b     c  d
0      Atlanta     Georgia  1.25  6
1  Tallahassee     Florida  2.60  3
2   Sacramento  California  1.70  5
```

## 6.5 总结

`访问数据`通常是数据分析的第一步。在本章中，我们已经学了一些有用的工具。在接下来的章节中，我们将深入研究数据规整、数据可视化、时间序列分析和其它主题。 

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/python/python%E6%95%B0%E6%8D%AE%E5%88%86%E6%9E%90/ch06-%E6%95%B0%E6%8D%AE%E5%8A%A0%E8%BD%BD%E5%AD%98%E5%82%A8%E4%B8%8E%E6%96%87%E4%BB%B6%E6%A0%BC%E5%BC%8F/  

