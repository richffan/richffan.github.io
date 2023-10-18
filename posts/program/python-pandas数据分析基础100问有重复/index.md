# Python pandas数据分析基础100问(有重复)

题目来自B站上看的一个讲爬虫的老师发的视频，不过代码没完全照那人的来，由于在平板上写的代码，就没有执行结果，感兴趣自行到B站搜原视频。  

```
import pandas as pd
```

1、将list转为dataframe

```
ls = [("join", 25, "male"), ("lisa", 29, "female"), ("david", 27, "male")``df = pd.DataFrame(ls, columns=["name", "age", "gender"])``display(df)
```

2、从csv文件中读取数据

```
df = pd.read_csv("filenpath", encoding="utf8")``# filepath_or_buffer: filepath 文件路径``# usecols: 需要读取的行列表``# nrows: 读取的行数``# na_values: 空值``# encoding: 编码格式
```

3、将dataframe导入到mysql

```
from sqlalchemy import create_engine``from sqlclchemy.type import *``df = pd.DataFrame() # 数据集``# 创建引擎``conn = create_engine("mysql+mysqlconnector://username:password@127.0.0.1:3306/db_name")``# 写入到数据库``df.to_sql("table_name", engine)
```

4、查看一个dataframe的行数、列数

```
# 行``df.shape[0]``# 列``df.shape[1]
```

5、查看dataframe的列名

```
df.columns # serise格式``df.columns.tolist() # list格式
```

6、查看dataframe索引

```
df.index # series格式``df.index.tolist()
```

7、查看pandas库的版本号

```
pd.__version__
```

8、从csv文件里读取数据，并创建一个dataframe

```
df = pd.read_csv("filepath", encoding='utf8')
```

9、查看dataframe的列的数据类型

```
df.dtypes
```

10、查看dataframe摘要统计描述信息

```
df.describe() # count/mean/std/min/max 1-4分位数，1-10分位数``# include: 统计哪些列，默认为None，全部``# exclude: 排除列``# percentiles: 分位点,默认为None,4分位
```

11、dataframe提取行

```
df.loc[index_name] # 单行``df.loc[index_name1, index_name2] # 多行``df.iloc[start_index: end_index] # 连续多行，不包含end
```

12、dataframe提取列

```
df["col_name"] # 单列``df[["col_1", "col_2", ....]] # 多列``df.iloc[:, col_index] # 单列``df.iloc[:, col_start_index: col_end_index] # 按列索引取 连续多列
```

13、选择datafram的行和列

```
df.loc[[indexName1, indexName2..], [colName1, colName2...]]  # 按行列名提取``   ``df.iloc[[startIndex: endIndex], [colStartIndex: colEndIndex]] # 按行列索引提取，不包括end
```

14、筛选dataframe的行

```
# 使用列条件筛选行``df[df[列名条件]] # 单条件``df[(条件1) & (条件2)] # 多条件，每一个单条件用()包起来，连接符号：&-与，|-或，～非``# 示例：筛选出年龄小于等于30，且收入大于等于30万的男性``df[(df["age"]<=30) & (df["income"]>=300000) & df["gender"]=="male"]
```

15、筛选dataframe的行和列

```
df[(条件1) & (条件2)][[colName1, colName2....]] # 筛选行之后再选择需要的列``df.loc[行筛选条件, [colName1, colName2]] # 使用列进行行条件筛选，同时限定输出列
```

16、根据某一列的值对一个dataframe进行排序

```
df.sort_values(by=["col1", "col2"], ascending=True) # 默认升序
```

17、对dataframe进行数据透视操作，即实现excel中的数据透视表功能

```
df.pivot_table(index='行索引', columns='列索引', values="统计列", aggfunc='函数名')``# aggfunc 支持 sum/mean/count/min/max等
```

18、对dataframe进行聚合操作

```
df.groupby(by=["索引列1", "索引列2"])['统计列'].agg(["函数1", "函数2", "函数3"])
```

19、对多个dataframe实现合并操作

```
# 纵向合并` `pd.merge(df_left, df_right, on=["合并列1", "合并列2"], how="inner or left or right or outer")``# 横向合并``pd.concat([df_1, df_2, df_3]) # 常用于拥有相同格式的多个数据集
```

20、删除dataframe中的一列或多列数据

```
df.drop(columns=["需要删除的列"], inplace=True) # 默认是生成一个新的数据集，inplace=False``   ``df.dorp(subset=["需要删除的列"], axis=1, inplace=True)
```

21、向dataframe中添加一行数据

```
df.iloc[df.shape[0]] = index_array
```

22、删除dataframe中的一行或多行数据

```
df.drop(subset=[1, 3], axis=0, inplace=True)
```

23、使用索引选择dataframe中某个范围内的行

```
df[1: 4] # 提取第一行到第四行
```

24、使用某一列选择dataframe中某个范围内的行

```
df.set_index("name").loc["start_indexName": "end_indexName"] # 连续索引名``# 不连续行，在列表中用逗号分隔即可
```

25、在dataframe中按特定条件取行

```
df[df["name"] == "小杨"] # 单条件、单取值``df[df["name"] in ["小杨", "小李"]] # 单条件、多取值``df[(df["name"] in ["小杨", "小李"]) & (df["age"] >= 30)] # 多条件
```

26、修改dataframe的列名

```
# 修改``df.rename(columns={"原列名": "修改后的列名"}, inplace=True)``# 重新设置``df.columns = [新的列名数组]
```

27、计算dataframe中某一列的总和

```
df["colName"].sum()
```

28、计算dataframe中某一列的平均值

```
df["col_name"].mean() # 返回浮点类型
```

29、计算dataframe中某一列的中位数

```
df["col_name"].median() # 返回浮点类型
```

30、计算dataframe中某一列的标准差

```
df["col_name"].std()
```

31、计算dataframe中某一列的方差

```
df["col_name"].var()
```

32、计算dataframe中某一列的最大值和最小值

```
df["col_name"].max()``df["col_name"].min()
```

33、查找dataframe某个单元格

```
df.loc[row, col]
```

34、dataframe重命名列名

```
df.rename(columns={"key": "value"}, inplace=True)
```

35、替换dataframe中的特定值

```
# 单列-1``df["col_name"].replace(to_replace="replace_str", value="str_new", regex=True)``# 单列-2``df["col_name"].map(lambda x: str(x).replace("replace_str", "str_new"))
```

36、如何将dataframe中的特定制替换成缺失值

```
df = df.replace(replace_value, np.nan) # 空值：np.nan pd.NaT
```

37、将dataframe中的空值进行填充

```
df.fillna(value, inplace=True)``# value = -999, 直接填充值``# value = "ffill", 使用上一个非空值进行填充``# value = "bfill", 使用后一个非空值进行填充``# value = {"col1": value1, "col2": value2}, 传入字典，不同列填充不同的值
```

38、删除dataframe中的缺失值

```
df.dropna(subset=["col1", "col2"], how='any', axis=0, inplace=True)``# subset: 用来判断是否删除的列``# axis: 0-删除行, 1-删除列``# how: any-一旦存在空值就删除, all-全为空时删除``# thresh: 0-1的浮点数，空值占比大于一定比例时删除
```

39、在dataframe中使用聚合函数

```
df.groupby(by=["col1", "col2"]).mean()``# 也可以使用count/sum等函数
```

40、在pandas中创建一个空的数据帧

```
df = pd.DataFrame()
```

41、在pandas中进行分组、聚合

```
df.groupby(by=["col1_groupby", "col2_groupby"])["col1_calculate", "col2_calculate"].agg("func1", "func2"[, ....])
```

42、对dataframe中的数据进行类型转换

```
df["col_name"] = df["col_name"].astype(np.int)
```

43、如何使用pandas中的迭代循环方法

```
for index,row in df.iterrows():`    `print(index, row)
```

44、如何使用pandas中的交叉表(列联表)

```
pd.crosstab(df["col_index"], df["col_feature"]) # 默认是统计次数
```

45、如何在pandas中使用重塑(reshape)数据？

```
df.pivot(index="col_index", columns="col_feature", values="col_calculate")
```

46、如何在pandas中使用变因子(factorize)函数

```
df["new"] = pd.factorize(df["class"]) # 相当于是把字符串类别变量进行数值化
```

47、在dataframe中创建随机日期列

```
start_date = '2023-01-01'``end_date = '2023-07-27'``data_range = pd.date_range(start=start_date, end=end_date)``   ``# size=df.shape[0] 即可创建一个跟dataframe等长的列``df["random_date"] = pd.Series(np.random.choice(date_range, size=df.shape[0]))
```

48、如何在pandas中使用字符串函数

```
# to_uppercase 字符串中的字母转为大写``# to_lowercase 字符串中的字母转为小写``# strip 去除字符串两侧的空格``# replace 将字符串中的一个子字符串替换成另外一个子字符串``# contains 判断字符串是否包含给定的子字符串``# startwith 判断字符串是否以给定字符串开头``# endwith 判断字符串是否以给定字符串结尾``# split 切分字符串``   ``df['col_name'] = df['col_name'].str.strip() # 对某一列去除空格``df['col_name'] = df['col_name'].str.strip().str.to_lowercase() # 对某一列去除空格，然后字母转小写
```

49、如何在dataframe中使用字符串拆分函数

```
df[['col_1', 'col2']] = df['col_to_split'].str.split("用来切分的字符串", expand=True)
```

50、如何对dataframe进行标准化处理

```
import pandas as pd``from sklearn.preprocessing import StandardScaler``   ``df_new = StandardScaler().fit_transform(df)
```

51、如何对dataframe进行归一化处理

```
import sklearn.preprocessing import MinMaxScaler``df_new = MinMaxScaler().fit_transform(df)
```

52、在pandas中使用数据透视表进行数据汇总

```
pd.pivot_table(df, values="计算列", index='索引列', columns="特征列", agg="func_name")``# 跟pivot也没有太大差别
```

53、在dataframe中使用groupby进行数据汇总

```
df.groupby(by=["col_name1", "col_name2"]).sum() # .mean .max .std
```

54、如何在Series中使用rolling函数进行移动窗口统计？

```
import pandas as pd``import numpy as np``se = pd.Series(np.random.randn(10))``>>> se``0   -0.974864``1    1.916511``2    0.760756``3    0.110363``4   -0.137077``5    0.186016``6   -0.759734``7   -1.529321``8   -1.571819``9    0.123134``   ``se_rolling = se.rolling(window=3).mean()``>>> se_rolling` `0         NaN``1         NaN``2    0.567468``3    0.929210``4    0.244681``5    0.053101``6   -0.236932``7   -0.701013``8   -1.286958``9   -0.992669``   ``计算逻辑：由于index为0、1时，由于不足移动数量，结果为NaN``index2 =（index0+index1+index2 ）/ 3``index3 =（ index1+index2+index3）/ 3``   ``# 语法``DataFrame.rolling(`    `window, # 表示时间窗的大小，数值表示计算统计量的观测值的数量即向前几个数据`    `min_periods=None, # 每个窗口最少包含的观测值数量，小于这个值的窗口结果为NA，值可以是int，默认None。`    `center=False, # 把窗口的标签设置为居中。布尔型，默认False，居右`    `win_type=None, # 窗口的类型。截取窗的各种函数。字符串类型，默认为None。`    `on=None, # 可选参数。对于dataframe而言，指定要计算滚动窗口的列。值为列名。`    `axis=0, # 默认为0，即对列进行计算`    `closed=None # 定义区间的开闭，默认是左开右闭的即默认为right,默认是左开右闭的即默认为right`    `)
```

55、在pandas中使用shift函数进行时间序列处理

```
import pandas as pd``import numpy as np``import datetime``   ``# 创建一个时间数组``date_range = pd.date_range(start='2023-01-01', end='2023-07-28', periods=10)``# 创建一个Serise``se = pd.Series(np.random.randn(10), index=date_range)``>>> se``2023-01-01 00:00:00    0.294050``2023-01-24 02:40:00   -0.324300``2023-02-16 05:20:00    0.345552``2023-03-11 08:00:00   -1.519163``2023-04-03 10:40:00    0.446149``2023-04-26 13:20:00    1.109210``2023-05-19 16:00:00    1.061693``2023-06-11 18:40:00   -1.417034``2023-07-04 21:20:00   -0.218537``2023-07-28 00:00:00   -0.365196``   ``# shift数据前滞 或 后滞``result = se.shift(periods=3, freq=datetime.timedelta(-3))``>>> result``2022-12-23 00:00:00    0.294050``2023-01-15 02:40:00   -0.324300``2023-02-07 05:20:00    0.345552``2023-03-02 08:00:00   -1.519163``2023-03-25 10:40:00    0.446149``2023-04-17 13:20:00    1.109210``2023-05-10 16:00:00    1.061693``2023-06-02 18:40:00   -1.417034``2023-06-25 21:20:00   -0.218537``2023-07-19 00:00:00   -0.365196``   ``# DataFrame.shift(periods=1, freq=None, axis=0)``# periods可以理解为移动幅度的次数，负数表示前滞，正数表示后滞。``# freq为None时，移动的是其他数据的值，即移动periods*1个单位长度。可以设为一个timedelta对象。适用于索引为时间序列数据时。``# axis默认为0，表示对列操作。如果为行则表示对行操作。
```

56、如何在pandas中使用set\_index函数设置索引

```
df.set_index(keys=['a', 'b']) # 将a\b两列设置为索引``   ``# set_index(keys, drop=True, inplace=False)``# drop: 是否删除原本的列，默认True-删除``# inplace: 是否原地修改，默认False-否
```

57、如何在pandas中使用reset\_index函数重置索引

```
df.reset_index() # 什么都不传入，默认把所有层级的索引重置为新的字段``   ``# reset_index(level=None, drop=False, inplace=False)``# level: 需要重置的索引层级``# drop: 是否删除，默认False``# inplace: 是否原地操作，默认False
```

58、如何在pandas中使用map进行数据操作

```
# map函数用于列操作``# 数据映射``mapping_dict = {"male":0, "female":1}``df["gender"] = df["gender"].map(mapping_dict)``   ``# 更常见的操作是向map传入一个函数``def mapping(x):`    `mapping_dict = {"male":0, "female":1}`    `return mapping_dict[x]``df["gender"] = df["gender"].map(mapping)``   ``# 传入匿名函数``df["gender"] = df["gender"].map(lambda x: lambda 1 if x=="female" else 0)
```

59、如何在pandas中使用apply函数进行元素级函数应用

```
# apply函数用于单元格级的操作``def multiply_by_2(x)`    `return x*2``df_new = df.apply(multiply_by_2) # df中的每一个元素都乘以2``   ``# 个人比较常用的操作：``df_new = df.groupby(groupby).apply(lambda df: pd.Series({`    `'cnt': df.shape[0],`    `'bad_cnt': df[label].sum(),`    `# 当标签类型只有0、1的时候，可以使用均值计算坏账率`    `'bad_rate': df[label].mean(),`    `'begin_date': df[apply_time_label].min(),`    `'end_date': df[apply_time_label].max()})).reset_index()`    `# 造特征``df["col_new"] = df.apply(lambda x: x["col1"] if x["col"]>=100 else x["col2"], axis=1)
```

60、如何在pandas中使用agg函数进行分组聚合操作

```
df.groupby(['col_1', "col_2"]).agg({"col3":"sum", "col4":"mean"})
```

61、如何在pandas中使用transform函数进行分组变换操作

```
# 通过分组后的计算，取到对分组结果进行赋值的效果``df["col_new"] = df.groupby('col1')['col_calculate'].transform(lambda x: x.mean())
```

62、如何在pandas中使用stack和unstack函数进行数据透视操作

```
# stack: 列转行``import pandas as pd``import numpy as np``df = pd.DataFrame({"name": ['alice', 'bob', 'charlie'],`        `"year": [2020, 2021, 2022],`        `"age": [20, 30, 40]})``df.set_index("year", inplace=True)``>>> df`         `name  age``year`              `2020    alice   20``2021      bob   30``2022  charlie   40``   ``df.stack() # 列转行-创建新dataframe``>>>` `year`      `2020  name      alice`      `age          20``2021  name        bob`      `age          30``2022  name    charlie`      `age          40``   ``# unstack: 行转列``df.stack().unstack() # 行转列-还原为原来的效果
```

63、如何在pandas中使用crosstab函数进行数据透视操作

```
pd.crosstab(index, # 行`            `columns, # 列`            `values=None, # 聚合列`            `aggfunc=None, # 聚合函数`            `rownames=None, # 行-分组名称`            `colnames=None, # 列-分组名称`            `dropna=True, # 是否删除空值`            `normalize=False) # 是否标准化
```

64、在pandas中使用cut函数对数据进行离散化(分箱)

```
pd.cut(x=df["age"], bins=10, right=True, labels=None)``# bins: 为整数表示分箱数量(等宽分)，为列表时表示切分点``# right: True-左开右闭``# labels: 分箱后显示的分箱名``# retbins=False, 是否返回分箱边界值
```

65、如何在pandas中使用qcut函数对数据进行等频分箱

```
qcut(x, # 待分箱的数据`     `q, # 等频分箱数量`     `labels=None, # 分箱后箱体名`     `retbins=False, # 是否返回分箱边界值`     `precision=3, # 分位数精度`     `duplicates='raise' # 切分点重复如何处理`     `)
```

66、在pandas中使用interpolate函数进行数据插值操作

```
df.interpolate(method='liner', axis=0, limit=None, inplace=False)``# liner-线性插值
```

67、在pandas中使用bfill 和 ffill 函数进行数据填充操作

```
df.bfill() # 使用后值填充``df.ffill() # 使用前值填充
```

68、在pandas中使用正则表达式进行数据过滤操作

[【python 基础】re 正则表达式](http://mp.weixin.qq.com/s?__biz=MzU1NzAyODM4NQ==&mid=2247485198&idx=1&sn=69c501b9727f18d654fc3fecf06f2f64&chksm=fc3d4240cb4acb563bd285ce223d997bd5797713abbac6f14ce3d3804172d02837bc3b86a428&scene=21#wechat_redirect)  

```
# pd.Series.str.contains(pat, case=True, flags=0, na=None, regex=True)``# pat: 字符类型 字符序列或者正则表达式``# case: 布尔类型 是否区分大小写``# flags: int类型 默认0 标志传递到re模块 比如flags=re.INGNORECASE 等价于 case=False``# na: 默认NaN 填写缺失值的值 可以随意指定``# regex: 布尔类型 默认True 假定pat为正则表达式 针对特殊字符需要设置regex=False``df = pd.DataFrame({"name": ["alice", "bob", "charlie", "diana", "emily"],`                   `"age": [10, 20, 30, 40, 50],`                   `"email": ["alice@qq.com", "bob@163.com", "charlie@yahoo.com", "diana@gmail.com", "emily@hotmail.com"]})``>>> df`      `name  age              email``0    alice   10       alice@qq.com``1      bob   20        bob@163.com``2  charlie   30  charlie@yahoo.com``3    diana   40    diana@gmail.com``4    emily   50  emily@hotmail.com``   ``df[df["email"].str.contains(pat="gmail")]``>>> ``    name  age            email``3  diana   40  diana@gmail.com``   ``df[df["email"].str.contains(pat="gmail|qq", case=False)]``>>> ``    name  age            email``0  alice   10     alice@qq.com``3  diana   40  diana@gmail.com
```

69、如何在pandas中使用diff函数进行时间序列处理

```
# 使用场景：现有某个用户的操作日志，想知道上一次操作跟下一次操作的间隔时间是多少。``date_range = pd.date_range('2023-01-01', periods=5)``se = pd.Series(data=date_range)``>>> se``0   2023-01-01``1   2023-01-02``2   2023-01-03``3   2023-01-04``4   2023-01-05``   ``se_diff = se.diff() # 不填数字表示一阶，填了数字表示多阶，递减``>>> se_diff``0      NaT``1   1 days``2   1 days``3   1 days``4   1 days
```

70、在dataframe中使用dropna进行删除空值的操作

```
pd.DataFrame.dropna(axis=0, how='any', thresh=None, subset=None, inplace=True)``# subset: 用来判断是否删除的列``# axis: 0-删除行, 1-删除列``# how: any-一旦存在空值就删除, all-全为空时删除``# thresh: 0-1的浮点数，空值占比大于一定比例时删除``# inplace: 是否原地操作
```

71、如何在panda中使用join函数进行数据拼接操作

```
# 类似于sql中的join，但不会自动合并相同列名的字段``# 只要两个表列名不同，不加任何参数就可以直接用。``# 如果两个表有重复的列名，需指定 lsuffix（左侧数据重列列使用的列名后缀）, rsuffix 参数。``# 参数的意义与 merge 方法基本相同。``   ``df1.join(df2, on=["col1", "col2"], how='inner')
```

![](https://mmbiz.qpic.cn/mmbiz_png/phUvdvw2lot9RwMjBXxQxEic8dFea6rnKToZ0oSjlbGH3DZDDav6YaYn2fHDG4iaKMmpIqhyicqic6FWkY4NYT5dzQ/640?wx_fmt=png)

72、如何在pandas中使用get\_dummies函数进行哑变量编码操作

```
df = pd.DataFrame({"color": ["red", "blue", "green", "red", "grey"]})``>>> df`   `color``0    red``1   blue``2  green``3    red``4   grey``   ``pd.get_dummies(df['color'])``>>> ``   blue  green  grey  red``0     0      0     0    1``1     1      0     0    0``2     0      1     0    0``3     0      0     0    1``4     0      0     1    0
```

73、如何将dataframe中的某一列转为int类型

```
df['col_name'] = df['col_name'].astype(np.int)
```

74、如何将dataframe中的某一列转为float类型

```
# 方法1``df['col_name'] = df['col_name'].astype(np.float)``# 方法2``df['col_name'] = df['col_name'].map(lambda x: float(x))``   ``不管是转为int还是float，都需要把数据中的特殊字符(如：NA, NULL等)转为空值，再转为需要的类型
```

75、如何使用to\_excel进行excel数据写入操作

```
df.to_excel(excel_writer="文件路径",`    `sheet_name="Sheet1", # sheet表名`    `)
```

76、如何使用to\_json将数据保存为json文件

```
df.to_json(path_or_buf="df.json", orient=None)``# orient:``#     split: 将json文件转换为多个文件，每个文件包含dafaframe中的一行``#     records: 将json文件转换为一个列表，每个元素是dataframe中的一行``#     index: 将json文件转换为一个字典，键是列名，值是一行记录``#     columns: 将json文件转换为一个字典，键是列名，值是一列记录``#     values: 将json文件转换为一个二维数组，每行是dataframe的一行记录
```

77、如何在pandas中使用plot函数画出柱形图

```
import pandas as pd``import numpy as np``import matplotlib.pyplot as plt``   ``df = pd.DataFrame({"name": ["tom", "jerry", "mickey", "minnie", "donald"],`        `"age": [25, 23, 28, 27, 26],`        `"score": [80, 90, 75, 85, 95]})``df.plot(kind="bar", x="name", y='score')``# plt.show() # jupyter需要执行plt.show才能在界面上显示图片
```

78、如何在pandas中使用plot函数画出折线图

```
# 当设置的标签信息有中文时，一般要先做一些设置``plt.rcParams['font.sans-serif']=['SimHei'] # 显示中文``plt.rcParams['axes.unicode_minus']=False # 取消使用unicode的负号``   ``plt.plot(df["col_x"], df["col_y"])``plt.title("图标题")``plt.xlabel("x轴名")``ply.ylabel("y轴名")``plt.show()
```

79、如何在pandas中使用plot函数画散点图

```
df.plot(kind="scatter", x="name", y='score')
```

80、如何在pandas中使用plot函数画出箱线图

```
# 方法1``df.plot(kind="box", x="name", y='score')``# 方法2``df.plot.box() # 默认对所有数值列绘制箱线图
```

![](https://mmbiz.qpic.cn/mmbiz_png/phUvdvw2lot9RwMjBXxQxEic8dFea6rnK6eO7MOA4IHpPQ6vjiatXzXkXdR04qK1KwvSoxYvBWeMroH6055l8rWQ/640?wx_fmt=png)

81、如何在pandas中使用plot函数画出面积图

```
pd.DataFrame.plot.area(x=None, y=None)``# 不传入x/y则是一个多列堆叠的面积图
```

82、如何在pandas中使用plot函数绘制饼图

```
df.plot(kind='pie', y="绘制饼图的列")
```

83、如何在pandas中对一列数据的正数和负数进行分组聚合操作

```
import pandas as pd``df = pd.DataFrame({"val": [1, -2, 3, -4, 5]})``df["positive"] = df["val"] > 0``df["negative"] = df["val"] < 0``df_result = df.groupby(by=["positive", "negative"]).agg({"val": "sum"})``df_result` `# 不知道意义在哪....
```

84、如何在pandas中实现两列的拼接

```
df["new"] = df["col_1"] + df["col_2"]
```

85、如何对数据框中的字符串进行模糊匹配

```
# pd.Series.str.contains(pat, case=True, flags=0, na=None, regex=True)``# pat: 字符类型 字符序列或者正则表达式``# case: 布尔类型 是否区分大小写``# flags: int类型 默认0 标志传递到re模块 比如flags=re.INGNORECASE 等价于 case=False``# na: 默认NaN 填写缺失值的值 可以随意指定``# regex: 布尔类型 默认True 假定pat为正则表达式 针对特殊字符需要设置regex=False
```

86、如何使用pandas处理文本数据

```
# lower() / upper() 字母转小写、大写``# strip() 去除首尾空格``# split() 字符串切分``# contains() 判断字符串是否包含指定文本``# replace() 文本替换``# extract() 从字符串中提取指定的文本``df['col_new'] = df["col"].str.func()
```

87、如何在dataframe中找到重复的行

```
df.duplicated(subset=["col_1", "col2"])``# subset: 判断是否重复的字段``# axis=0,默认判断重复行
```

88、如何查看数的异常值

```
# 通过数据描述信息，判断是否存在异常值``import pandas as pd``df.describe()``   ``# 通过箱线图，判断是否存在异常值``from matplotlib import pyplot as plt``plt.boxplot(df)` `# df可以是一个Series，也可以是一个DataFrame``# 当df是一个DataFrame时，建议选取处于同一数量级，不然画出来的图就很难判断
```

89、如何查找dataframe中的缺失值

```
# 使用某一列判断，得到存在空值的行``df[df['col'].isnull()]``   ``# 判断整个dataframe每个字段的数量``df.isnull().sum()``   ``# info可以获取数据集信息``df.info()
```

90、如何在dataframe中获取某一列的值只出现一次的所有行

```
se_value_counts = df['col'].value_counts()``df[df["col"].isin(se_value_counts[se_value_counts==1].index)]
```

91、如何在pandas中求相邻元素的差异

```
df['col'].diff()` `# 当col字段是时间时，求的是前后两条记录的时间差``# 当col字段是数值时，求的是间隔元素的差值
```

92、如何将dataframe中的某一列中包含的字符串进行拆分

```
# 新字段是一个列表，并且没找到切分符号会返回空值``df['col_new'] = df['col'].str.split("_")``   ``# 取切分后的第一个值``df["col_new"] = df["col"].map(lambda x: str(x).split("_")[0])
```

93、如何使用pandas 对某一列数据进行md5加密

```
# 联合建模时，对姓名、手机号、身份证等敏感信息进行加密``import hashlib``df['col_md5'] = df["col"].map(lambda x: hashlib.md5(x).hexdigest())
```

94、如何在pandas中使用groupby和agg进行分组聚合

```
df.groupby(by=["col_1", "col_2"]).agg({"score": "sum", "age": "mean"})
```

95、如何在dataframe中使用cumsum函数进行累计计算

```
# 整个dataframe横向或纵向求和, 0-横向,1-纵向``df.cumsum(axis=1)
```

96、如何将dataframe根据某一列的值进行过滤

```
# 筛选出年龄 大于 30 的 男性``df[(df['age']>30) & df['gender']=='male']
```

97、如何使用pandas中的值计算出新的一列

```
df['new'] = 0.4 * df["col"] + 0.6*100
```

98、如何为dataframe的每个元素使用一个自定义函数

```
def test(x)`    `return x+1``df_new = df.apply(test)
```

99、如何使用pandas对数据进行预处理

```
# 读取数据``df = pd.read_csv(datapath) # pd.read_excel(datapath, sheet_name)``# 描述信息``df.describe()``# 数据详情``df.info()``# 类型转换``df["col"] = df["col"].astype(np.int)``# 删除空值``df.dropna()``# 删除重复值``df.drop_duplicates()``# 填充空值``df.fillna()
```

100、如何从一个csv文件中读取数据，将其中的数值数据进行聚合，并计算每个聚合后的组的中位数

```
# 读取数据``df = pd.read_csv(datapath, encoding='utf-8')``# 删除重复值``df.drop_duplicates(inplace=True)``# 删除空值``df.dropna(subset=[...], inplace=True, how='any')``# 分组聚合``df.groupby(by=["col_1", "col_2"]).agg({"col_cal": "median"})
```

  

本文转自 [https://mp.weixin.qq.com/s/ER3s5yjMwMzQqiEr8YwYJQ](https://mp.weixin.qq.com/s/ER3s5yjMwMzQqiEr8YwYJQ)，如有侵权，请联系删除。

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/posts/program/python-pandas%E6%95%B0%E6%8D%AE%E5%88%86%E6%9E%90%E5%9F%BA%E7%A1%80100%E9%97%AE%E6%9C%89%E9%87%8D%E5%A4%8D/  

