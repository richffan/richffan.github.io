# Pandas最常用的15个函数

Pandas是一个功能强大的数据处理库，它提供了大量的函数和方法，用于数据的读取、清洗、转换、分析和可视化。下面是Pandas中最常用的15个函数及其代码示例：

1. read_csv()：用于读取CSV文件中的数据并创建DataFrame对象。
```python
import pandas as pd

df = pd.read_csv('data.csv')
print(df)
```

2. head()：用于查看DataFrame中的前几行数据，默认显示前5行。

```python
print(df.head())
```

3. tail()：用于查看DataFrame中的后几行数据，默认显示后5行。

```python
print(df.tail())
```

4. info()：用于查看DataFrame的基本信息，包括列名、数据类型、非空值数量等。

```python
print(df.info())
```

5. describe()：用于生成DataFrame中数值列的统计摘要，包括计数、均值、标准差、最小值、25%分位数、中位数、75%分位数和最大值。

```python
print(df.describe())
```

6. shape：用于获取DataFrame的形状，即行数和列数。

```python
print(df.shape)
```

7. columns：用于获取DataFrame的列名。

```python
print(df.columns)
```

8. fillna()：用于填充DataFrame中的缺失值。

```python
df.fillna(0, inplace=True)
print(df)
```

9. groupby()：用于对DataFrame进行分组操作。

```python
grouped = df.groupby('category')
print(grouped.mean())
```

10. sort_values()：用于对DataFrame按照指定列进行排序。

```python
sorted_df = df.sort_values('column_name')
print(sorted_df)
```

11. drop_duplicates()：用于移除DataFrame中的重复行。

```python
df.drop_duplicates(inplace=True)
print(df)
```

12. value_counts()：用于计算DataFrame中某一列的唯一值计数。

```python
print(df['column_name'].value_counts())
```

13. apply()：用于对DataFrame的某一列或多列应用自定义函数。

```python
def custom_function(x):
    # 自定义处理逻辑
    return x * 2

df['column_name'] = df['column_name'].apply(custom_function)
print(df)
```

14. merge()：用于合并两个DataFrame对象。

```python
merged_df = pd.merge(df1, df2, on='key_column')
print(merged_df)
```

15. plot()：用于绘制DataFrame中的数据图表。

```python
df.plot(x='column_name1', y='column_name2', kind='line')
plt.show()
```

通过掌握这些常用函数，您可以更加高效地使用Pandas进行数据处理和分析。它们提供了基本的数据操作和转换功能，为您的数据工作提供了便利。

总结：本文介绍了Pandas中最常用的15个函数，并提供了相应的代码示例。这些函数涵盖了数据读取、查看、描述统计、缺失值处理、分组、排序、去重、计数、自定义函数应用、合并和可视化等常见的数据处理任务。掌握这些函数将帮助您更加熟练地使用Pandas进行数据分析和处理。

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/posts/program/python/pandas/pandas%E6%9C%80%E5%B8%B8%E7%94%A8%E7%9A%8415%E4%B8%AA%E5%87%BD%E6%95%B0/  

