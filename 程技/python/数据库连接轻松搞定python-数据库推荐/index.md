# 数据库连接轻松搞定：Python 数据库推荐

## SQLite3：轻巧便捷的数据库连接

首先，让我们认识一下 SQLite3 这个轻巧便捷的库。它是 Python 中自带的数据库模块，适用于小型应用和快速原型开发。

让我们来看看 SQLite3 的魔法：

```python
import sqlite3

## 连接数据库
conn = sqlite3.connect(&#34;mydatabase.db&#34;)

## 创建表格
cursor = conn.cursor()
cursor.execute(&#34;CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, name TEXT, age INTEGER)&#34;)

## 插入数据
cursor.execute(&#34;INSERT INTO users (name, age) VALUES (?, ?)&#34;, (&#34;Alice&#34;, 25))
cursor.execute(&#34;INSERT INTO users (name, age) VALUES (?, ?)&#34;, (&#34;Bob&#34;, 30))

## 查询数据
cursor.execute(&#34;SELECT * FROM users&#34;)
rows = cursor.fetchall()
for row in rows:
    print(row)

## 关闭连接
conn.close()
```

运行上述代码，你将看到插入的数据和查询结果输出在终端上，SQLite3 让数据库连接变得如此简单！

## MySQL Connector：连接MySQL

接下来，我们要介绍 MySQL Connector 这个强大的库。它是用于连接 MySQL 数据库的官方驱动，适用于中小型应用和生产环境。



MySQL Connector 的安装非常简单，使用以下命令：
```python
pip install mysql-connector-python
```

让我们看看 MySQL Connector 的魔法：

```python

import mysql.connector

## 连接数据库
conn = mysql.connector.connect(
    host=&#34;localhost&#34;,
    user=&#34;root&#34;,
    password=&#34;your_password&#34;,
    database=&#34;mydatabase&#34;
)

## 创建表格
cursor = conn.cursor()
cursor.execute(&#34;CREATE TABLE IF NOT EXISTS users (id INT AUTO_INCREMENT PRIMARY KEY, name VARCHAR(255), age INT)&#34;)

## 插入数据
cursor.execute(&#34;INSERT INTO users (name, age) VALUES (%s, %s)&#34;, (&#34;Alice&#34;, 25))
cursor.execute(&#34;INSERT INTO users (name, age) VALUES (%s, %s)&#34;, (&#34;Bob&#34;, 30))

## 查询数据
cursor.execute(&#34;SELECT * FROM users&#34;)
rows = cursor.fetchall()
for row in rows:
    print(row)

## 关闭连接
conn.close()
```

运行上述代码，你将看到插入的数据和查询结果输出在终端上，MySQL Connector 让数据库连接变得如此高效！

## SQLAlchemy：强大的ORM框架

最后，让我们来认识 SQLAlchemy 这个强大的库。它是 Python 中最流行的 ORM（Object-Relational Mapping）框架，适用于复杂的数据库操作和应用。



SQLAlchemy 的安装非常简单，使用以下命令：

```python
pip install sqlalchemy
```

让我们看看 SQLAlchemy 的魔法：
```python

from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

## 连接数据库
engine = create_engine(&#34;sqlite:///mydatabase.db&#34;, echo=True)
Base = declarative_base()

## 定义数据表模型
class User(Base):
    __tablename__ = &#34;users&#34;
    id = Column(Integer, primary_key=True)
    name = Column(String)
    age = Column(Integer)

## 创建数据表
Base.metadata.create_all(engine)

## 创建会话
Session = sessionmaker(bind=engine)
session = Session()

## 插入数据
user1 = User(name=&#34;Alice&#34;, age=25)
user2 = User(name=&#34;Bob&#34;, age=30)
session.add_all([user1, user2])
session.commit()

## 查询数据
users = session.query(User).all()
for user in users:
    print(user.name, user.age)

## 关闭会话
session.close()
```
运行上述代码，你将看到插入的数据和查询结果输出在终端上，SQLAlchemy 让数据库操作变得如此高端！



## 数据库连接案例

下面我们来看一个实际应用的数据库连接案例：从 CSV 文件中读取数据并存储到数据库。

```python
import csv

## 连接数据库（使用 SQLite3 示例）
conn = sqlite3.connect(&#34;mydatabase.db&#34;)
cursor = conn.cursor()

## 创建表格
cursor.execute(&#34;CREATE TABLE IF NOT EXISTS books (id INTEGER PRIMARY KEY, title TEXT, author TEXT)&#34;)

## 从 CSV 文件中读取数据
with open(&#34;books.csv&#34;, &#34;r&#34;) as file:
    reader = csv.reader(file)
    next(reader)  ## 跳过标题行
    for row in reader:
        title, author = row
        cursor.execute(&#34;INSERT INTO books (title, author) VALUES (?, ?)&#34;, (title, author))

## 查询数据
cursor.execute(&#34;SELECT * FROM books&#34;)
rows = cursor.fetchall()
for row in rows:
    print(row)

## 关闭连接
conn.close()
```

运行上述代码，你将看到读取的数据和查询结果输出在终端上，通过这样简单的数据库连接，你可以更好地管理和查询数据。



---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/python/%E6%95%B0%E6%8D%AE%E5%BA%93%E8%BF%9E%E6%8E%A5%E8%BD%BB%E6%9D%BE%E6%90%9E%E5%AE%9Apython-%E6%95%B0%E6%8D%AE%E5%BA%93%E6%8E%A8%E8%8D%90/  

