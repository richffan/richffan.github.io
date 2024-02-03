# 3 栈


&gt; 栈是一种数据结构，只能从一端插入和删除操作，遵循着先进后出原则存储数据。

![栈](https://gitee.com/wugenqiang/images/raw/master/02/image-20201220170815652.png)

## 3.1 栈的初始化

```python
def __init__(self):
    self.stack = []  # 栈列表
    self.size = 20  # 栈大小
    self.top = -1  # 栈顶位置
```



## 3.2 元素进栈

```python
# 元素进栈
def push(self, element):
    self.stack.append(element)
    self.top &#43;= 1
```



## 3.3 元素出栈

```python
# 元素出栈
def pop(self):
    element = self.stack[-1]
    self.top -= 1
    del self.stack[-1]
    return element
```

这里可以直接调用pop函数，使用如下：

```python
self.stack.pop() # 弹出栈顶元素
```



## 3.4 获取栈顶元素

```python
# 获取栈顶位置
def getTop(self):
    return self.top
```

这里也可以直接使用列表，使用如下：

```python
self.stack[-1]
```



## 3.5 清空栈

```python
# 清空栈
def empty(self):
    self.stack = []
    self.top = -1
```



## 3.6 判断是否为空栈

```python
# 判断是否为空栈
def isEmpty(self):
    if self.top == -1:
        return True
    else:
        return False
```

这里可以直接判断列表是否为空即可，使用如下：

```python
return self.stack is []
```



## 3.7 判断是否为满栈

```python
# 是否为满栈
def isFull(self):
    if self.top == self.size - 1:
        return True
    else:
        return False
```



## 3.8 完整代码

```python
#!usr/bin/env python
# encoding:utf-8

class Stack:
    def __init__(self):
        self.stack = []  # 栈列表
        self.size = 20  # 栈大小
        self.top = -1  # 栈顶位置

    # 元素进栈
    def push(self, element):
        if self.isFull():  # 如果栈满，引发异常
            raise StackException(&#39;Stack is full&#39;)
        else:
            self.stack.append(element)
            self.top &#43;= 1

    # 元素出栈
    def pop(self):
        if self.isEmpty():  # 如果栈为空，则引发异常
            raise StackException(&#39;Stack is empty&#39;)
        else:
            element = self.stack[-1]
            self.top -= 1
            del self.stack[-1]
            return element

    # 获取栈顶位置
    def getTop(self):
        return self.top

    # 清空栈
    def empty(self):
        self.stack = []
        self.top = -1

    # 判断是否为空栈
    def isEmpty(self):
        if self.top == -1:
            return True
        else:
            return False

    # 是否为满栈
    def isFull(self):
        if self.top == self.size - 1:
            return True
        else:
            return False


class StackException(Exception):  # 自定义异常类
    def __init__(self, data):
        self.data = data

    def __str__(self):
        return self.data


&#39;&#39;&#39;
主函数
&#39;&#39;&#39;
if __name__ == &#39;__main__&#39;:
    stack = Stack()  # 创建栈
    for i in range(10):
        stack.push(i)  # 元素进栈
    print(&#39;栈顶位置：&#39;, stack.getTop())
    while not stack.isEmpty():
        print(stack.pop())  # 元素出栈

    stack.empty()  # 清空栈
    for i in range(21):
        stack.push(i)  # 引发异常
```





---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/python/python%E6%95%B0%E6%8D%AE%E7%BB%93%E6%9E%84/3-%E6%A0%88/  

