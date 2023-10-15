---
title: " python运行时：ModuleNotFoundError: No module named 'tensorflow'"
categories: ["编程"]
tags: ["python"]
date: 2023-09-25
---

TensorFlow报错：

![](https://gitee.com/wugenqiang/images/raw/master/image/1632384432159.png)

一般的解决方案：

```python
pip install --upgrade --ignore-installed tensorflow
```

或者：

```python
pip install --user --upgrade --ignore-installed tensorflow
```

运行结果如下：

![](https://gitee.com/wugenqiang/images/raw/master/image/1632384620726.png)



