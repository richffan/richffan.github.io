# Golang入门笔记-CH04-Go语言流程控制

## 前言

流程控制是每门语言控制程序逻辑和执行顺序的重要组成部分，Go 语言中常见的流程控制有 `if`，`for`，`switch`；`break`、`continue` 和 `goto` 是为了简化流程控制，降低代码复杂度。

## if-else

`if` 分支结构的基本写法为：

&gt; `if` 分支结构多用于条件判断。

```go
if 表达式1 {
    分支1
} else if 表达式2 {
    分支2
} else {
    分支3
}
```

上述代码中，若**表达式1**的值为 `true`，程序将会执行**分支1**；若**表达式1**的值为 `false`，继续判断**表达式2**，若**表达式2**为 `true`，将会执行**分支2**；若**表达式1**和**表达式2**都为`false`，将会执行**分支3**。

`if-else` 分支结构会逐层判断表达式是否为 `true`，若为 `true`，则执行该表达式中对应的分支，否则继续判断下一个表达式，依次类推。

我们来看一个例子：

```go
package main

import &#34;fmt&#34;

func testIf(score int) {
	if score &gt;= 90 {
		fmt.Println(&#34;优秀&#34;)
	} else if score &gt;= 75 {
		fmt.Println(&#34;良好&#34;)
	} else if score &gt; 60 {
		fmt.Println(&#34;及格&#34;)
	} else {
		fmt.Println(&#34;不及格&#34;)
	}
}

func main() {
	testIf(90)
}
```

运行结果为:

```shell
优秀
```

`if` 语句还有一个特殊的用法：

&gt; 可以在判断条件之前执行一条语句，例如赋值等。

```go
if a := 10; a &gt; 6 {
    fmt.Println(a)
}
```

## for

`for` 语句常用于循环，例如循环遍历字符串、数组、切片和Map等类型数据，Go 语言中没有`while`循环语句。

`for` 语句的基本格式如下：

```go
for 初始语句;条件表达式;结束语句 {
    循环体语句
}
```

当 `for` 循环中条件表达式为`true`，循环会继续执行；否则，循环终止。

```go
package main

import &#34;fmt&#34;

func main() {
	for i := 0; i &lt; 5; i&#43;&#43; {
		fmt.Println(i)
	}
}
```

上述代码的执行结果为：

```shell
0
1
2
3
4
```

`for` 循环中的初始化语句可以省略，但 `;` 分号不能省略：

```go
package main

import &#34;fmt&#34;

func main() {
    i := 0
	for ; i &lt; 5; i&#43;&#43; {
		fmt.Println(i)
	}
}
```

`for`循环中的初始化语句和结束语句都可以省略，类似于 C 语言中的 `while` 语句：

```go
package main

import &#34;fmt&#34;

func main() {
    i := 0
	for i &lt; 5 {
		fmt.Println(i)
		i&#43;&#43;
	}
}
```

我们可以通过省略 `for` 循环中的初始化语句，条件表达式和结束语句来实现**无限循环**：

```go
for {
    循环体语句
}
```

## for range

`for range` 是**键值循环**，一般用于遍历字符串、数组、切片、map 和 channel。

例如，遍历数组：

```go
package main

import &#34;fmt&#34;

func main() {
	nameList := [3]string{&#34;张三&#34;, &#34;李四&#34;, &#34;赵二&#34;}
	genderMap := map[string]string{
		&#34;0&#34;: &#34;男&#34;,
		&#34;1&#34;: &#34;女&#34;,
	}
	
    	// k 是键，v 是值，
	for k, v := range nameList { // 对于数组而言，k 就是下标
		fmt.Println(k, v)
	}

	for k, v := range genderMap {
		fmt.Println(k, v)
	}
}
```

运行结果为：

```shell
0 张三
1 李四
2 赵二
0 男
1 女
```

## switch

`switch` 语句和 `if` 语句类似，一般用于多条件判断，且这些条件易于枚举：

&gt; 每个 `switch` 语句只能有一个 `default` 分支。 

```go
package main

import &#34;fmt&#34;

func main() {
	finger := 3
	
	switch finger {
	case 1:
		fmt.Println(&#34;大拇指&#34;)
	case 2:
		fmt.Println(&#34;食指&#34;)
	case 3:
		fmt.Println(&#34;中指&#34;)
	case 4:
		fmt.Println(&#34;无名指&#34;)
	case 5:
		fmt.Println(&#34;小拇指&#34;)
	default:
		fmt.Println(&#34;无效的输入！&#34;)
	}
}
```

一个分支可以有多个值：

```go
package main

import &#34;fmt&#34;

func main() {
	n := 3

	switch n {
	case 1, 3, 5, 7, 9:
		fmt.Println(&#34;奇数&#34;)
	case 0, 2, 4, 6, 8:
		fmt.Println(&#34;偶数&#34;)
	default:
		fmt.Println(&#34;无效的输入！&#34;)
	}
}
```

`switch` 的 `case` 中还可以使用表达式，一旦使用表达式，`switch` 后面不需要填变量：

```go
package main

import &#34;fmt&#34;

func main() {
	n := 60

	switch {
	case n &gt;= 90:
		fmt.Println(&#34;优秀&#34;)
	case n &gt;= 75 &amp;&amp; n &lt; 90:
		fmt.Println(&#34;良好&#34;)
	case n &gt;= 60 &amp;&amp; n  &lt; 75:
		fmt.Println(&#34;及格&#34;)
	case n &lt; 60:
		fmt.Println(&#34;不及格&#34;)
	default:
		fmt.Println(&#34;成绩无效&#34;)
	}
}
```

Go 语言中还保留了 `fallthrough` ，主要是为了兼容 C 语言，`fallthrough` 可以继续执行满足条件的 `case` 的下一个 `case`：

```go
package main

import &#34;fmt&#34;

func main() {
	s := &#34;a&#34;

	switch {
	case s == &#34;a&#34;:
		fmt.Println(&#34;a&#34;)
		fallthrough
	case s == &#34;b&#34;:
		fmt.Println(&#34;b&#34;)
	case s == &#34;c&#34;:
		fmt.Println(&#34;c&#34;)
	default:
		fmt.Println(&#34;...&#34;)
	}
}
```

运行结果为：

```go
a
b
```

## break

`break` 用于主动跳出循环，例如：

```go
package main

import &#34;fmt&#34;

func main() {
	for i := 0; i &lt; 5; i&#43;&#43; {
		if i == 3 {
			break
		}
		fmt.Println(i)
	}
}
```

上述代码中，当 `i`  的值为 3 时，将会跳出循环，所以只会打印 0，1 和 2。

## continue

`continue` 用于跳过这次循环，继续执行下一次循环，注意和 `break`  语句的区别。

```go
package main

import &#34;fmt&#34;

func main() {
	for i := 0; i &lt; 5; i&#43;&#43; {
		if i == 3 { // 当 i 值为 3 时，跳过这次循环，执行下一次循环，因此 3 不会被打印
			continue
		}
		fmt.Println(i)
	}
}
```

运行结果为：

```shell
0
1
2
4
```

## goto

我们可以用 `goto` 语句跳转到指定标签，来简化代码。但同时，代码的可读性会下降，一般在编码中，尽量不要使用 `goto` 语句。

在双层 `for` 循环中，如果我们使用 `break` 来退出循环，可以定义一个标识变量来标记是否要退出：

```go
package main

import &#34;fmt&#34;

func main() {
	var breakFlag bool
	for i := 0; i &lt; 10; i&#43;&#43; {
		for j := 0; j &lt; 10; j&#43;&#43; {
			if j == 2 {
				// 设置退出标签
				breakFlag = true
				break
			}
			fmt.Printf(&#34;%v-%v\n&#34;, i, j)
		}
		// 外层 for 循环判断
		if breakFlag {
			break
		}
	}
}
```

上述代码运行结果为：

```shell
0-0
0-1
```

我们可以用 `goto` 语句来简化代码，直接跳转到指定标签即可：

```go
package main

import &#34;fmt&#34;

func main() {
	for i := 0; i &lt; 10; i&#43;&#43; {
		for j := 0; j &lt; 10; j&#43;&#43; {
			if j == 2 {
				// 设置退出标签
				goto breakTag
			}
			fmt.Printf(&#34;%v-%v\n&#34;, i, j)
		}
	}
	return
	// 标签
breakTag:
	fmt.Println(&#34;结束 for 循环&#34;)
}
```

上述代码运行结果为：

```shell
0-0
0-1
结束 for 循环
```

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/golang/golang%E5%85%A5%E9%97%A8%E7%AC%94%E8%AE%B0-ch04-go%E8%AF%AD%E8%A8%80%E6%B5%81%E7%A8%8B%E6%8E%A7%E5%88%B6/  

