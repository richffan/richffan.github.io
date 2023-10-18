# sudo命令免密码设置

> 如果某台linux只有自己在使用，比如个人系统，每次调用 `sudo` 时都需要输入密码，长期下来着实厌烦，因此本文介绍如何配置 `sudo` 命令，使其在运行时不需要输入密码。

<!--more-->

### 步骤
1. 执行命令
```bash
$ sudo visudo
```

2. 添加以下两行， 下面的 sys 表示 sys 组成员不用密码使用sudo
```bash
aaronkilik ALL=(ALL) NOPASSWD: ALL
%sys ALL=(ALL) NOPASSWD: ALL
```

现在在使用 `sudo` 命令， 将不再需要输入密码。

### 扩展
如果只允许用户使用 `kill` 和 `rm` 命令时，不需要输入密码，见如下配置
```
%sys ALL=(ALL) NOPASSWD: /bin/kill, /bin/rm
```


---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/posts/program/linux/sudo%E5%91%BD%E4%BB%A4%E5%85%8D%E5%AF%86%E7%A0%81%E8%AE%BE%E7%BD%AE/  

