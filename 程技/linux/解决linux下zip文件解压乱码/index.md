# 解决linux下zip文件解压乱码

### 原因
由于zip格式并没有指定编码格式，Windows下生成的zip文件中的编码是GBK/GB2312等，因此，导致这些zip文件在Linux下解压时出现乱码问题，因为Linux下的默认编码是UTF8。

<!--more-->

### 解决方案
使用7z解压。

1. 安装p7zip和convmv
```bash
## fedora
$ su -c 'yum install p7zip convmv'
## ubuntu
$ sudo apt-get install p7zip convmv
```

2. 执行一下命令解压缩
```bash
## 使用7z解压缩
$ LANG=C 7za x your-zip-file.zip
## 递归转码
$ convmv -f GBK -t utf8 --notest -r .
```


---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/linux/%E8%A7%A3%E5%86%B3linux%E4%B8%8Bzip%E6%96%87%E4%BB%B6%E8%A7%A3%E5%8E%8B%E4%B9%B1%E7%A0%81/  

