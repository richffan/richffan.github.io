# 使用Git将文件夹上传到Github以及使用Git LFS上传大文件


GIT


<!--more-->


## 一、上传普通大小文件
1.安装和配置Git
2.在Github上新建一个仓库，并复制下地址
3.找一个空文件，右键-Git Bash Here

git clone 复制下的仓库地址
1
4.进入克隆下来的仓库的文件夹

cd 仓库的文件夹名
1
再将要上传的文件夹粘贴到这里面
5.添加文件到暂存区

git add .
1
6.将暂存区内容添加到仓库中

git commit -m "写说明信息"
1
7.上传远程代码并合并

git push -u origin main
1
再输入账号密码，等待即可

## 二、使用Git LFS上传大文件
1.安装和配置Git
2.安装git lfs: windows installer，并运行（Windows端）
3.在Github上新建一个仓库，并复制下地址
4.找一个空文件，右键-Git Bash Here
5. 开启lfs功能

git lfs install
1
6.使用 git lfs track 命令进行大文件追踪 例如git lfs track “*.png” 追踪所有后缀为png的文件

git lfs track "*.文件类型"
1
7.查看现有的文件追踪模式

git lfs track
1
8.将管理文件.gitattributes添加入git仓库

git add .gitattributes
1
提交代码需要将gitattributes文件提交至仓库，它保存了文件的追踪记录
9.显示当前跟踪的文件列表

git lfs ls-files 
1
将代码 push 到远程仓库后，LFS 跟踪的文件会以『Git LFS』的形式显示:
10.克隆仓库

git lfs clone 复制下的仓库地址
1
11.进行配置

git config lfs.复制下的仓库地址/info/lfs.locksver ify false
1
再把postBuffer的值配置成500M，可以根据你需要下载的文件大小，将postBuffer值配置成合适的大小。

git config --global http.postBuffer 524288000
1
12.进入克隆下来的仓库的文件夹

cd 仓库的文件夹名
1
再将要上传的文件夹粘贴到这里面
13. 添加文件到暂存区

git add .
1
文件较大可能会等待时间较久
14.将暂存区内容添加到仓库中

git commit -m "写说明信息"
1
15.上传远程代码并合并

git push -u origin main
1
再输入账号密码，等待即可

## 三、取消代理设置
git config --global --unset http.proxy
git config --global --unset httpx.proxy
1
2
打开C:\Users\当前用户里的.gitconfig文件(这个默认是隐藏文件),将里面关于proxy的行删掉

## 四、常见问题
1.上传远程代码并合并

git push -u origin main
1
报错：

error: src refspec main does not match any
error: failed to push some refs to 'https://github.com...
1
2
原因：
本地 GIT 客户端创建了一个默认分支，称为master（git init初始化存储库时），但 GitHub 上的远程存储库没有master，而是默认分支称为main。
解决方法：
运行

git push -u origin main
1
之前先

git checkout -B main
1
或者改为master

git pull origin master


---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/blog/%E4%BD%BF%E7%94%A8git%E5%B0%86%E6%96%87%E4%BB%B6%E5%A4%B9%E4%B8%8A%E4%BC%A0%E5%88%B0github%E4%BB%A5%E5%8F%8A%E4%BD%BF%E7%94%A8git-lfs%E4%B8%8A%E4%BC%A0%E5%A4%A7%E6%96%87%E4%BB%B6/  

