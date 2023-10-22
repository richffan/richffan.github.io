# WSL（适用于 Linux 的 Windows 子系统）Kali在线安装


WSL-Kali安装说明：https://www.kali.org/docs/wsl/win-kex/
WSL-Kali软件安装：https://blog.csdn.net/NDASH/article/details/124147241
WSL-Kali报错解决：https://blog.csdn.net/m0_65228980/article/details/124748981
本文已过时，应适用WSL旧版本

PS：不建议使用WSL安装Kali，安装后Kali问题很多，不好解决，使用虚拟机更实用

## 1 安装虚拟镜像

powershell中执行

wsl -l -v

确认用的是version 2，才能开始下一步。

### 什么是WSL

```
WSL——适用于 Linux 的 Windows 子系统可让开发人员直接在 Windows 上按原样运行 GNU/Linux 环境（包括大多数命令行工具、实用工具和应用程序），且不会产生传统虚拟机或双启动设置开销。
```

### WSL的安装

#### Windows环境配置

```
使用版本：Windows 10 （自动更新到最新版） 
```

##### 1、启用Windows功能
控制面板–>程序与功能–>启用或关闭Windows功能，选中 适用于Linux的Windows子系统 和 虚拟机平台。
见官方文档：https://docs.microsoft.com/zh-cn/windows/wsl/install-on-server


##### 2、设置Windows预览体验计划
开始–>设置–>更新和安全–>Windows预览体验计划，如下图。按提示重启电脑。
见官方文档：https://docs.microsoft.com/zh-cn/windows/wsl/install-win10
Windows预览体验计划文档：https://insider.windows.com/en-us/getting-started#flight


#### WSL简单安装（需PowerShell管理员权限）

```
官方文档给出简单安装方法。
官方文档：https://devblogs.microsoft.com/commandline/install-wsl-with-a-single-command-now-available-in-windows-10-version-2004-and-higher/
```

##### 1、查看支持系统列表（默认Ubuntu，在线安装）

```
wsl --list --online
```

##### 2、选中系统安装

```
本文选中 Kali linux 安装
输入命令，等待安装完成。
wsl --install -d kali-linux
```

##### 3、配置用户名密码
```
用户名不支持大写首字母，输入密码是看不见的。
官方文档：https://docs.microsoft.com/zh-cn/windows/wsl/user-support
```

##### 4、成功

## 2 启用systemd

在Kali环境终端输入

echo -e "[boot]\nsystemd=true" | sudo tee -a /etc/wsl.conf

配置完后重启你的电脑。

这一步很重要，这个是0号线程的基础service，没有它的话很多基础服务起不来，会出现很多奇奇怪怪的问题，这里我也卡了很久。

我的是win11用的这个方法，如果是win10是另外一种配置方式，可以网上搜一下。

最后要执行下面的命令，主要是检查service是否起得来，如果命令正常执行才能继续下一步

systemctl list-units --type=service

## 3 安装win-kex
sudo apt update

这里不建议大家换源，使用官方默认的就好，慢点没关系。

安装标准包

sudo apt install kali-linux-default

安装kex

sudo apt install kali-win-kex

kex会把tigerVNC打包安装，但包里面的VNC版本不是最新，要安装最新的。

sudo apt remove tigervnc-standalone-server

sudo apt install tigervnc-standalone-server

然后执行下面的命令验证，如果显示找不到-list这个option说明vnc不是最新，命令正常执行了才能继续下一步。

vncserver -list

## 4 开启文件权限

cd ~

检查有没有.Xauthority这个文件，没有的话就新建一个

touch ~/.Xauthority

虚拟显示器X0权限设置为可读写，验证方式很简单，看看能不能删除它，注意这里是root下的操作

$ rm -rf /tmp/.X11-unix

如果出现Read only file system，执行

$ umount /tmp/.X11-unix

再次执行rm -rf /tmp/.X11-unix，如果正常删除的话就进行下一步

## 5 配置servername
vim /etc/resolv.conf

加上这一行

nameserver your_ip_address

ip是win下的物理网卡地址，ipconfig查看

## 6 启动kex
kex内部打包了两种远程访问方式，一个是tigerVNC，一个是vcxsrv，vcxsrv的方式配置要复杂，而且远程访问很卡顿，不是很推荐，如果文章有热度就再出vcxsrv的教程，如果没有就算了。避免混乱，这里只给tigerVNC的使用方法。

kex --win -s

vncserver成功启动会有下面的打印，里面会分配一个port，第一次的话会提示设置密码，记住这个port和密码，client会用到。

Win-KeX server sessions:

X DISPLAY #     RFB PORT #      RFB UNIX PATH   PROCESS ID #    SERVER

1               5901                            775             Xtigervnc

## 7 Windows环境下载开源的vnc client
https://sourceforge.net/projects/tigervnc/

或者

https://sourceforge.net/projects/turbovnc/

打开软件

ip address为

localhost:port

port 为step 6 出现的数字，比如默认为5901，那IP为 localhost:5901

密码

即step 6初始start vncserver设置的密码。

登陆后就是这个样子了。可以happy了。。。

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/posts/program/system/wsl%E9%80%82%E7%94%A8%E4%BA%8E-linux-%E7%9A%84-windows-%E5%AD%90%E7%B3%BB%E7%BB%9Fkali%E5%9C%A8%E7%BA%BF%E5%AE%89%E8%A3%85/  

