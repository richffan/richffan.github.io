---
title: QEMU - QuickStart
date: 2023-08-06 18:00:00
categories: ["系统"]
tags:
    - qemu
---
<https://cyan-io.github.io/>
## Intro
QEMU是一款跨平台的开源虚拟机工具，因为无GUI配置工具，有一定使用门槛

根据冯诺依曼结构，一个计算机可分为：

- 运算器
- 控制器
- 存储器
- 输入设备
- 输出设备
使用QEMU时，应该尽量往该模型联想

## 创建存储器
创建一个16G的虚拟磁盘文件

```bash
qemu-img create -f qcow2 drive 16G
```

## 虚拟硬件环境
运算器、控制器 -> CPU、各种加速器*
输入、输出设备 -> 网卡、CXL设备等
QEMU本身没有像VMware那样的GUI配置界面，所有配置通过命令行参数指定

```bash
qemu-system-x86_64 -machine q35 \
	-smp 4,sockets=1,cores=4,threads=1 \
	-m 4096
```
现在启动了一个机型为q35，处理器1插槽4核4线程，内存4096MB，使用默认网络配置的虚拟机，由于没有启动盘，并不会启动某个系统

## 利用虚拟硬件环境启动系统
我们已经有了一块硬盘drive作为启动盘，指定给虚拟环境：

```bash
qemu-system-x86_64 -machine q35 \
	-smp 4,sockets=1,cores=4,threads=1 \
	-m 4096 \
	-drive file=drive
```
会发现，指定后并没有什么区别，是因为这个“硬盘”我们没有给它“安装”系统

## 安装系统
参考为物理机安装系统：

物理机硬盘 -> 虚拟磁盘文件drive
安装媒介（写入了镜像的U盘） -> 系统镜像文件

```bash
qemu-system-x86_64 -machine q35 \
	-smp 4,sockets=1,cores=4,threads=1 \
	-m 4096 \
	-drive file=drive \
	-drive file=<系统镜像，如ubuntu-22.04.2-live-server-amd64.iso>,media=cdrom
```
按照ISO文件的引导进行安装即可，完成后可移除ISO文件重启