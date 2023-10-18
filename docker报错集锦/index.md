# docker报错集锦

### docker创建容器

<!--more-->


#### 1. iptables failed
创建 `tale` 容器时，如下命令：
```bash
docker run -d --privileged --hostname tale --name tale \
-v /etc/localtime:/etc/localtime:ro \
-v /home/tale:/var/tale_home -p 127.0.0.1:234:9000 \
-m 1024m --memory-swap -1 tale:1.0
```
然后就报了以下错误：
```bash
docker: Error response from daemon: driver failed programming external connectivity on endpoint tale (263775ff559176224428ec44dcec416a1c20e6c69198d9760b38f35849914260): iptables failed: iptables --wait -t nat -A DOCKER -p tcp -d 127.0.0.1 --dport 234 -j DNAT --to-destination 172.17.0.4:9000 ! -i docker0: iptables: No chain/target/match by that name.
 (exit status 1).
```
解决办法：重启 docker 服务：
```bash
$ service docker restart
```


---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/docker%E6%8A%A5%E9%94%99%E9%9B%86%E9%94%A6/  

