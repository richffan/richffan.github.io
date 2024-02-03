# linux下修改按键ESC&lt;=&gt;CAPSLOCK和Control=&gt;ALT_R

使用 `vim` 过程中发现 `esc` 和 `ctrl` 按键很难按，小拇指没有那么长啊～～，而 `caps_lock` 和 `alt_r`(右alt) 很少用。

本教程将 `esc` 和 `caps_lock` 两个按键交换， `alt_r`(右alt) 改为 `ctrl`。

&lt;!--more--&gt;

### 一、 esc 与 caps_lock 按键交换
①. 创建 `.xmodmaprc` 文件。
②. 加入以下内容：
```bash
remove Lock = Caps_Lock
add Lock = Escape
keysym Caps_Lock = Escape
keysym Escape = Caps_Lock
```
③. 执行 `xmodmap .xmodmaprc` 使之生效。
### 二、 将 右alt 改为 ctrl
①. 查看需要修改键位的 keysym
通过 `xev | grep keycode` 获取右 `alt` 的 keysym 为 `Alt_R`。如下图所示：
![通过xev获取右alt的keysym](http://img.saodiyang.com/FvuqjLi5czeBluMTyIfv_xUOcu5k.png)

②. 查看 `Alt_R` 是哪个 modifier 使用的
通过 `xmodmap -pm` 查看，发现 `Alt_R` 是作为 modifier `mod1` 使用的。如下图所示：
![查看 Alt_R 是作为 mode1 使用的](http://img.saodiyang.com/Fib8QjT-Ccx30DCf2rF4WkzHsbOH.png)

③. 修改 modifier
```bash
xmodmap -e &#39;remove mod1 = Alt_R&#39; ## 解除原来绑定
xmodmap -e &#39;add control = Alt_R&#39; ## 作为 control 使用
```


---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/linux/linux%E4%B8%8B%E4%BF%AE%E6%94%B9%E6%8C%89%E9%94%AEesc-capslock%E5%92%8Ccontrol-alt-r/  

