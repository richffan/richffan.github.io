---
title: 关于 樊刹
date: 2019-08-02T11:04:49+08:00
draft: false
comment: false
pageStyle: wide
lightgallery: true
math: true
aliases: /about/
---

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/hugo-fixit/FixIt?style=flat)](https://github.com/hugo-fixit/FixIt/releases)
[![Hugo](https://img.shields.io/badge/Hugo-%5E0.109.0-ff4088?style=flat&logo=hugo)](https://gohugo.io/)
[![License](https://img.shields.io/github/license/hugo-fixit/FixIt?style=flat)](https://github.com/hugo-fixit/FixIt/blob/master/LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/hugo-fixit/FixIt?style=social)](https://github.com/hugo-fixit/FixIt)
[![GitHub forks](https://img.shields.io/github/forks/hugo-fixit/FixIt?style=social)](https://github.com/hugo-fixit/FixIt/fork)


```goat
          .               .                .               .--- 1          .-- 1     / 1
         / \              |                |           .---+            .-+         +
        /   \         .---+---.         .--+--.        |   '--- 2      |   '-- 2   / \ 2
       +     +        |       |        |       |    ---+            ---+          +
      / \   / \     .-+-.   .-+-.     .+.     .+.      |   .--- 3      |   .-- 3   \ / 3
     /   \ /   \    |   |   |   |    |   |   |   |     '---+            '-+         +
     1   2 3   4    1   2   3   4    1   2   3   4         '--- 4          '-- 4     \ 4
```


{{< mermaid >}}
graph LR;
    A[Hard edge] -->|Link text| B(Round edge)
    B --> C{Decision}
    C -->|One| D[Result one]
    C -->|Two| E[Result two]
{{< /mermaid >}}