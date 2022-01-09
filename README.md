# Cc2333

为了应对我校教务部要求改实验报告封面而诞生的脚本

## 效果
<div align=center>
<img width="398" alt="截屏2022-01-09 上午10 56 15" src="https://user-images.githubusercontent.com/13096985/148667390-253b0af5-cf22-42b4-b7d4-82f6a93f20c4.png">
  <p>原封面</p>
</div>

<div align=center>
<img width="413" alt="截屏2022-01-09 上午10 54 46" src="https://user-images.githubusercontent.com/13096985/148667363-3087810f-cbdd-4845-a538-7559f65ee9ec.png">
  <p>新封面</p>
</div>

## 原理
使用 python-docx 读取原报告封面信息，并利用 python-docx 以新封面模板构造新的实验报告封面，
新封面和原报告都转换为 pdf 文件，最后将取新封面和原报告除了封面的页面，将它们合并到一个 pdf 中。

并利用一些简单的操作实现了实验信息的居中。

噢对了，这里用的封面模板是被我魔改过的，下划线对齐。

## 依赖安装

```
pip3 install -r requirements.txt
```

## 参数说明

-i 课程编号

-c 班级名称

-l 实验地点

-s 原报告文件夹路径

-o 输出文件夹路径

## 用法示例

```
python3 main.py -i IB01017 -c 2019级物联网五班 -l C5-428 -s ./Python -o ./output
```
