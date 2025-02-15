# wps-excel-embed-image

wps excel embed image

在wps中插入内嵌单元格的图片

# 起因
因个人想法想在excel中插入**内嵌于单元格**的图片

网络上搜索了一下，搜索到了```XlsxWriter```库的 https://xlsxwriter.readthedocs.io/example_embedded_images.html

和微软官网的问答 https://answers.microsoft.com/en-us/msoffice/forum/all/how-to-use-place-in-cell-insert-image-option-in/cc895416-720b-4643-9104-9fdabca83cbf

如页面所说，仅支持office 365 版本

后经过大量搜索没找到答案

偶然看到一个读取wps excel内嵌单元格图片的帖子 https://blog.csdn.net/maudboy/article/details/133145278

后考虑到我这边的使用场景都是用的wps，所以就尝试实现一下

本库的操作原理就是基于此文章
# 原理简述

将excel文件解压，解压出各种xml文件，根据内嵌单元格图片的格式对这些xml文件进行修改，并将图片添加到指定目录，xml中将单元格-图片的对应关系进行关联，操作完毕后将目录重新压缩成excel文件

# 运行示例
创建虚拟环境
```shell
#导入requirements.txt
pip install -r requirements.txt
```
运行main.py脚本

# 建议

仅适用于最终处理阶段，因为嵌入后如果再用代码对excel进行修改则图片会丢失