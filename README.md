# 由word模板文档生成pdf文件
* word模板生成pdf
````
ip地址+ /worldTemplate2Pdf
````

* 查看生成的pdf
````
ip地址 +/uploadFiles/test0001%E5%A7%94%E6%89%98%E6%8B%85%E4%BF%9D%E7%94%B3%E8%AF%B7%E4%B9%A6.pdf
````

# 在centos7生成pdf文件异常解决

问题：在windows上运行程序一切正常，但是在centos7中，内容全是##

解决：因为centos系统上缺少字体，导致程序解析内容异常

````
下载中文字体，下载simsun.ttc
并根据下面链接
https://blog.csdn.net/cuiyaonan2000/article/details/82881516
配置字体，重启服务后重新尝试
````
