# java将office文件转换为pdf文件的三种方法

方法1.poi读取doc + itext生成pdf （实现最方便，效果最差，跨平台）
方法2.jodconverter + openOffice （一般格式实现效果还行，复杂格式容易有错位，跨平台）
方法3.jacob + msOfficeWord + SaveAsPDFandXPS （完美保持原doc格式，效率最慢，只能在windows环境下进行）

由于方法1效果比较差,本文只介绍后两种方法

## 方法2：使用jodconverter来调用openOffice的服务来转换，openOffice有个各个平台的版本，所以这种方法跟方法1一样都是跨平台的。

jodconverter的下载地址：http://www.artofsolving.com/opensource/jodconverter
首先要安装openOffice，下载地址：http://www.openoffice.org/download/index.html

## 方法3：效果最好的一种方法，但是需要window环境，而且速度是最慢的需要安装msofficeWord以及SaveAsPDFandXPS.exe(word的一个插件，用来把word转化为pdf)

Office版本是2007，因为SaveAsPDFandXPS是微软为office2007及以上版本开发的插件
SaveAsPDFandXPS下载地址：http://www.microsoft.com/zh-cn/download/details.aspx?id=7
有Microsoft Office软件的可以不安装SaveAsPDFandXPS,Office软件会自带插件
jacob 包下载地址：http://sourceforge.net/projects/jacob-project/