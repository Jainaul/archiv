# archiv2.2

日喀则电力公司文件归档小助手2.2

![interface](./config/interface.png)

## 主要功能

1. 自动添加档案章，档案号和序号
2. 自动添加卷内页码
3. 自动生成卷册封面，目录，备考表
4. 自动生成档案盒侧边贴条
5. 自动生成资料交接单

## 使用方法

1. 下载[archiv](https://github.com/rcqed/archiv/releases/download/2.2/archiv_v2.2.7z)（国内下载：[archiv](https://gitee.com/rcqed/archiv/releases/download/2.2/archiv_v2.2.7z)），解压后使用。
3. 将归档好的PDF文件放入``archiv``文件夹中（不要有其它文件），并按"档号__名字"的格式命名卷册名和文件名

```
目录结构如下：
├─archiv
   └─XXXXXX-XXXXXX-XX-XXXX-XXX__卷册名称1
   │   ├─XXXXXX-XXXXXX-XX-XXXX-XXX-XXX__文件名1.pdf
   │   ├─XXXXXX-XXXXXX-XX-XXXX-XXX-XXX__文件名2.pdf
   └─XXXXXX-XXXXXX-XX-XXXX-XXX__卷册名称2
       ├─XXXXXX-XXXXXX-XX-XXXX-XXX-XXX__文件名3.pdf
       ├─XXXXXX-XXXXXX-XX-XXXX-XXX-XXX__文件名4.pdf
```

4. ``双击运行.bat``即可
5. 输出文件在``output``目录下

## 感谢

Python库：openpyxl--3.1.2，PyPDF2--2.10.8，reportlab--4.0.4，python-docx-0.8.11

软件：[WinPython](http://winpython.github.io/)

## 请作者喝奶茶

![support](./config/support.png)
