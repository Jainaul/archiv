# archiv

日喀则电力公司文件归档小助手

## 主要功能

1. 自动添加档案章，档案号和序号
2. 自动添加卷内页码
3. 自动生成卷册封面，目录，备考表
4. 自动生成档案盒侧边贴条
5. 自动生成资料交接单

## 使用方式

1. 安装python3.7及以上版本
2. 安装模块：pip install openpyxl PyPDF2 reportlab
3. 将归档好的文件放入archiv文件夹中，并按"档号__名字"的格式命名卷册名和文件名
4. 运行start1.bat，实现1，2功能
5. 运行start2.bat，自动在main，mini，minor文件夹中生成用于VBA的表格
6. 使用工具将excle生成word（例如：excle精灵）
