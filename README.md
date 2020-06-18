# xlxs2csv

### 为了统计上的相关工作而写的一个小工具（针对特定报表，非通用）。主要作用就是将收集好信息的xlsx格式文件按要求转换成csv格式报文。

### 打包成可执行文件使用
- 可以利用pyinstaller打包。
> `pyinstaller -F -w Handler.py`(单个文件模式)
>
> `pyinstaller -D -w Handler.py`(文件夹模式)

### 注意
- 此工具并非xlsx转csv通用工具，是针对特定报表进行了定制的工具，使用前请注意。
- 第一次使用需要初始化ini配置文件，主要就是填写单位统一社会信用代码。（不知道可不填，此处主要为文件命名需要）