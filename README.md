# 这是一个用Nodejs来操作excel表格的项目（练手）

## 使用class的方式进行编程（面向对象的方式）
1. 文件读读取是有一个excel文件存放位置，通过node内置的fs方法来读取文件流，在通过xlsx插件对excel表格进行读取。
2. 使用了`dayjs`对日期进行计算，主要是为了拿去表格中的月份，然后匹配月份是否匹配天数来捕获边界条件。
3. 使用`lodash`简化计算能力。

## ```注意!!``` 这个脚本只是针对我现在处理的表格数据。 不同的表格，数据处理也不一样。
 不过， 读取文件路径，操作excel等操作请前往`xlsx`文档。可以尝试用js对进行处理 处理完在丢到xlsx生成新的excel表格。