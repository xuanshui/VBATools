## 改动：

#### 0、代码优化思路

1. readDataFromExcel函数用三层循环来遍历所有单元格，这个操作是很耗时的，可以把三层循环削减为两层，第三层使用Range("e1:e4") = Application.Transpose(d.Keys)    Range("f1:f4") = Application.Transpose(d.Items)这样类似的语句去获取内容

##### 一、全局变量

- 增加g_arr_num_excel_colInSheets()：对应g_arr_num_excel_lineInSheets()，即Excel表格中每个sheet实际使用的列数
- 增加g_arr_testCase_key()：存放每个sheet的第一行内容，从而作为临时字典tmp_dic_testCase的键

##### 二、getExcelLineInfo函数

1. 增加初始化填写g_arr_num_excel_colInSheets的代码

```
 g_arr_num_excel_colInSheets(tmp) = g_JL_Excel.worksheets(tmp).usedrange.Columns.count 
```

##### 三、readDataFromExcel函数

1. 在第一层遍历sheet时，赋值全局数组g_arr_testCase_key()来存放每个sheet的第一行内容
1. 在第三层遍历colNum时，当遍历到第二列测试用例编号时，完善对第一行特殊情况的处理：If rowNum >= C_excel_start_row Then       '从表格的用例起始行开始，临时字典中才保存有用例信息
1. 第三层遍历中，realCaseNumInSubItem = realCaseNumInSubItem + 1 '在测试项编号为空的情况下，设定实际用例编号+1
