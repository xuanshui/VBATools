Attribute VB_Name = "JL_EXCL_WD"
'-----------------------------------------------------------------------------------------------------------------------------------------
'version:       1.0.0
'author:        lihong
'update date:   20230902
'function：
'1、从Excel版测试记录提取所有的测试用例信息
'2、检查Excel版测试记录中的部分低级错误
'3、把测试用例信息写入Word版的测试记录（最好是通过另一个脚本自动生成的测试记录）
'
'
'
'-----------------------------------------------------------------------------------------------------------------------------------------

Option Base 1
Option Explicit '强制程序中的变量声明为显式声明，即必须使用Dim或者ReDim诺包声明交量

'-----------------------------------------------------------------------------------------------------------------------------------------
'           ↓文件操作
'-----------------------------------------------------------------------------------------------------------------------------------------
Dim g_JL_Excel As Object
Dim g_JL_Word As Document
Public ExcelApp As Object
Public WordApp As Object
Dim workPath As String
'-----------------------------------------------------------------------------------------------------------------------------------------
'           ↓Excel表格操作
'-----------------------------------------------------------------------------------------------------------------------------------------
Dim g_arr_num_excel_rowInSheets() As Long
Dim g_arr_num_excel_colInSheets() As Long
Dim g_num_excel_actual_row As Long
Dim g_num_excel_actual_column As Integer
'-----------------------------------------------------------------------------------------------------------------------------------------
'           ↓数据结构：存放测试用例的所有信息
'-----------------------------------------------------------------------------------------------------------------------------------------
Public g_arr_testCase() As String
Public g_dic_subItem_testCase As Object
Public g_arr_testCase_key() As String                 '键数组：每个sheet的第一行的每个单元格内容就是每个用例字典的键
'-----------------------------------------------------------------------------------------------------------------------------------------
'           ↓错误统计
'-----------------------------------------------------------------------------------------------------------------------------------------
Dim g_num_err_excel As Integer
Dim g_arr_err_excel() As String
'-----------------------------------------------------------------------------------------------------------------------------------------
'           ↓可选参数
'-----------------------------------------------------------------------------------------------------------------------------------------
Const C_excel_start_row As Integer = 2      'Excel测试记录的测试用例的实际起始行
Const C_opt As Boolean = False              '脚本速度优化开关

Const C_word_rowNum_caseName As Integer = 1         'Word表格中的“测试用例名称”在第一行
Const C_word_rowNum_caseCode As Integer = 2         'Word表格中的“测试用例标识”在第二行
Const C_word_rowNum_caseInit As Integer = 3         'Word表格中的“测试用例初始化”在第二行
Const C_word_rowNum_premise  As Integer = 6         'Word表格中的“前提与约束”在第六行
Const C_word_rowNum_caseStart As Integer = 10       'Word表格中的测试用例实际内容从第十行开始
Const C_word_rowNum_designer As Integer = 11        'Word表格中的“设计人员”在第十一行
Const C_word_rowNum_tester As Integer = 12          'Word表格中的“测试人员”在第十二行

Const C_excel_fileName As String = "测试记录.xlsx"
Const C_word_fileName As String = "测试记录.doc"
Const C_word_fileName_saved As String = "测试记录-脚本填写.doc"
'Excel测试记录的列名与对应的列序号
Public Enum E_colNum_JLExcel
    subItemCode = 2                     '第2列：测试子项的标识号
    caseNumInSubItem = 3                '第3列：每个测试子项下的测试用例序号
    caseName = 4                        '第4列：测试用例名称
    casePass = 8                        '第8列：测试用例是否通过
End Enum
'Excel测试记录本身的错误情况
Private Enum E_errCode_JL
    'Excel测试记录本身的错误情况
    subItemCode_Null             '"测试项编号为空"
    subItemCode_Repeat           '"测试项编号重复"
    caseNumInSubItem_NotNum      '"用例编号不是数字"
    caseNumInSubItem_Diff        '"用例编号与实际不一致"
    caseName_Null                '"用例名称为空"
    caseName_SeqNull             '"用例名称的序号为空"
    caseName_SeqNotNum           '"用例名称的序号不是数字"
    caseName_SeqDiff             '"用例名称的序号与实际不一致"
    casePass_Null                '"测试用例的结论为空"
    casePass_Wrong               '"测试用例的结论不是"通过"或"未通过"或"不通过"或"建议改进"
    
    '从Excel测试记录获取信息后，填入Word测试记录过程时的错误情况
    subItemNum_Diff              'Excel和Word两个版本的测试记录中的测试子项数量不相同
End Enum

'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------
'           主函数
'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------
Sub JL_EXCL_WD()
    '0、捕获异常，错误处理
    'On Error GoTo errhandlemain

    '1、初始化函数，进行提示框、判断版本、设置光标和视图等
    init

    '2、检查Excel版测试记录中的错误，写入一个txt文件，然后提示测试人员修改Excel表格
    checkErr

    '3、获取excel表格中的测试用例信息，填入数组g_dic_subItem
    readDataFromExcel
    
    '4、关闭且不保存Excel版的测试记录
    'Call printDic2(g_dic_subItem_testCase)
    '5、填写word版的测试记录
    writeDataToWord

    '6、关闭和保存相关文件，另存word版的测试记录
    saveJL

    Exit Sub

errhandlemain:
    Debug.Print Chr(13) & "----------- Error Happend! ----------- " & Chr(13)
    Debug.Print "错误 " & Err.Number & " : " & Err.Description & Chr(13) & "错误定位 : " & Err.Source
    If Err.Number = 4608 Then
    Else
        'err.raise err.number
    End If
    Debug.Print Chr(13) & "-----------   Error  End   ----------- " & Chr(13)
    Resume Next
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
'           1、Init：初始化
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub init()

    workPath = ThisDocument.Path
    
    Set ExcelApp = CreateObject("excel.application")
    'Set WordApp = CreateObject("word.application")
    
    '3、打开测试记录
    Set g_JL_Excel = ExcelApp.workbooks.Open(workPath & "\" & C_excel_fileName)
    ExcelApp.Visible = True
    Set g_JL_Word = Application.Documents.Open(workPath & "\" & C_word_fileName)
    
    
    g_JL_Word.Activate
    Selection.HomeKey wdStory
    
    '5、可选优化：将Word版记录转为普通视图，省去Word分页，提高运行速度
    If C_opt Then
        g_JL_Word.ActiveWindow.View.Type = wdNormalView
    End If
    
    '6、获取Excel表格行数、列数信息，然后根据Excel实际行数初始化测试用例的数组、全局字典第二层字典的键数组
    getExcelLineInfo
    ReDim g_arr_testCase(g_num_excel_actual_row, g_num_excel_actual_column)
    ReDim g_arr_testCase_key(g_JL_Excel.worksheets.count, g_num_excel_actual_column)
    
    '7、初始化全局字典
    Set g_dic_subItem_testCase = CreateObject("scripting.dictionary")
    
    '8、初始化错误参数
    g_num_err_excel = 0
    
End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
'           2、checkErr：检查Excel版测试记录中的错误，写入一个txt文件，然后提示测试人员修改Excel表格
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub checkErr()



End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
'           3、readDataFromExcel：获取excel表格中的测试用例信息，填入数组g_dic_subItem
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub readDataFromExcel()

    '1、优化选项
    If C_opt Then
        ExcelApp.ScreenUpdating = False          '脚本操作Exce1时不会刷新屏幕
        Application.ScreenUpdating = False       '脚本操作word时不会刷新屏幕
    End If
    
    
    '2、定义变量
    Dim sheetNum, rowNum, colNum As Long         '遍历表格用的三层计数
    Dim caseNum As Long: caseNum = 0             '测试用例总计数
    Dim cellStr As String                        '暂存保存单元格的内容
    Dim subItemCode As String                    '上一个测试用力的测试子项标识号
    'Dim tmp_dic_subItemCode_repeat As Object              '重复的测试用力的测试子项标识号
    'Dim subItemCodeRepeatRow As Long             '在第几行重复
    'Dim caseNumInSubItem As String              'Excel表格中测试用例的计数，可能是错误的
    Dim realCaseNumInSubItem As Integer          '正确的测试子项中的测试用例的计
    Dim tmp_arr_caseNumInSubItem() As String     '存放字符“-”分割用例名称后的数组
    Dim tmp_numInCaseName As String              '存放用例名称“-”后的序号
    '字典相关
    Dim tmpKey, tmpVal As String                 '存放临时字典的键、值
   
    
    
    
    '3、遍历Excel所有sheet、所有row、所有column
'----------------------第一层循环------------------------------------------------------------------------------------------
    For sheetNum = 1 To g_JL_Excel.worksheets.count
        '3-1 获取第一行的内容，填入键数组g_arr_testCase_key
        'ReDim g_arr_testCase_key(g_JL_Excel.worksheets(sheetNum).usedrange.Columns.count)
        For colNum = 1 To g_arr_num_excel_colInSheets(sheetNum)
            cellStr = Escape_Text(g_JL_Excel.worksheets(sheetNum).Cells(1, colNum).Value)
            '键值为空时，使用列序号作为键
            If cellStr = "" Then
                g_arr_testCase_key(sheetNum, colNum) = colNum
            Else
                g_arr_testCase_key(sheetNum, colNum) = cellStr
            End If
        Next colNum
        '重置重复错误的临时字典
        'Set tmp_dic_subItemCode_repeat = CreateObject("scripting.dictionary")
'----------------------第二层循环------------------------------------------------------------------------------------------
        For rowNum = C_excel_start_row To g_arr_num_excel_rowInSheets(sheetNum) '从每个sheet的实际起始行开始 '从表格的用例起始行开始，字典中才保存有用例信息

            caseNum = caseNum + 1 '统计总行数，作为测试用例数组的第一个下标
            Dim tmp_dic_testCase As Object
            Set tmp_dic_testCase = CreateObject("scripting.dictionary")  '如果要重复利用某一个临时的字典变量，那就必须在每次使用前set一下，才能清空字典的内容，否则会出错
'----------------------第三层循环------------------------------------------------------------------------------------------
            For colNum = 1 To g_arr_num_excel_colInSheets(sheetNum)
                '3-2 获取单元格内容
                cellStr = g_JL_Excel.worksheets(sheetNum).Cells(rowNum, colNum).Value '存储原始字符串，后续根据不同的列进行处理
                
                '3-3 获取临时字典tmp_dic_testCase的键
                tmpKey = g_arr_testCase_key(sheetNum, colNum)
                
                '3-4 根据不同的列，进行不同的处理：获取信息，判断错误
                Select Case colNum
                
                    '1、测试项编号（测试子项标识号）
                    Case E_colNum_JLExcel.subItemCode
                        tmpVal = Escape_Text(cellStr)
                        '1-1 ---------发生错误1---------：测试项编号为空
                        If tmpVal = "" Then
                            tmpVal = subItemCode    '在测试项编号为空的情况下，继续使用上一条测试用例的测试项编号
                            realCaseNumInSubItem = realCaseNumInSubItem + 1 '在测试项编号为空的情况下，设定实际用例编号+1
                            recErrInfo E_errCode_JL.subItemCode_Null & "测试项编号为空", sheetNum, rowNum, colNum        '记录错误信息
                            
                        '1-2 当前测试项编号与上一条用例的测试项编号相同：即当前用例和上一条用例是同一个测试子项的用例
                        ElseIf cellStr = subItemCode Then
                            realCaseNumInSubItem = realCaseNumInSubItem + 1 '实际用例编号+1

                        '1-4 当前测试项编号与上一条用例的测试项编号不同：上一条测试子项的所有用例已统计完，进入下一条测试子项的用例，写入信息到全局字典
                        Else
                            '每次进入下一个测试子项时，都重新初始化测试子项的临时字典
                            Dim tmp_dic_subItem As Object
                            '重置实际用例编号
                            realCaseNumInSubItem = 1
                            If rowNum > C_excel_start_row Then  '要等脚本运行到内容起始行的下一行，这样临时字典中才会有数据，才能进行写入
                                '--------------------------------------------
                                '本条用例的测试项编号重复不影响上一条测试用例，之前的测试子项的所有用例信息已经统计完成，正常写入全局字典
                                '--------------------------------------------
                                g_dic_subItem_testCase.Add subItemCode, tmp_dic_subItem
                            End If
                            
                            ' ---------发生错误2---------：测试项编号重复：当前这条用例的测试项编号重复了，把subItemCode加上“重复”标记，但是上面的测试项编号是没问题的，正常写入
                            If g_dic_subItem_testCase.Exists(tmpVal) Then
                               
                                recErrInfo E_errCode_JL.subItemCode_Repeat & "测试项编号重复", sheetNum, rowNum, colNum        '记录错误信息
                                '脚本不进行修正，测试人员自行修改
                                
'                                '对重复测试项编号的处理-----
'                                If tmp_dic_subItemCode_repeat.Exists(tmpVal) Then
'                                    'tmp_dic_subItemCode_repeat(tmpVal) = tmp_dic_subItemCode_repeat(tmpVal) + 1
'                                    lastRepeatRowNum = rowNum
'                                Else
'                                    tmp_dic_subItemCode_repeat.Add tmpVal, rowNum
'                                    lastRepeatRowNum = rowNum
'                                End If
'                                If rowNum = lastRepeatRowNum Then
'                                    subItemCode = tmpVal & "-重复-" & tmp_dic_subItemCode_repeat(tmpVal)
'                                End If
                                subItemCode = cellStr & "-重复-" & g_num_err_excel
                                
                            '不重复的测试项编号，正常情况的分支
                            Else

                                '写入全局字典后，更新subItemCode为当前的测试项编号，用于判断下一条用例是否还属于同一个测试子项
                                subItemCode = tmpVal
                            End If
                            
                            '1-3-2 每次进入下一个测试子项时，都重新初始化测试子项的临时字典
                            Set tmp_dic_subItem = CreateObject("scripting.dictionary")
                        End If
                        
                    '2、用例编号
                    Case E_colNum_JLExcel.caseNumInSubItem
                        tmpVal = Escape_Text(cellStr)
                        ' ---------发生错误3---------：用例编号不是数字
                        If Not IsNumeric(tmpVal) Then
                            recErrInfo E_errCode_JL.caseNumInSubItem_NotNum & "用例编号不是数字", sheetNum, rowNum, colNum        '记录错误信息
                            '脚本不进行修正，测试人员自行修改
                        ' ---------发生错误4---------：用例编号与实际序号不一致
                        ElseIf realCaseNumInSubItem <> CInt(tmpVal) Then
                            recErrInfo E_errCode_JL.caseNumInSubItem_Diff & "用例编号与实际不一致", sheetNum, rowNum, colNum         '记录错误信息
                            '脚本不进行修正，测试人员自行修改
                        End If
                        
                    '3、用例名称
                    Case E_colNum_JLExcel.caseName
                        tmpVal = Escape_Text(cellStr)
                         ' ---------发生错误5---------：'用例名称为空
                        If tmpVal = "" Then
                            recErrInfo E_errCode_JL.caseName_Null & "用例名称为空", sheetNum, rowNum, colNum           '记录错误信息
                            '脚本不进行修正，测试人员自行修改
                        Else
                            tmp_arr_caseNumInSubItem = Split(tmpVal, "-")
                            tmp_numInCaseName = tmp_arr_caseNumInSubItem(UBound(tmp_arr_caseNumInSubItem))  '获取用例名称里“-”后面的序号
                            ' ---------发生错误5---------：用例名称后面没有加“-”，即用例名称后没有序号
'                            If UBound(tmp_arr_caseNumInSubItem) = -1 Then
                            If InStr(tmpVal, "-") = 0 Then
                                recErrInfo E_errCode_JL.caseName_SeqNull & "用例名称的序号为空", sheetNum, rowNum, colNum          '记录错误信息
                                '错误修正
                                tmpVal = tmpVal & "-" & realCaseNumInSubItem
                            ' ---------发生错误6---------：用例名称“-”后面的不是数字
                            ElseIf Not IsNumeric(tmp_numInCaseName) Then
                                recErrInfo E_errCode_JL.caseName_SeqNotNum & "用例名称的序号不是数字", sheetNum, rowNum, colNum           '记录错误信息
                                tmpVal = getSubItemName(tmpVal) & "-" & realCaseNumInSubItem '脚本修正
                                
                            ' ---------发生错误7---------：用例名称“-”后的序号与实际不一致
                            ElseIf realCaseNumInSubItem <> CInt(tmp_numInCaseName) Then
                                recErrInfo E_errCode_JL.caseName_SeqDiff & "用例名称的序号与实际不一致", sheetNum, rowNum, colNum            '记录错误信息
                                tmpVal = getSubItemName(tmpVal) & "-" & realCaseNumInSubItem '脚本修正
                                
                            End If
                        End If
                        
                    '4、是否通过
                    Case E_colNum_JLExcel.casePass
                        tmpVal = Escape_Text(cellStr)
                         ' ---------发生错误8---------：测试结论为空
                        If tmpVal = "" Then
                            recErrInfo E_errCode_JL.casePass_Null, sheetNum & "测试用例的结论为空", rowNum, colNum             '记录错误信息
                            '脚本不进行修正，测试人员自行修改
                         ' ---------发生错误9---------：测试用例的结论必须是"通过"或"未通过"或"建议改进"
                        ElseIf tmpVal <> "通过" And tmpVal <> "未通过" And tmpVal <> "不通过" And tmpVal <> "建议改进" Then
                            recErrInfo E_errCode_JL.casePass_Wrong & "测试用例的结论不是""通过""或""未通过""或""不通过""或""建议改进""", sheetNum, rowNum, colNum             '记录错误信息
                            '脚本不进行修正，测试人员自行修改
                        End If
'
                    '5、除去上面的4列，其他列不判断错误，直接使用原始内容
                    Case Else
                        tmpVal = cellStr
                        
                End Select
                '--------------------------------------------
                '测试用例的一列信息已处理，将其写入临时字典tmp_dic_testCase
                '--------------------------------------------
                tmp_dic_testCase.Add tmpKey, tmpVal
            Next colNum
'----------------------第三层循环------------------------------------------------------------------------------------------

        '--------------------------------------------
        '一条测试用例的所有信息已经写入临时字典tmp_dic_testCase，将tmp_dic_testCase写入更高一层的临时字典tmp_dic_testCase
        '--------------------------------------------
        'Call printDic1(tmp_dic_testCase)
        tmp_dic_subItem.Add realCaseNumInSubItem, tmp_dic_testCase
        '当脚本已经遍历到每个sheet的最后一行时，此时脚本就到达了最后一条测试用例，直接把当前的临时字典tmp_dic_subItem放入全局字典
'        If sheetNum = g_JL_Excel.worksheets.count And rowNum = g_arr_num_excel_rowInSheets(g_JL_Excel.worksheets.count) Then
        If rowNum = g_arr_num_excel_rowInSheets(g_JL_Excel.worksheets.count) Then
            '--------------------------------------------
            '到达整个测试用例的最后一条用例，直接将临时字典tmp_dic_subItem写入全局字典
            '--------------------------------------------
            g_dic_subItem_testCase.Add subItemCode, tmp_dic_subItem
        End If
        Next rowNum
        'Call printDic2(tmp_dic_subItem)
'----------------------第二层循环------------------------------------------------------------------------------------------
    Next sheetNum
'----------------------第一层循环------------------------------------------------------------------------------------------
    'Call printDic3(g_dic_subItem_testCase)
    '检验二层字典是否存入了预期的数据
    
End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
'           4、writeDataToWord：填写word版的测试记录
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub writeDataToWord()

'    init
'    readDataFromExcel

    ' ---------发生错误---------：测试记录的Excel(脚本统计-全局字典)和Word(表格数量)两个版本中的测试子项数量不相同
    If ActiveDocument.Tables.count <> g_dic_subItem_testCase.count Then
        recErrInfo E_errCode_JL.subItemNum_Diff & "测试记录的Excel(脚本统计-全局字典)和Word(表格数量)两个版本中的测试子项数量不相同"         '记录错误信息
    End If
    
    '思路：
    '1、根据实际测试内容（即Excel表格内容）去填写Word
    '2、可能会产生两个严重问题：一是Excel有而Word无（可能是Excel的测试项编号写错），二是Word有Excel无（Excel漏测漏写）
    '具体实现：
    '1、根据全局字典，在Word的表格中以此搜索测试用例标识，1、如果测试用例标识一致，则进行填写，2、否则认为Excel的测试项编号写错，不填写这条测试子项的用例，记录错误消息，测试人员手动填写
    '2、遍历完全局字典后，再遍历Word的全部表格，搜索“输入及操作说明”这一栏下的单元格，1、如果全都有内容，说明Excel没有漏测漏写，2、否则有空白的话，则认为Excel存在漏测漏写，记录错误信息
    
    'Word表格操作的相关变量
    Dim tbl As table '当前正在浏览的Word表格
    Dim tblNum As Long 'Word表格序号
    'Dim cellStr As String  'Word表格的单元格内容
    
    '记录错误
    Dim dic_errSubItem_Excel As Object: Set dic_errSubItem_Excel = CreateObject("scripting.dictionary")     '存放Excel中没法写入Word的测试子项（包括其下的多个测试用例）：Excel的测试子项标识号写错
    Dim dic_errSubItem_Word As Object: Set dic_errSubItem_Word = CreateObject("scripting.dictionary")       '存放Word有但是Excel没有的测试子项：Excel版测试记录漏测漏写
    
    '
    Dim arr_subItemCode() As Variant: arr_subItemCode = g_dic_subItem_testCase.keys       '所有的测试子项标识号
    
'    Dim tmp_arr_caseName() As String  '存放分割测试用例名称的数组
    Dim excelCaseName As String 'Excel测试用例中的测试用例名称
    
    'Word表格内容
    Dim caseName As Range  '测试用例名称
    Dim caseCode As Range  '测试用例标识
    Dim caseInit As Range  '测试用例初始化
    Dim premise As Range  '前提与约束
    'Dim caseStart As Range  '测试用例实际内容
    Dim caseInput As Range  '输入
    Dim caseExpectedOutcome As Range    '期望结果
    Dim caseCert As Range       '评估准则
    Dim caseActualResult As Range   '实际测试结果
    Dim isPass As Range             '执行结果
    Dim designer As Range  '设计人员
    Dim dsgnDate As Range   '设计日期
    Dim execCond As Range   '执行情况
    Dim tester As Range  '测试人员
    Dim supervisor As Range '监督人员
    Dim execDate As Range   '执行日期
    
    For tblNum = 1 To g_JL_Word.Tables.count
        Set tbl = g_JL_Word.Tables(tblNum)
        
        Set caseName = tbl.Rows(C_word_rowNum_caseName).Cells(2).Range
        Set caseCode = tbl.Rows(C_word_rowNum_caseCode).Cells(2).Range
        Set caseInit = tbl.Rows(C_word_rowNum_caseInit).Cells(2).Range
        Set premise = tbl.Rows(C_word_rowNum_premise).Cells(2).Range
        
        Set caseInput = tbl.Rows(C_word_rowNum_caseStart).Cells(2).Range
        Set caseExpectedOutcome = tbl.Rows(C_word_rowNum_caseStart).Cells(3).Range
        Set caseCert = tbl.Rows(C_word_rowNum_caseStart).Cells(4).Range
        Set caseActualResult = tbl.Rows(C_word_rowNum_caseStart).Cells(5).Range
        Set isPass = tbl.Rows(C_word_rowNum_caseStart).Cells(6).Range
        
        Set designer = tbl.Rows(C_word_rowNum_designer).Cells(2).Range
        Set dsgnDate = tbl.Rows(C_word_rowNum_designer).Cells(4).Range
        Set execCond = tbl.Rows(C_word_rowNum_designer).Cells(6).Range
        
        Set tester = tbl.Rows(C_word_rowNum_tester).Cells(2).Range
        Set supervisor = tbl.Rows(C_word_rowNum_tester).Cells(4).Range
        Set execDate = tbl.Rows(C_word_rowNum_tester).Cells(6).Range
        
        Call printDic3(g_dic_subItem_testCase)
        
        '如果测试用例标识一致("用例名称"))
        If arr_subItemCode(tblNum - 1) & "-" = Escape_Text(caseCode.Text) Then
        
            Dim dic_subCodeItem As Object: Set dic_subCodeItem = CreateObject("scripting.dictionary")
            Dim dic_testCase As Object: Set dic_testCase = CreateObject("scripting.dictionary")
            Dim caseNumInSubItem As Integer
            
            '获取测试子项下的所有测试用例字典
            Set dic_subCodeItem = g_dic_subItem_testCase(arr_subItemCode(tblNum - 1))
            '获取keys
            Dim arr_caseNumInSubItem() As Variant: arr_caseNumInSubItem = dic_subCodeItem.keys

            '遍历测试子项下的所有测试用例
            For caseNumInSubItem = 1 To dic_subCodeItem.count
                Set dic_testCase = dic_subCodeItem(arr_caseNumInSubItem(caseNumInSubItem - 1))
                Dim colName() As Variant: colName = dic_testCase.keys
                '如果测试用例的测试项编号与Word表格的编号一直
                If dic_testCase(colName(subItemCode - 1)) & "-" = Escape_Text(caseCode.Text) Then
                    Debug.Print dic_testCase(colName(subItemCode - 1))
                    '根据不同的列，填写这个用例到Word的表格中
                End If
            Next caseNumInSubItem
        End If
        
    
    Next tblNum
    
End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
'           5、saveJL：关闭和保存相关文件，另存word版的测试记录
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub saveJL()
    '1、另存word版的测试记录，并关闭原文件日不保存
    g_JL_Word.SaveAs (workPath & "\" & C_word_fileName_saved)
    g_JL_Word.Close wdDoNotSaveChanges
    '2、关闭exce1版的测试记录，且不保存
    g_JL_Excel.Close wdDoNotSaveChanges
    '3、退出创建的APP，释放APP内存
    'WordApp.Quit
    ExcelApp.Quit
    Set g_JL_Word = Nothing
    Set g_JL_Excel = Nothing
End Sub


'-------------------------------------------------------------------------------------------
'     1、过程getExcelLineInfo
'-------------------------------------------------------------------------------------------
Private Sub getExcelLineInfo()

    Dim sheetNum, rowCount, colCountOne, colCountAll, tmp As Long
      
    sheetNum = g_JL_Excel.worksheets.count
    ReDim g_arr_num_excel_rowInSheets(sheetNum)    '根据实际的sheet值重定义行数组
    ReDim g_arr_num_excel_colInSheets(sheetNum)     '根据实际的sheet值重定义列数组

    '获取每个sheet的行数、所有sheet的行数的总和
    For tmp = 1 To sheetNum
        g_arr_num_excel_rowInSheets(tmp) = g_JL_Excel.worksheets(tmp).usedrange.Rows.count '填充全局变量数组
        rowCount = rowCount + g_JL_Excel.worksheets(tmp).usedrange.Rows.count
    Next tmp
    
    g_num_excel_actual_row = rowCount '全局变量：获取测试记录所有sheet使用的行数
    
    
    '获取每个sheet的列数、所有sheet中列数的最大值
    colCountOne = g_JL_Excel.worksheets(1).usedrange.Columns.count
    colCountAll = g_JL_Excel.worksheets(1).usedrange.Columns.count
    For tmp = 1 To sheetNum
        '目前认为sheet中列数不能小于8（E_colNum_JLExcel.casePass），如果小于8，则置为8
        If g_JL_Excel.worksheets(tmp).usedrange.Columns.count < E_colNum_JLExcel.casePass Then
            colCountOne = E_colNum_JLExcel.casePass
        Else
            colCountOne = g_JL_Excel.worksheets(tmp).usedrange.Columns.count
        End If
        g_arr_num_excel_colInSheets(tmp) = colCountOne   '新增-获取每个sheet的列数
        
        If g_JL_Excel.worksheets(tmp).usedrange.Columns.count > colCountAll Then
            colCountAll = g_JL_Excel.worksheets(tmp).usedrange.Columns.count
        End If
    Next tmp
    
    g_num_excel_actual_column = colCountAll '全局变量：获取测试记录所有sheet的最大列数
    
End Sub
'-------------------------------------------------------------------------------------------
'     2、函数Escape_Text
'-------------------------------------------------------------------------------------------
Private Function Escape_Text(ByVal originStr As String)
    Dim tmp As String
    tmp = Replace(originStr, Chr(13), "")   '换行符
    tmp = Replace(tmp, Chr(10), "")         '换行符
    tmp = Replace(tmp, Chr(7), "")          'word表格中的小黑点
    tmp = Replace(tmp, Chr(9), "")          'Tab-水平制表符
    tmp = Replace(tmp, Chr(11), "")         '垂直制表符
    tmp = Replace(tmp, " ", "") '           即Chr(32) = 一个空格符
    Escape_Text = tmp
End Function
'-------------------------------------------------------------------------------------------
'     3、过程recErrInfo
'-------------------------------------------------------------------------------------------
Private Sub recErrInfo(ByVal errMsg As String, Optional ByVal sheetNum As String = "", Optional ByVal rowNum As String = "", Optional ByVal colNum As String = "")
    g_num_err_excel = g_num_err_excel + 1
    ReDim Preserve g_arr_err_excel(g_num_err_excel)
    If sheetNum <> "" Then
        g_arr_err_excel(g_num_err_excel) = g_num_err_excel & " --" & errMsg & "--位置：sheet=" & sheetNum & " 行=" & rowNum & " 列=" & colNum
    Else
        g_arr_err_excel(g_num_err_excel) = g_num_err_excel & " --" & errMsg
    End If
    
End Sub
'-------------------------------------------------------------------------------------------
'     4、函数getSubItemName：从用例名称中找到测试子项名称，比如从“初始化功能-01”中得到“初始化功能”
'-------------------------------------------------------------------------------------------
Private Function getSubItemName(ByVal str As String)
    
    Dim charPos As Integer
    charPos = InStrRev(str, "-")                '函数InStr()和函数InStrRev()，功能相同但方向相反，前者从左往右查找指定字符出现的位置，后者从右往左查找
    getSubItemName = Mid(str, 1, charPos - 1)
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'打印字典，printDic
'参数是引用传递，对于较大的字典可以节省时间
'调用的时候需要在函数前面加上 “Call”
'input：字典
'output：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function printDic1(ByRef dic As Object)
    'Debug.Print "printDic()"
    Dim count As Integer, keys, items
    
    keys = dic.keys
    items = dic.items
    
    Debug.Print "一维字典共有" & dic.count & "个元素"
    For count = 1 To dic.count
        Debug.Print "No." & count & ": " & "(" & keys(count - 1) & ", " & items(count - 1) & ")"
        'Debug.Print "第" & count & "个元素内容：" & items(count - 1) & Chr(13)
    Next count
    Debug.Print "打印完成。" & Chr(13)
End Function
Private Function printDic2(ByRef dic As Object)
    'Debug.Print "printDic()"
    Dim count As Integer, keys, items
    Dim dic2 As Object
    
    Set dic2 = CreateObject("scripting.dictionary")
    
    keys = dic.keys
    items = dic.items
    
    Debug.Print "二维字典共有" & dic.count & "个元素" & Chr(13)
    For count = 1 To dic.count
        Set dic2 = CreateObject("scripting.dictionary")
        Set dic2 = dic(keys(count - 1))
        Debug.Print "二维字典的第" & count & "个测试用例序号为：" & keys(count - 1)
        '第二层字典即dic2
        Call printDic1(dic2)
    Next count
    
    Debug.Print "打印完成。" & Chr(13)
End Function
Private Function printDic3(ByRef dic As Object)
    Dim count As Integer, keys, items
    Dim dic2 As Object
    
    Set dic2 = CreateObject("scripting.dictionary")
    
    keys = dic.keys
    items = dic.items
    
    Debug.Print "三维字典共有" & dic.count & "个元素" & Chr(13)
    For count = 1 To dic.count
        Set dic2 = CreateObject("scripting.dictionary")
        Set dic2 = dic(keys(count - 1))
        Debug.Print "三维字典的第" & count & "个测试子项标识号为：" & keys(count - 1) & Chr(13)
        '第二层字典即dic2
        Call printDic2(dic2)
    Next count
    
    Debug.Print "打印完成。"
End Function

Private Sub test()
    getSubItemName ("01-初始化-功能-01")
End Sub

    


