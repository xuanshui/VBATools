Attribute VB_Name = "模块3"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'统计文件夹下的Word文档的页数，在新的文件"file_pages.doc"里以表格形式展现
'完成度：100%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub 统计文件夹下所有文档的页数()
    '设置参数
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '需要读取信息的文档的路径
    Dim workPath As String: workPath = ThisDocument.Path & "\"
    '保存信息的文件名
    Dim infoFileName As String: infoFileName = "file_pages.doc"
    '如果保存信息的文件不存在，新建后保存的路径
    Dim wholeInfoFileName As String: wholeInfoFileName = workPath & infoFileName
    'Word文档的后缀类型
    Dim fileType As String: fileType = "*.docx"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '判断目录下是否存在文件infoFileName，不存在则新建，存在则打开
    Dim FSO As Object: Set FSO = CreateObject("scripting.fileSystemObject")
    '如果文件已存在，打开
    If FSO.FileExists(wholeInfoFileName) Then
        Documents.Open wholeInfoFileName
    '文件不存在，新建
    Else
        Dim newDoc As Document
        Set newDoc = Documents.Add '或者使用模板创建文件Set newDoc = Documents.Add template.docx
        '将文档保存在指定目录下
'        ChangeFileOpenDirectory ThisDocument.Path  '或者指定目录 ChangeFileOpenDirectory "D:\Files\"
        ActiveDocument.SaveAs2 wholeInfoFileName
    End If
    
    '在文末插入表格，记录信息
    Selection.EndKey wdStory '光标跳至文末
    Selection.Text = Chr(13) & workPath & "-文件夹信息: " & Chr(13)
    Selection.EndKey wdStory '光标跳至文末
    '插入1行3列的表格
    ActiveDocument.Tables.Add Selection.Range, 1, 3
    With ActiveDocument.Tables(ActiveDocument.Tables.count).Rows.First
        .Cells(1).Range.Text = "文件名"
        .Cells(2).Range.Text = "页数"
        .Cells(3).Range.Text = "备注"
    End With
    
    '获取文件所在目录下的所有doc和docx文件
    Dim lastRow As Row
    Dim fileName As String, docReading As Document
    fileName = Dir(workPath & fileType)
    Do
        Set docReading = Documents.Open(workPath & fileName) '打开需要读取信息的文件，不要遗漏字符“\”
        
        '向infoFileName文件中写入信息
        Documents(infoFileName).Activate
        ActiveDocument.Tables(ActiveDocument.Tables.count).Rows.Last.Select '选中最后一个表格的最后一行
        Selection.InsertRowsBelow '在后面插入新的一行
        Set lastRow = ActiveDocument.Tables(ActiveDocument.Tables.count).Rows.Last
        '向新插入行的第一栏写数据
        lastRow.Range.Cells(1).Range.Text = docReading.Name '文件名
        lastRow.Range.Cells(2).Range.Text = docReading.ComputeStatistics(wdStatisticPages) '文件页数
        lastRow.Range.Cells(3).Range.Text = "" '备注栏留空
        '保存
        ActiveDocument.Save
        '关闭读取页数的文件，不保存
        docReading.Close 0
        
        '读取workPath目录里的下一个docx文件
        fileName = Dir
    'Dir函数会重复读取上一次Dir函数的目录参数下的所有指定文件，直到读取完毕，返回空字符串""
    Loop Until fileName = ""
    Windows(infoFileName).Activate
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'填写问题报告的问题表的签署和日期
'完成度：100%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub 问题报告_填写签名和日期()
    '设置参数
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim testTextStr As String: testTextStr = "签字：李雅莹  日期：20201202" '新的测试人员签署
    Dim dsgnTextStr As String: dsgnTextStr = "签字：吴金平  日期：20201202" '新的开发人员签署
    Dim doc As Document: Set doc = ActiveDocument '默认问题报告文档是当前的活跃文档
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '同时满足如下三个条件，则认为该表是问题表，然后才进行修改
    ''条件1、表格的第一个单元格是这个内容，就认为该表格是问题表，但是并不一定会修改
    ''条件2、表格最后一个单元格或者倒数第三个单元格的最后一行包含关键字“签字”
    ''条件3、表格最后一个单元格或者倒数第三个单元格的最后一行包含关键字“日期”
    Dim idStr As String: idStr = "报告单编号" '条件1
    Dim idStr1 As String: idStr1 = "签字" '条件2
    Dim idStr2 As String: idStr2 = "日期" '条件3
    
    Dim tblCount As Integer
    Dim tbl As table, testCel As cell, dsgnCel As cell
    Dim testArr() As String, dsgnArr() As String '存放要操作的单元格文本
    Dim textLen As Integer
    
    For Each tbl In doc.Tables
        '判断表格是不是问题表
        If tbl.Range.Cells.count > 20 And InStr(1, tbl.Range.Cells(1).Range.Text, idStr) And tbl.Range.Cells(1).Range.Characters.count < 7 Then
            Set testCel = tbl.Range.Cells(tbl.Range.Cells.count - 2)
            Set dsgnCel = tbl.Range.Cells(tbl.Range.Cells.count)
            
            ''''''''''''''''''''
            '1、处理测试签字
            ''''''''''''''''''''
            '删除有效内容后面多余的空白行
            Call deleteCellWhiteLine(testCel)
            testArr = Split(testCel.Range.Text, Chr(13))
            textStr = testArr(UBound(testArr) - 1)
            '如果文本内容里既有“签名”又有“日期”，就认为找到了正确的字符串
            If InStr(1, textStr, idStr1) > 0 And InStr(1, textStr, idStr2) > 0 Then
                
'                testArr(UBound(testArr) - 1) = textStr
'                testCel.Range.Text = Join(testArr, Chr(13))
                textLen = Len(textStr)
                
                '获取单元格的最后一个字符，并选中它，有两种方法进行移动，注释的方法(利用area获得单元格的最后一个字符)也是可以的
                ThisDocument.Range(testCel.Range.End - 1, testCel.Range.End).Select
                Selection.EndKey wdLine  '移动光标到最后一个字符所在行（也就是最后一行）的末尾
                endIndex = Selection.End '获取最后一行的结尾位置
                ThisDocument.Range(endIndex - textLen, endIndex).Delete wdCharacter '首尾位置不同，执行delete
                
                Selection.Text = testTextStr
                Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '将所在行的格式设置为靠右对齐
            End If
            
            ''''''''''''''''''''
            '2、处理开发签字
            ''''''''''''''''''''
            '删除有效内容后面多余的空白行
            Call deleteCellWhiteLine(dsgnCel)
            dsgnArr = Split(dsgnCel.Range.Text, Chr(13))
            textStr = dsgnArr(UBound(dsgnArr) - 1) '得到最后一行的文本内容
            '如果文本内容里既有“签名”又有“日期”，就认为找到了正确的字符串
            If InStr(1, textStr, idStr1) > 0 And InStr(1, textStr, idStr2) > 0 Then
                '获取最后一行的字符长度
                textLen = Len(textStr)
                
                '获取单元格的最后一个字符，并选中它，有两种方法进行移动，注释的方法(利用area获得单元格的最后一个字符)也是可以的
                ThisDocument.Range(dsgnCel.Range.End - 1, dsgnCel.Range.End).Select
                Selection.EndKey wdLine  '移动光标到最后一个字符所在行（也就是最后一行）的末尾
                endIndex = Selection.End '获取最后一行的结尾位置
                ThisDocument.Range(endIndex - textLen, endIndex).Delete wdCharacter '执行delete
                
                Selection.Text = dsgnTextStr
                Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '将所在行的格式设置为靠右对齐
            End If
            
            tblCount = tblCount + 1
        End If
    Next tbl
    Debug.Print "问题报告脚本: 己填充/修改 " & tblCount & " 个表格"
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'测试函数deleteCellWhiteLine能否正常工作
'结论：功能正常
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub 测试函数deleteCellWhiteLine能否正常工作()
    Dim cel As cell
    Set cel = ActiveDocument.Tables(1).cell(8, 2)
    Call deleteCellWhiteLine(cel)
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'删除单元格有效内容后的空白行,deleteCellWhiteLine
'如果单元格的最后一行内容是空白，那就删除这一行，直到最后一行不为空
'如果最后一行只有一个换行符，那么就只删除这个换行符
'参数是引用传递，对于较大的表格可以节省时间
'调用的时候需要在函数前面加上 “Call”
'input：单元格
'output：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function deleteCellWhiteLine(ByRef cel As cell)
    Dim num As Integer, startIndex As Integer, endIndex As Integer
    Dim area As Range

    Dim celText() As String
    celText = Split(cel.Range.Text, Chr(13)) '将单元格的内容按换行符进行分割
    Dim lineNum As Integer: lineNum = UBound(celText) - LBound(celText) + 1 '获取单元格内容的行数
    Dim i As Integer
    '遍历单元格所有行
    For i = 1 To lineNum
        '获取单元格的最后一个字符，并选中它，有两种方法进行移动，注释的方法(利用area获得单元格的最后一个字符)也是可以的
    '        Set area = cel.Range.Characters(cel.Range.Characters.count)
    '        area.Select
        ThisDocument.Range(cel.Range.End - 1, cel.Range.End).Select
        
        Selection.EndKey wdLine '移动光标到最后一个字符所在行（也就是最后一行）的末尾
        endIndex = Selection.End '获取最后一行的结尾位置
        Selection.HomeKey wdLine '移动光标到最后一行的行首
        startIndex = Selection.End '获取最后一行的开始位置
    '        Debug.Print endIndex & " --- " & startIndex
    
        '首尾的位置相同，表示只有换行符，删除这个换行符
        If startIndex = endIndex Then
            Selection.TypeBackspace
        '首尾位置不同，那么最后一行包含非空字符或者空白字符，需要进一步判断
        Else
            Dim tmpRng As Range: Set tmpRng = ThisDocument.Range(startIndex, endIndex) '获取最后一行的起止点
            '判断这一行的内容里是不是全是空白字符或者换行符
            If isStrBlank(tmpRng.Text) Then '如果字符串里都是空白字符“ ”或者换行符Chr(13)
                tmpRng.Delete wdCharacter '执行delete
                Selection.TypeBackspace '删除该行数据后，本行还会剩余一个换行符，删除这个换行符
            Else
            '如果这一行的内容不为空，退出循环
                Exit For
            End If
        End If
    Next i
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'判断字符串是否只有空白字符“ ”或者换行符Chr(13),isStrBlank
'input：字符串
'output：字符串为空：True，字符串非空：False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function isStrBlank(str As String) As Boolean
    Dim arr() As String
    Dim isBlank As Boolean: isBlank = True
    '用一个空白字符“ ”分割字符串
    arr = Split(str, " ")
    Dim count As Integer: count = 0
    Do
        '如果分割之后的数组里有不为空的字符，且该字符不是换行符，结果置假，退出循环
        If Len(arr(count)) <> 0 And arr(count) <> Chr(13) And arr(count) <> Chr(10) Then
            isBlank = False
            Exit Do
        End If
        count = count + 1
    Loop Until count >= (UBound(arr) - LBound(arr) + 1)
    isStrBlank = isBlank
End Function

Sub 测试说明_填写人员和日期()
    Dim tblCount As Integer: tblCount = 0
    Dim idStr As String: idStr = "执行日期"
    Dim tbl As table, 设计人员 As Range, 设计日期 As Range, 执行情况 As Range, 测试人员 As Range, 监督人员 As Range, 执行日期 As Range
        
    For Each tbl In ActiveDocument.Tables
        If tbl.Range.Cells.count > 20 And InStr(1, tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Text, idStr) And tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Characters.count < 6 Then
            
            tblCount = tblCount + 1
            
            Set 设计人员 = tbl.Range.Cells(tbl.Range.Cells.count - 10).Range
            Set 设计日期 = tbl.Range.Cells(tbl.Range.Cells.count - 8).Range
            Set 执行情况 = tbl.Range.Cells(tbl.Range.Cells.count - 6).Range
            Set 测试人员 = tbl.Range.Cells(tbl.Range.Cells.count - 4).Range
            Set 监督人员 = tbl.Range.Cells(tbl.Range.Cells.count - 2).Range
            Set 执行日期 = tbl.Range.Cells(tbl.Range.Cells.count).Range
            
            设计人员.Text = "杨成勇"
            设计日期.Text = "20201010"
            执行情况.Text = ""
            测试人员.Text = ""
            监督人员.Text = ""
            执行日期.Text = ""
            
            Debug.Print tblCount & ":" & 设计人员.Text
        End If
    Next tbl
    Debug.Print "测试说明脚本: 已填充/修改 " & tblCount & " 个表格"
End Sub

Sub 测试记录_填写人员和日期()
    Dim tblCount As Integer: tblCount = 0
    Dim idStr As String: idStr = "执行日期"
    Dim tbl As table, 设计人员 As Range, 设计日期 As Range, 执行情况 As Range, 测试人员 As Range, 监督人员 As Range, 执行日期 As Range
    
    '遍历文档所有表格
    For Each tbl In ActiveDocument.Tables
        '1-185页是第一轮测试
        If tbl.Range.Information(wdActiveEndPageNumber) <= 18 Then
        '表格的单元格数量大于20,表格的倒数第二个单元格包含字符串“执行日期”，表格的倒数第二个单元格的字符个数小于6
            If tbl.Range.Cells.count > 20 And InStr(1, tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Text, idStr) And tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Characters.count < 6 Then
                
                tblCount = tblCount + 1
                
                Set 设计人员 = tbl.Range.Cells(tbl.Range.Cells.count - 10).Range
                Set 设计日期 = tbl.Range.Cells(tbl.Range.Cells.count - 8).Range
                Set 执行情况 = tbl.Range.Cells(tbl.Range.Cells.count - 6).Range
                Set 测试人员 = tbl.Range.Cells(tbl.Range.Cells.count - 4).Range
                Set 监督人员 = tbl.Range.Cells(tbl.Range.Cells.count - 2).Range
                Set 执行日期 = tbl.Range.Cells(tbl.Range.Cells.count).Range
                
                设计人员.Text = "杨成勇"
                设计日期.Text = "20201010"
                执行情况.Text = "已执行"
                测试人员.Text = "李雅莹"
                监督人员.Text = "杨豹"
                执行日期.Text = "20201015"
                
                Debug.Print tblCount & ":" & 设计人员.Text
            End If
        '185页之后是回归测试
        ElseIf tbl.Range.Information(wdActiveEndPageNumber) > 185 Then
            If tbl.Range.Cells.count > 20 And InStr(1, tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Text, idStr) And tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Characters.count < 6 Then
                
                tblCount = tblCount + 1
                
                Set 设计人员 = tbl.Range.Cells(tbl.Range.Cells.count - 10).Range
                Set 设计日期 = tbl.Range.Cells(tbl.Range.Cells.count - 8).Range
                Set 执行情况 = tbl.Range.Cells(tbl.Range.Cells.count - 6).Range
                Set 测试人员 = tbl.Range.Cells(tbl.Range.Cells.count - 4).Range
                Set 监督人员 = tbl.Range.Cells(tbl.Range.Cells.count - 2).Range
                Set 执行日期 = tbl.Range.Cells(tbl.Range.Cells.count).Range
                
                设计人员.Text = "杨成勇"
                设计日期.Text = "20201010"
                执行情况.Text = "已执行"
                测试人员.Text = "李雅莹"
                监督人员.Text = "杨豹"
                执行日期.Text = "20201115"
                
                Debug.Print tblCount & ":" & 设计人员.Text
            End If
        End If
    Next tbl
    Debug.Print "测试记录脚本: 已填充/修改 " & tblCount & " 个表格"
End Sub
