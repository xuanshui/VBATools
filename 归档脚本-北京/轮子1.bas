Attribute VB_Name = "模块1"
Sub 训练_测试表格操作()
    Dim a As table
    For Each a In ThisDocument.Tables
'        Call showCellIndex(a)
'        Call delCellLastLine(a)
    Next a
    Call ProcessCellLastLine(ThisDocument.Tables(1))
'    Call isLineEmpty(ThisDocument.Range(1, 3))
    Dim doc As Document, rlt As Boolean, rng As Range
    Set doc = Application.ActiveDocument
    Set rng = doc.Range(0, 0)
    Debug.Print doc.Name
'    Debug.Print rng.Characters.count
'    Call isLineEmpty(rng, rlt)
'    If rlt Then
'        Debug.Print "Range is empty"
'    Else
'        Debug.Print "Range is not empty"
'    End If
'    Call showCellIndex(ThisDocument.Tables(1)) 'showCellIndex
'    Call delCellLastLine(ThisDocument.Tables(1)) 'delCellLastLine
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'处理单元格最后一行, ProcessCellLastLine
'处理表格里所有单元格的最后一行内容，例如：把最后一行内容(空白行不算)替换为指定字符串
'参数是引用传递，对于较大的表格可以节省时间
'调用的时候需要在函数前面加上 “Call”
'input：表格
'output：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ProcessCellLastLine(ByRef table As table)
    Dim findResult, textLen As Integer, textLenStart As Integer, textLenEnd As Integer
    Dim identifyStr As String, cellNum As Integer
    '先判断这个表格的是不是我们要找的表格，怎么判断？一般表格的某个固定位置单元格的内容是固定的，那就通过这个来识别
    ''''''''''''''''''每次使用都要设置这几个变量''''''''''''''''''''''''''''''''''''
    textLenStart = 5 '作为特征进行识别的单元格，该单元格的长度要求
    textLenEnd = 50
    identifyStr = "维修显控" '作为特征进行识别的单元格的特征字符串
    cellNum = 1 '该表格的第几个单元格内容是固定的、能够用来找到表格？
    ''''''''''''''''''设置完成''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    textLen = Len(table.Range.Cells(cellNum).Range.Text)
    'InStr()函数最后一个参数如果为vbTextCompare表示不比较大小写，vbBinaryCompare会比较大小写
    findResult = InStr(1, table.Range.Cells(cellNum).Range.Text, identifyStr, vbBinaryCompare)
    '判断表格的某个单元格里有没有特定的字符,并且该单元格的字符串长度要在指定的范围内，不然认为该单元格不能标识这个表格
    If findResult > 0 And textLen >= textLenStart And textLen <= textLenEnd Then
'        Debug.Print "find the string."
        Dim cel As cell, celStart As Integer, celEnd As Integer, result As Boolean
        Set cel = table.Range.Cells(table.Range.Cells.count) '找到表格的最后一个单元格
        Call isLineEmpty(cel, result)
    Else
        Debug.Print "Not find."
    End If
    
    If table.Range.Cells(1).Range.Text = "维修显控软件" Then '判断表格的某个单元是否为特定的字符，然后进行操作
        
    End If
    
    
    If table.Columns.count = 9 And table.Rows.count = 7 Then '判断表格的行、列的长度
    End If
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'设置文档字体，setFileFont
'暂时无法设置字号
'封面、标题字体为黑体，正文字体为宋体
'调用的时候需要在函数前面加上 “Call”
'input：文档对象
'output：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function setFileFont(doc As Document)
    Dim chars As Characters, char As Range
    Set chars = doc.Characters
    For Each char In chars
        if char.Information(
    Next char

End Function
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
    
'    Dim str As String: str = "     " & Chr(13)
'    str = "  21  你好 " & Chr(13)

    '用一个空白字符“ ”分割字符串
    arr = Split(str, " ")
    
    Dim count As Integer: count = 0
    Do
'        Debug.Print "字符" & count & ": " & Len(arr(count))
        '如果分割之后的数组里有不为空的字符，且该字符不是换行符，结果置假，退出循环
        If Len(arr(count)) <> 0 And arr(count) <> Chr(13) And arr(count) <> Chr(10) Then
            isBlank = False
'            Debug.Print "str is not blank."
            Exit Do
        End If
        count = count + 1
    Loop Until count >= (UBound(arr) - LBound(arr) + 1)
    isStrBlank = isBlank
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'isLineEmpty()
'判断某一Range类型的Text是否只有" "和Chr(13),如果是则认为该Range为空，返回True，如果有除了这两个字符外的其他字符就返回False
'调用的时候需要在函数前面加上 “Call”
'input：Range类型，可以set rng = Range(12, 78)，然后传入rng到本函数。或者传入Range对象
'output：通过引用传递函数结果，存放在变量result中
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function isLineEmpty(ByRef rng As Range, ByRef result As Boolean)
    Dim isEmpty As Boolean
    isEmpty = True
    If Len(rng.Text) = 0 Then '若长度为1(回车符占一个字符)，则Range为空
        isEmpty = True
        
    ElseIf rng.Text = Chr(13) Then '若Range的文本内容是换行符，则Range为空
        isEmpty = True
        
    Else '遍历Range的每一行，如果每一行都为空，则Range为空
        Dim char As Range, count As Integer
        For count = 1 To rng.Characters.count Step 1 '不是一整行或者不包括行末，那就判断rng的所有内容
            Set char = rng.Characters(count)
'            Debug.Print "文本内容：" & char.Text
            If char.Text <> " " And char.Text <> Chr(13) Then '如果Range内有任何一个字符不是空白，那就判断Range非空
                isEmpty = False
                Exit For
            End If
        Next count
    End If
    result = isEmpty
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'清除换行符,delEnter
'input：字符串
'output：无换行符的字符串
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function delEnter(str As String) As String
    Debug.Print "delNull()"
    Dim tempStr As String
'    '判断有无非打印字符
'    If InStr(1, str, Chr(13), vbBinaryCompare) > 0 Then
'        'Debug.Print "find Chr(13)."
'        tempStr = Replace(str, Chr(13), "")
'    End If
    tempStr = Split(str, Chr(13))(0)
    delNull = tempStr
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'打印数组,printArr
'参数是引用传递，对于较大的数组可以节省时间
'input：数组
'output：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function printArr(ByRef arr)
    Debug.Print "printArr()"
    Dim count As Integer
    Do
        Debug.Print arr(count)
        count = count + 1
    Loop Until count >= (UBound(arr) - LBound(arr) + 1)
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'打印字典，printDic
'参数是引用传递，对于较大的字典可以节省时间
'调用的时候需要在函数前面加上 “Call”
'input：字典
'output：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function printDic(ByRef dic As Object)
    Debug.Print "printDic()"
    Dim count As Integer, keys, items
    keys = dic.keys
    items = dic.items
    count = 0
    Debug.Print "打印字典..."
    Do
        Debug.Print "第" & count + 1 & "个: (" & keys(count) & ", " & items(count) & ")"
        count = count + 1
    Loop Until count >= dic.count
    Debug.Print "打印完成。"
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'单元格添加序号,showCellIndex
'展示一个表格里的所有单元格的行列号，参数是引用传递，对于较大的表格可以节省时间
'调用的时候需要在函数前面加上 “Call”
'input：表格
'output：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function showCellIndex(ByRef selectTable As Object)
    Debug.Print "showCellIndex()"
    Dim cel As cell
    For Each cel In selectTable.Range.Cells
        With cel
        '这里有两种插入方法，前者是直接连接，后者是附加新的内容到原文后面，后者可以选择要不要换行，前者不能选择
'            .Range.Text = .Range.Text & "(" & .RowIndex & "," & .ColumnIndex & ")"
            cel.Range.InsertAfter Chr(13) & "(" & .RowIndex & "," & .ColumnIndex & ")" '如果删除Chr(13)那就不会在新的一行添加
        End With
    Next
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'删除单元格最后一行,delCellLastLine
'删除表格里所有单元格的最后一行内容，如果最后一行只有一个换行符，那么就只删除这个换行符
'参数是引用传递，对于较大的表格可以节省时间
'调用的时候需要在函数前面加上 “Call”
'input：表格
'output：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function delCellLastLine(ByRef tbl As Object)
    Dim cel As cell, num As Integer, startIndex As Integer, endIndex As Integer
    Dim area As Range
    For Each cel In tbl.Range.Cells
        '获取单元格的最后一个字符，并选中它，有两种方法进行移动，注释的方法(利用area获得单元格的最后一个字符)也是可以的
'        Set area = cel.Range.Characters(cel.Range.Characters.count)
'        area.Select
        ThisDocument.Range(cel.Range.End - 1, cel.Range.End).Select
        
        Selection.EndKey wdLine '移动光标到最后一个字符所在行（也就是最后一行）的末尾
        
        endIndex = Selection.End '获取最后一行的结尾位置
       
        Selection.HomeKey wdLine '移动光标到最后一行的行首
        
        startIndex = Selection.End '获取最后一行的开始位置
'        Debug.Print endIndex & " --- " & startIndex
        If startIndex = endIndex Then
            
            Selection.TypeBackspace '首尾的位置相同，表示只有换行符，删除这个换行符
        Else
           
            ThisDocument.Range(startIndex, endIndex).Delete wdCharacter '首尾位置不同，执行delete
            
            Selection.TypeBackspace '删除该行数据后，本行还会剩余一个换行符，删除这个换行符
        End If
    Next
End Function

Function 光标移动和各种杂项()
    '光标右移一个字符
    Selection.MoveEnd
    Selection.MoveRight Unit:=wdCharacter, count:=1
    '光标向左移动一个字符
    Selection.MoveLeft Unit:=wdCharacter, count:=2
    '光标向下移动一次(即向下移动一行)
    Selection.MoveDown Unit:=wdLine, count:=1
    '光标向上移动一次(即向上移动一行)
    Selection.MoveUp Unit:=wdLine, count:=1
    '光标移动到行末
    Selection.EndKey Unit:=wdLine
    '光标移动到行首
    Selection.HomeKey wdLine
    
    '光标移动到第一个表格的第一个单元格的开始位置
    Selection.GoTo wdGoToTable, wdGoToFirst
    
    '执行一次backspace，即删除光标前面的字符
    Selection.TypeBackspace
    '执行一次delete，即删除光标后面的字符
    Selection.Delete Unit:=wdCharacter, count:=1
    
    
    With ThisDocument.Tables(2)
        Dim count As Integer, scoreCell As cell, content As String
        '创建字典存储信息
        Dim dic As Object, valArr, keyArr
        Set dic = CreateObject("scripting.dictionary")
        '遍历表格的最后一栏
        For Each scoreCell In .Columns.Last.Cells
            '排除掉表格第一行，这一行没有分数
            If scoreCell.Row.Index <> 1 Then
                '删除表格结尾的换行符
    '                content = Split(scoreCell.Range.Text, Chr(13))(0)
                content = delNull(scoreCell.Range.Text)
                '按照顺序放入字典
                dic.Add scoreCell.Row.Index, content
            End If
        Next scoreCell
    End With
    '    打印字典
    Call printDic(dic)
End Function
    
