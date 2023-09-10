Attribute VB_Name = "ģ��1"
Sub ѵ��_���Ա�����()
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
'����Ԫ�����һ��, ProcessCellLastLine
'�����������е�Ԫ������һ�����ݣ����磺�����һ������(�հ��в���)�滻Ϊָ���ַ���
'���������ô��ݣ����ڽϴ�ı����Խ�ʡʱ��
'���õ�ʱ����Ҫ�ں���ǰ����� ��Call��
'input�����
'output����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ProcessCellLastLine(ByRef table As table)
    Dim findResult, textLen As Integer, textLenStart As Integer, textLenEnd As Integer
    Dim identifyStr As String, cellNum As Integer
    '���ж���������ǲ�������Ҫ�ҵı����ô�жϣ�һ�����ĳ���̶�λ�õ�Ԫ��������ǹ̶��ģ��Ǿ�ͨ�������ʶ��
    ''''''''''''''''''ÿ��ʹ�ö�Ҫ�����⼸������''''''''''''''''''''''''''''''''''''
    textLenStart = 5 '��Ϊ��������ʶ��ĵ�Ԫ�񣬸õ�Ԫ��ĳ���Ҫ��
    textLenEnd = 50
    identifyStr = "ά���Կ�" '��Ϊ��������ʶ��ĵ�Ԫ��������ַ���
    cellNum = 1 '�ñ��ĵڼ�����Ԫ�������ǹ̶��ġ��ܹ������ҵ����
    ''''''''''''''''''�������''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    textLen = Len(table.Range.Cells(cellNum).Range.Text)
    'InStr()�������һ���������ΪvbTextCompare��ʾ���Ƚϴ�Сд��vbBinaryCompare��Ƚϴ�Сд
    findResult = InStr(1, table.Range.Cells(cellNum).Range.Text, identifyStr, vbBinaryCompare)
    '�жϱ���ĳ����Ԫ������û���ض����ַ�,���Ҹõ�Ԫ����ַ�������Ҫ��ָ���ķ�Χ�ڣ���Ȼ��Ϊ�õ�Ԫ���ܱ�ʶ������
    If findResult > 0 And textLen >= textLenStart And textLen <= textLenEnd Then
'        Debug.Print "find the string."
        Dim cel As cell, celStart As Integer, celEnd As Integer, result As Boolean
        Set cel = table.Range.Cells(table.Range.Cells.count) '�ҵ��������һ����Ԫ��
        Call isLineEmpty(cel, result)
    Else
        Debug.Print "Not find."
    End If
    
    If table.Range.Cells(1).Range.Text = "ά���Կ����" Then '�жϱ���ĳ����Ԫ�Ƿ�Ϊ�ض����ַ���Ȼ����в���
        
    End If
    
    
    If table.Columns.count = 9 And table.Rows.count = 7 Then '�жϱ����С��еĳ���
    End If
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�����ĵ����壬setFileFont
'��ʱ�޷������ֺ�
'���桢��������Ϊ���壬��������Ϊ����
'���õ�ʱ����Ҫ�ں���ǰ����� ��Call��
'input���ĵ�����
'output����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function setFileFont(doc As Document)
    Dim chars As Characters, char As Range
    Set chars = doc.Characters
    For Each char In chars
        if char.Information(
    Next char

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ɾ����Ԫ����Ч���ݺ�Ŀհ���,deleteCellWhiteLine
'�����Ԫ������һ�������ǿհף��Ǿ�ɾ����һ�У�ֱ�����һ�в�Ϊ��
'������һ��ֻ��һ�����з�����ô��ֻɾ��������з�
'���������ô��ݣ����ڽϴ�ı����Խ�ʡʱ��
'���õ�ʱ����Ҫ�ں���ǰ����� ��Call��
'input����Ԫ��
'output����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function deleteCellWhiteLine(ByRef cel As cell)
    Dim num As Integer, startIndex As Integer, endIndex As Integer
    Dim area As Range

    Dim celText() As String
    celText = Split(cel.Range.Text, Chr(13)) '����Ԫ������ݰ����з����зָ�
    Dim lineNum As Integer: lineNum = UBound(celText) - LBound(celText) + 1 '��ȡ��Ԫ�����ݵ�����
    Dim i As Integer
    '������Ԫ��������
    For i = 1 To lineNum
        '��ȡ��Ԫ������һ���ַ�����ѡ�����������ַ��������ƶ���ע�͵ķ���(����area��õ�Ԫ������һ���ַ�)Ҳ�ǿ��Ե�
    '        Set area = cel.Range.Characters(cel.Range.Characters.count)
    '        area.Select
        ThisDocument.Range(cel.Range.End - 1, cel.Range.End).Select
        
        Selection.EndKey wdLine '�ƶ���굽���һ���ַ������У�Ҳ�������һ�У���ĩβ
        endIndex = Selection.End '��ȡ���һ�еĽ�βλ��
        Selection.HomeKey wdLine '�ƶ���굽���һ�е�����
        startIndex = Selection.End '��ȡ���һ�еĿ�ʼλ��
    '        Debug.Print endIndex & " --- " & startIndex
    
        '��β��λ����ͬ����ʾֻ�л��з���ɾ��������з�
        If startIndex = endIndex Then
            Selection.TypeBackspace
        '��βλ�ò�ͬ����ô���һ�а����ǿ��ַ����߿հ��ַ�����Ҫ��һ���ж�
        Else
            Dim tmpRng As Range: Set tmpRng = ThisDocument.Range(startIndex, endIndex) '��ȡ���һ�е���ֹ��
            '�ж���һ�е��������ǲ���ȫ�ǿհ��ַ����߻��з�
            If isStrBlank(tmpRng.Text) Then '����ַ����ﶼ�ǿհ��ַ��� �����߻��з�Chr(13)
                tmpRng.Delete wdCharacter 'ִ��delete
                Selection.TypeBackspace 'ɾ���������ݺ󣬱��л���ʣ��һ�����з���ɾ��������з�
            Else
            '�����һ�е����ݲ�Ϊ�գ��˳�ѭ��
                Exit For
            End If
        End If
    Next i
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�ж��ַ����Ƿ�ֻ�пհ��ַ��� �����߻��з�Chr(13),isStrBlank
'input���ַ���
'output���ַ���Ϊ�գ�True���ַ����ǿգ�False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function isStrBlank(str As String) As Boolean
    Dim arr() As String
    Dim isBlank As Boolean: isBlank = True
    
'    Dim str As String: str = "     " & Chr(13)
'    str = "  21  ��� " & Chr(13)

    '��һ���հ��ַ��� ���ָ��ַ���
    arr = Split(str, " ")
    
    Dim count As Integer: count = 0
    Do
'        Debug.Print "�ַ�" & count & ": " & Len(arr(count))
        '����ָ�֮����������в�Ϊ�յ��ַ����Ҹ��ַ����ǻ��з�������ü٣��˳�ѭ��
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
'�ж�ĳһRange���͵�Text�Ƿ�ֻ��" "��Chr(13),���������Ϊ��RangeΪ�գ�����True������г����������ַ���������ַ��ͷ���False
'���õ�ʱ����Ҫ�ں���ǰ����� ��Call��
'input��Range���ͣ�����set rng = Range(12, 78)��Ȼ����rng�������������ߴ���Range����
'output��ͨ�����ô��ݺ������������ڱ���result��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function isLineEmpty(ByRef rng As Range, ByRef result As Boolean)
    Dim isEmpty As Boolean
    isEmpty = True
    If Len(rng.Text) = 0 Then '������Ϊ1(�س���ռһ���ַ�)����RangeΪ��
        isEmpty = True
        
    ElseIf rng.Text = Chr(13) Then '��Range���ı������ǻ��з�����RangeΪ��
        isEmpty = True
        
    Else '����Range��ÿһ�У����ÿһ�ж�Ϊ�գ���RangeΪ��
        Dim char As Range, count As Integer
        For count = 1 To rng.Characters.count Step 1 '����һ���л��߲�������ĩ���Ǿ��ж�rng����������
            Set char = rng.Characters(count)
'            Debug.Print "�ı����ݣ�" & char.Text
            If char.Text <> " " And char.Text <> Chr(13) Then '���Range�����κ�һ���ַ����ǿհף��Ǿ��ж�Range�ǿ�
                isEmpty = False
                Exit For
            End If
        Next count
    End If
    result = isEmpty
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'������з�,delEnter
'input���ַ���
'output���޻��з����ַ���
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function delEnter(str As String) As String
    Debug.Print "delNull()"
    Dim tempStr As String
'    '�ж����޷Ǵ�ӡ�ַ�
'    If InStr(1, str, Chr(13), vbBinaryCompare) > 0 Then
'        'Debug.Print "find Chr(13)."
'        tempStr = Replace(str, Chr(13), "")
'    End If
    tempStr = Split(str, Chr(13))(0)
    delNull = tempStr
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ӡ����,printArr
'���������ô��ݣ����ڽϴ��������Խ�ʡʱ��
'input������
'output����
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
'��ӡ�ֵ䣬printDic
'���������ô��ݣ����ڽϴ���ֵ���Խ�ʡʱ��
'���õ�ʱ����Ҫ�ں���ǰ����� ��Call��
'input���ֵ�
'output����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function printDic(ByRef dic As Object)
    Debug.Print "printDic()"
    Dim count As Integer, keys, items
    keys = dic.keys
    items = dic.items
    count = 0
    Debug.Print "��ӡ�ֵ�..."
    Do
        Debug.Print "��" & count + 1 & "��: (" & keys(count) & ", " & items(count) & ")"
        count = count + 1
    Loop Until count >= dic.count
    Debug.Print "��ӡ��ɡ�"
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��Ԫ��������,showCellIndex
'չʾһ�����������е�Ԫ������кţ����������ô��ݣ����ڽϴ�ı����Խ�ʡʱ��
'���õ�ʱ����Ҫ�ں���ǰ����� ��Call��
'input�����
'output����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function showCellIndex(ByRef selectTable As Object)
    Debug.Print "showCellIndex()"
    Dim cel As cell
    For Each cel In selectTable.Range.Cells
        With cel
        '���������ֲ��뷽����ǰ����ֱ�����ӣ������Ǹ����µ����ݵ�ԭ�ĺ��棬���߿���ѡ��Ҫ��Ҫ���У�ǰ�߲���ѡ��
'            .Range.Text = .Range.Text & "(" & .RowIndex & "," & .ColumnIndex & ")"
            cel.Range.InsertAfter Chr(13) & "(" & .RowIndex & "," & .ColumnIndex & ")" '���ɾ��Chr(13)�ǾͲ������µ�һ�����
        End With
    Next
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ɾ����Ԫ�����һ��,delCellLastLine
'ɾ����������е�Ԫ������һ�����ݣ�������һ��ֻ��һ�����з�����ô��ֻɾ��������з�
'���������ô��ݣ����ڽϴ�ı����Խ�ʡʱ��
'���õ�ʱ����Ҫ�ں���ǰ����� ��Call��
'input�����
'output����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function delCellLastLine(ByRef tbl As Object)
    Dim cel As cell, num As Integer, startIndex As Integer, endIndex As Integer
    Dim area As Range
    For Each cel In tbl.Range.Cells
        '��ȡ��Ԫ������һ���ַ�����ѡ�����������ַ��������ƶ���ע�͵ķ���(����area��õ�Ԫ������һ���ַ�)Ҳ�ǿ��Ե�
'        Set area = cel.Range.Characters(cel.Range.Characters.count)
'        area.Select
        ThisDocument.Range(cel.Range.End - 1, cel.Range.End).Select
        
        Selection.EndKey wdLine '�ƶ���굽���һ���ַ������У�Ҳ�������һ�У���ĩβ
        
        endIndex = Selection.End '��ȡ���һ�еĽ�βλ��
       
        Selection.HomeKey wdLine '�ƶ���굽���һ�е�����
        
        startIndex = Selection.End '��ȡ���һ�еĿ�ʼλ��
'        Debug.Print endIndex & " --- " & startIndex
        If startIndex = endIndex Then
            
            Selection.TypeBackspace '��β��λ����ͬ����ʾֻ�л��з���ɾ��������з�
        Else
           
            ThisDocument.Range(startIndex, endIndex).Delete wdCharacter '��βλ�ò�ͬ��ִ��delete
            
            Selection.TypeBackspace 'ɾ���������ݺ󣬱��л���ʣ��һ�����з���ɾ��������з�
        End If
    Next
End Function

Function ����ƶ��͸�������()
    '�������һ���ַ�
    Selection.MoveEnd
    Selection.MoveRight Unit:=wdCharacter, count:=1
    '��������ƶ�һ���ַ�
    Selection.MoveLeft Unit:=wdCharacter, count:=2
    '��������ƶ�һ��(�������ƶ�һ��)
    Selection.MoveDown Unit:=wdLine, count:=1
    '��������ƶ�һ��(�������ƶ�һ��)
    Selection.MoveUp Unit:=wdLine, count:=1
    '����ƶ�����ĩ
    Selection.EndKey Unit:=wdLine
    '����ƶ�������
    Selection.HomeKey wdLine
    
    '����ƶ�����һ�����ĵ�һ����Ԫ��Ŀ�ʼλ��
    Selection.GoTo wdGoToTable, wdGoToFirst
    
    'ִ��һ��backspace����ɾ�����ǰ����ַ�
    Selection.TypeBackspace
    'ִ��һ��delete����ɾ����������ַ�
    Selection.Delete Unit:=wdCharacter, count:=1
    
    
    With ThisDocument.Tables(2)
        Dim count As Integer, scoreCell As cell, content As String
        '�����ֵ�洢��Ϣ
        Dim dic As Object, valArr, keyArr
        Set dic = CreateObject("scripting.dictionary")
        '�����������һ��
        For Each scoreCell In .Columns.Last.Cells
            '�ų�������һ�У���һ��û�з���
            If scoreCell.Row.Index <> 1 Then
                'ɾ������β�Ļ��з�
    '                content = Split(scoreCell.Range.Text, Chr(13))(0)
                content = delNull(scoreCell.Range.Text)
                '����˳������ֵ�
                dic.Add scoreCell.Row.Index, content
            End If
        Next scoreCell
    End With
    '    ��ӡ�ֵ�
    Call printDic(dic)
End Function
    
