Attribute VB_Name = "ģ��3"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ͳ���ļ����µ�Word�ĵ���ҳ�������µ��ļ�"file_pages.doc"���Ա����ʽչ��
'��ɶȣ�100%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ͳ���ļ����������ĵ���ҳ��()
    '���ò���
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '��Ҫ��ȡ��Ϣ���ĵ���·��
    Dim workPath As String: workPath = ThisDocument.Path & "\"
    '������Ϣ���ļ���
    Dim infoFileName As String: infoFileName = "file_pages.doc"
    '���������Ϣ���ļ������ڣ��½��󱣴��·��
    Dim wholeInfoFileName As String: wholeInfoFileName = workPath & infoFileName
    'Word�ĵ��ĺ�׺����
    Dim fileType As String: fileType = "*.docx"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '�ж�Ŀ¼���Ƿ�����ļ�infoFileName�����������½����������
    Dim FSO As Object: Set FSO = CreateObject("scripting.fileSystemObject")
    '����ļ��Ѵ��ڣ���
    If FSO.FileExists(wholeInfoFileName) Then
        Documents.Open wholeInfoFileName
    '�ļ������ڣ��½�
    Else
        Dim newDoc As Document
        Set newDoc = Documents.Add '����ʹ��ģ�崴���ļ�Set newDoc = Documents.Add template.docx
        '���ĵ�������ָ��Ŀ¼��
'        ChangeFileOpenDirectory ThisDocument.Path  '����ָ��Ŀ¼ ChangeFileOpenDirectory "D:\Files\"
        ActiveDocument.SaveAs2 wholeInfoFileName
    End If
    
    '����ĩ�����񣬼�¼��Ϣ
    Selection.EndKey wdStory '���������ĩ
    Selection.Text = Chr(13) & workPath & "-�ļ�����Ϣ: " & Chr(13)
    Selection.EndKey wdStory '���������ĩ
    '����1��3�еı��
    ActiveDocument.Tables.Add Selection.Range, 1, 3
    With ActiveDocument.Tables(ActiveDocument.Tables.count).Rows.First
        .Cells(1).Range.Text = "�ļ���"
        .Cells(2).Range.Text = "ҳ��"
        .Cells(3).Range.Text = "��ע"
    End With
    
    '��ȡ�ļ�����Ŀ¼�µ�����doc��docx�ļ�
    Dim lastRow As Row
    Dim fileName As String, docReading As Document
    fileName = Dir(workPath & fileType)
    Do
        Set docReading = Documents.Open(workPath & fileName) '����Ҫ��ȡ��Ϣ���ļ�����Ҫ��©�ַ���\��
        
        '��infoFileName�ļ���д����Ϣ
        Documents(infoFileName).Activate
        ActiveDocument.Tables(ActiveDocument.Tables.count).Rows.Last.Select 'ѡ�����һ���������һ��
        Selection.InsertRowsBelow '�ں�������µ�һ��
        Set lastRow = ActiveDocument.Tables(ActiveDocument.Tables.count).Rows.Last
        '���²����еĵ�һ��д����
        lastRow.Range.Cells(1).Range.Text = docReading.Name '�ļ���
        lastRow.Range.Cells(2).Range.Text = docReading.ComputeStatistics(wdStatisticPages) '�ļ�ҳ��
        lastRow.Range.Cells(3).Range.Text = "" '��ע������
        '����
        ActiveDocument.Save
        '�رն�ȡҳ�����ļ���������
        docReading.Close 0
        
        '��ȡworkPathĿ¼�����һ��docx�ļ�
        fileName = Dir
    'Dir�������ظ���ȡ��һ��Dir������Ŀ¼�����µ�����ָ���ļ���ֱ����ȡ��ϣ����ؿ��ַ���""
    Loop Until fileName = ""
    Windows(infoFileName).Activate
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��д���ⱨ���������ǩ�������
'��ɶȣ�100%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ���ⱨ��_��дǩ��������()
    '���ò���
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim testTextStr As String: testTextStr = "ǩ�֣�����Ө  ���ڣ�20201202" '�µĲ�����Աǩ��
    Dim dsgnTextStr As String: dsgnTextStr = "ǩ�֣����ƽ  ���ڣ�20201202" '�µĿ�����Աǩ��
    Dim doc As Document: Set doc = ActiveDocument 'Ĭ�����ⱨ���ĵ��ǵ�ǰ�Ļ�Ծ�ĵ�
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ͬʱ����������������������Ϊ�ñ��������Ȼ��Ž����޸�
    ''����1�����ĵ�һ����Ԫ����������ݣ�����Ϊ�ñ������������ǲ���һ�����޸�
    ''����2��������һ����Ԫ����ߵ�����������Ԫ������һ�а����ؼ��֡�ǩ�֡�
    ''����3��������һ����Ԫ����ߵ�����������Ԫ������һ�а����ؼ��֡����ڡ�
    Dim idStr As String: idStr = "���浥���" '����1
    Dim idStr1 As String: idStr1 = "ǩ��" '����2
    Dim idStr2 As String: idStr2 = "����" '����3
    
    Dim tblCount As Integer
    Dim tbl As table, testCel As cell, dsgnCel As cell
    Dim testArr() As String, dsgnArr() As String '���Ҫ�����ĵ�Ԫ���ı�
    Dim textLen As Integer
    
    For Each tbl In doc.Tables
        '�жϱ���ǲ��������
        If tbl.Range.Cells.count > 20 And InStr(1, tbl.Range.Cells(1).Range.Text, idStr) And tbl.Range.Cells(1).Range.Characters.count < 7 Then
            Set testCel = tbl.Range.Cells(tbl.Range.Cells.count - 2)
            Set dsgnCel = tbl.Range.Cells(tbl.Range.Cells.count)
            
            ''''''''''''''''''''
            '1���������ǩ��
            ''''''''''''''''''''
            'ɾ����Ч���ݺ������Ŀհ���
            Call deleteCellWhiteLine(testCel)
            testArr = Split(testCel.Range.Text, Chr(13))
            textStr = testArr(UBound(testArr) - 1)
            '����ı���������С�ǩ�������С����ڡ�������Ϊ�ҵ�����ȷ���ַ���
            If InStr(1, textStr, idStr1) > 0 And InStr(1, textStr, idStr2) > 0 Then
                
'                testArr(UBound(testArr) - 1) = textStr
'                testCel.Range.Text = Join(testArr, Chr(13))
                textLen = Len(textStr)
                
                '��ȡ��Ԫ������һ���ַ�����ѡ�����������ַ��������ƶ���ע�͵ķ���(����area��õ�Ԫ������һ���ַ�)Ҳ�ǿ��Ե�
                ThisDocument.Range(testCel.Range.End - 1, testCel.Range.End).Select
                Selection.EndKey wdLine  '�ƶ���굽���һ���ַ������У�Ҳ�������һ�У���ĩβ
                endIndex = Selection.End '��ȡ���һ�еĽ�βλ��
                ThisDocument.Range(endIndex - textLen, endIndex).Delete wdCharacter '��βλ�ò�ͬ��ִ��delete
                
                Selection.Text = testTextStr
                Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '�������еĸ�ʽ����Ϊ���Ҷ���
            End If
            
            ''''''''''''''''''''
            '2��������ǩ��
            ''''''''''''''''''''
            'ɾ����Ч���ݺ������Ŀհ���
            Call deleteCellWhiteLine(dsgnCel)
            dsgnArr = Split(dsgnCel.Range.Text, Chr(13))
            textStr = dsgnArr(UBound(dsgnArr) - 1) '�õ����һ�е��ı�����
            '����ı���������С�ǩ�������С����ڡ�������Ϊ�ҵ�����ȷ���ַ���
            If InStr(1, textStr, idStr1) > 0 And InStr(1, textStr, idStr2) > 0 Then
                '��ȡ���һ�е��ַ�����
                textLen = Len(textStr)
                
                '��ȡ��Ԫ������һ���ַ�����ѡ�����������ַ��������ƶ���ע�͵ķ���(����area��õ�Ԫ������һ���ַ�)Ҳ�ǿ��Ե�
                ThisDocument.Range(dsgnCel.Range.End - 1, dsgnCel.Range.End).Select
                Selection.EndKey wdLine  '�ƶ���굽���һ���ַ������У�Ҳ�������һ�У���ĩβ
                endIndex = Selection.End '��ȡ���һ�еĽ�βλ��
                ThisDocument.Range(endIndex - textLen, endIndex).Delete wdCharacter 'ִ��delete
                
                Selection.Text = dsgnTextStr
                Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '�������еĸ�ʽ����Ϊ���Ҷ���
            End If
            
            tblCount = tblCount + 1
        End If
    Next tbl
    Debug.Print "���ⱨ��ű�: �����/�޸� " & tblCount & " �����"
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���Ժ���deleteCellWhiteLine�ܷ���������
'���ۣ���������
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ���Ժ���deleteCellWhiteLine�ܷ���������()
    Dim cel As cell
    Set cel = ActiveDocument.Tables(1).cell(8, 2)
    Call deleteCellWhiteLine(cel)
End Sub
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
    '��һ���հ��ַ��� ���ָ��ַ���
    arr = Split(str, " ")
    Dim count As Integer: count = 0
    Do
        '����ָ�֮����������в�Ϊ�յ��ַ����Ҹ��ַ����ǻ��з�������ü٣��˳�ѭ��
        If Len(arr(count)) <> 0 And arr(count) <> Chr(13) And arr(count) <> Chr(10) Then
            isBlank = False
            Exit Do
        End If
        count = count + 1
    Loop Until count >= (UBound(arr) - LBound(arr) + 1)
    isStrBlank = isBlank
End Function

Sub ����˵��_��д��Ա������()
    Dim tblCount As Integer: tblCount = 0
    Dim idStr As String: idStr = "ִ������"
    Dim tbl As table, �����Ա As Range, ������� As Range, ִ����� As Range, ������Ա As Range, �ල��Ա As Range, ִ������ As Range
        
    For Each tbl In ActiveDocument.Tables
        If tbl.Range.Cells.count > 20 And InStr(1, tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Text, idStr) And tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Characters.count < 6 Then
            
            tblCount = tblCount + 1
            
            Set �����Ա = tbl.Range.Cells(tbl.Range.Cells.count - 10).Range
            Set ������� = tbl.Range.Cells(tbl.Range.Cells.count - 8).Range
            Set ִ����� = tbl.Range.Cells(tbl.Range.Cells.count - 6).Range
            Set ������Ա = tbl.Range.Cells(tbl.Range.Cells.count - 4).Range
            Set �ල��Ա = tbl.Range.Cells(tbl.Range.Cells.count - 2).Range
            Set ִ������ = tbl.Range.Cells(tbl.Range.Cells.count).Range
            
            �����Ա.Text = "�����"
            �������.Text = "20201010"
            ִ�����.Text = ""
            ������Ա.Text = ""
            �ල��Ա.Text = ""
            ִ������.Text = ""
            
            Debug.Print tblCount & ":" & �����Ա.Text
        End If
    Next tbl
    Debug.Print "����˵���ű�: �����/�޸� " & tblCount & " �����"
End Sub

Sub ���Լ�¼_��д��Ա������()
    Dim tblCount As Integer: tblCount = 0
    Dim idStr As String: idStr = "ִ������"
    Dim tbl As table, �����Ա As Range, ������� As Range, ִ����� As Range, ������Ա As Range, �ල��Ա As Range, ִ������ As Range
    
    '�����ĵ����б��
    For Each tbl In ActiveDocument.Tables
        '1-185ҳ�ǵ�һ�ֲ���
        If tbl.Range.Information(wdActiveEndPageNumber) <= 18 Then
        '���ĵ�Ԫ����������20,���ĵ����ڶ�����Ԫ������ַ�����ִ�����ڡ������ĵ����ڶ�����Ԫ����ַ�����С��6
            If tbl.Range.Cells.count > 20 And InStr(1, tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Text, idStr) And tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Characters.count < 6 Then
                
                tblCount = tblCount + 1
                
                Set �����Ա = tbl.Range.Cells(tbl.Range.Cells.count - 10).Range
                Set ������� = tbl.Range.Cells(tbl.Range.Cells.count - 8).Range
                Set ִ����� = tbl.Range.Cells(tbl.Range.Cells.count - 6).Range
                Set ������Ա = tbl.Range.Cells(tbl.Range.Cells.count - 4).Range
                Set �ල��Ա = tbl.Range.Cells(tbl.Range.Cells.count - 2).Range
                Set ִ������ = tbl.Range.Cells(tbl.Range.Cells.count).Range
                
                �����Ա.Text = "�����"
                �������.Text = "20201010"
                ִ�����.Text = "��ִ��"
                ������Ա.Text = "����Ө"
                �ල��Ա.Text = "�"
                ִ������.Text = "20201015"
                
                Debug.Print tblCount & ":" & �����Ա.Text
            End If
        '185ҳ֮���ǻع����
        ElseIf tbl.Range.Information(wdActiveEndPageNumber) > 185 Then
            If tbl.Range.Cells.count > 20 And InStr(1, tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Text, idStr) And tbl.Range.Cells(tbl.Range.Cells.count - 1).Range.Characters.count < 6 Then
                
                tblCount = tblCount + 1
                
                Set �����Ա = tbl.Range.Cells(tbl.Range.Cells.count - 10).Range
                Set ������� = tbl.Range.Cells(tbl.Range.Cells.count - 8).Range
                Set ִ����� = tbl.Range.Cells(tbl.Range.Cells.count - 6).Range
                Set ������Ա = tbl.Range.Cells(tbl.Range.Cells.count - 4).Range
                Set �ල��Ա = tbl.Range.Cells(tbl.Range.Cells.count - 2).Range
                Set ִ������ = tbl.Range.Cells(tbl.Range.Cells.count).Range
                
                �����Ա.Text = "�����"
                �������.Text = "20201010"
                ִ�����.Text = "��ִ��"
                ������Ա.Text = "����Ө"
                �ල��Ա.Text = "�"
                ִ������.Text = "20201115"
                
                Debug.Print tblCount & ":" & �����Ա.Text
            End If
        End If
    Next tbl
    Debug.Print "���Լ�¼�ű�: �����/�޸� " & tblCount & " �����"
End Sub
