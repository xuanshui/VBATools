Sub ͳ���ļ�ҳ��_���հ�()

    Dim workPath As String: workPath = ThisDocument.Path & "\"
    Dim fileType As String: fileType = "*.doc?"
    Dim infoFileName As String: infoFileName = "ҳ��ͳ��.doc"
    Dim wholeInfoFileName As String: wholeInfoFileName = ThisDocument.Path & "\" & infoFileName
    
    Dim startTime: startTime = Now
    
    '1����ȡ����Ŀ¼�µĵ�һ����Ŀ¼���ƣ���Ž��ֵ���
    Dim FSO As Object: Set FSO = CreateObject("scripting.filesystemobject")
    Dim rootFolder As Object: Set rootFolder = FSO.GetFolder(workPath)
    Dim subFolders As Object: Set subFolders = rootFolder.subFolders
    Dim tmpFolder
    '���ļ�����Ϣ
    Dim foldersDic As Object: Set foldersDic = CreateObject("scripting.dictionary")
    For Each tmpFolder In subFolders
        foldersDic.Add tmpFolder.Name, tmpFolder.Path
    Next tmpFolder

    
    '2���жϱ�����Ϣ���ļ�infoDoc�Ƿ����
    Dim infoDoc As Document
    If FSO.FileExists(wholeInfoFileName) Then '���ھʹ�
        Set infoDoc = Documents.Open(wholeInfoFileName)
    Else
        Set infoDoc = Documents.Add '�����ھ��½�
        ActiveDocument.SaveAs wholeInfoFileName
    End If
    
    '3����infoDoc��д�뱾��ͳ�Ƶ���Ŀ��·������ʼʱ��
    With Selection
        .EndKey wdStory
        .TypeText Chr(13) & rootFolder.Name
        .Style = infoDoc.Styles("���� 1")
        .TypeParagraph
        .TypeText "�ű���ʼʱ�䣺" & startTime
        .TypeParagraph
        .TypeText "��Ŀ·��" & workPath
        .TypeParagraph
    End With
    
    '4��ͳ�ƹ���Ŀ¼���������ļ����е�Word�ļ���ҳ����ÿ�����ļ��е���Ϣ����һ�ű���
    '4.1---����һϵ�б������ڱ������ļ���ʱ��ʹ�õ�
    
    Dim filesDic As Object: Set filesDic = CreateObject("scripting.dictionary") '�����ļ�����Ϣ�ֵ�
    Dim quaDic As Object: Set quaDic = CreateObject("scripting.dictionary") '���������¼�ļ���ҳ����Ϣ
    Dim transDic As Object: Set transDic = CreateObject("scripting.dictionary") '����ƽ��嵥�ļ���ҳ����Ϣ
    
    Dim subFolder '�������ļ����ֵ�ʱʹ��
    Dim fileCount As Integer: fileCount = 1 '�ж�ѭ���������½������Ҫд���ͷ��Ϣ��ÿ�ζ�ȡ��һ���ļ�����Ҫ��ԭ
    
    
    '������Ҫ�õ��ı���
    Dim lastRow As row 'infoDoc�����һ��������һ��
    Dim fileName As String 'ʹ��dir�����ļ��л�ȡ���ļ���
    Dim docReading As Document '��Ҫ��ȡ��Ϣ���ĵ�
    Dim startPos As Integer 'ҳ����Ϣ�����
    Dim endPos As Integer 'ҳ����Ϣ���յ�
    Dim coverPage As Integer '�����ҳ��
    Dim pageInName As Integer '�ļ������ҳ��
    Dim realPage As Integer 'ʵ��ҳ��
    Dim errFileNum As Integer '����ҳ��������ļ�����
    Dim errNameNum As Integer '�ļ���ҳ��������ļ�����
    
    Dim RE As Object: Set RE = CreateObject("vbscript.regexp") '������ʽ
    
    '4.2---forѭ�������ֵ��еļ�ֵ�ԣ�ÿ����ֵ�Լ���ʾһ�����ļ���
    For Each subFolder In foldersDic
        
        '4.2.1 �������ļ���������ָ�����͵��ļ�����ȡҳ����д��infoDoc
        fileName = Dir(foldersDic(subFolder) & "\" & fileType)
        Do While fileName <> ""
            If fileName = ThisDocument.Name Then
                MsgBox ("���𽫽ű��ļ��������ļ�����")
                Exit Do
            End If
            
            '4.2.1.1 ������ļ�������ָ�����͵��ļ�����ͳ��ҳ��
            If fileCount = 1 Then
                '��һ��ѭ����ʱ����ϱ�ͷ
                ' ��infoDoc���������
                infoDoc.Activate
                Selection.EndKey wdStory
                Selection.TypeText subFolder
                Selection.Style = infoDoc.Styles("���� 6")
                Selection.TypeParagraph
                Selection.EndKey wdStory
                '����1��4�еı��д���ͷ��Ϣ
                Call insLastRow(infoDoc, "�ļ���", "��ʵҳ��", "��Ƥҳ��", "�ļ�����ҳ��", True)
                '����Ҫͳ����Ϣ���ļ�
                Set docReading = Documents.Open(foldersDic(subFolder) & "\" & fileName)
                realPage = docReading.ComputeStatistics(wdStatisticPages)
                Call insLastRow(infoDoc, docReading.Name, realPage)
                Set lastRow = infoDoc.Tables(infoDoc.Tables.count).Rows.Last
            '���ǵ�һ��ѭ��������д����Ϣ��infoDoc
            Else
                '����Ҫͳ����Ϣ���ļ�
                Set docReading = Documents.Open(foldersDic(subFolder) & "\" & fileName)
                realPage = docReading.ComputeStatistics(wdStatisticPages)
                
                Call insLastRow(infoDoc, docReading.Name, realPage)
                Set lastRow = infoDoc.Tables(infoDoc.Tables.count).Rows.Last
            End If
            
            '4.2.1.2 ��ȡ��Ƥ�����ҳ�������ж�������ʵҳ���Ƿ����
            If InStr(1, subFolder, "�����ĵ�") Then '���ļ����������������ĵ��������ҳ��
                docReading.Activate
                ' ��ȡҳ��
                With Selection.Find
                    .Forward = True
                    .ClearFormatting
                    .MatchWholeWord = True
                    .MatchCase = False
                    .Wrap = wdFindContinue
                    .Execute FindText:="ҳ��"
                End With
                Selection.MoveRight wdCharacter, 2
                startPos = Selection.Start
                Selection.EndKey wdLine
                endPos = Selection.Start
                
                coverPage = Val(docReading.Range(startPos, endPos).Text)
'                lastRow.Range.Cells(3).Range.Text = coverPage
                '�������ҳ������ʵҳ����һ��
                If coverPage > 0 And coverPage <> realPage Then
                    lastRow.Range.Cells(3).Range.Text = coverPage
                    lastRow.Range.Cells(3).Range.Text = lastRow.Range.Cells(3).Range.Text & "������ҳ�����󣡡�"
                    errFileNum = errFileNum + 1
                Else
                    lastRow.Range.Cells(3).Range.Text = "δ�ҵ�ҳ��..."
                End If
            End If
            
            
            '4.2.1.3 ��ȡ�ļ����������������ҳ�������ж�������ʵҳ���Ƿ����
            If InStr(1, fileName, "��") <> 0 And InStr(1, fileName, "��") <> 0 Then
                RE.Global = True
                RE.ignorecase = False
                RE.Pattern = "\d+(?=ҳ)"
                Set result = RE.Execute(fileName)
                '����ҵ�ҳ�������к�������
                If result.count > 0 Then
                    pageInName = Val(result(0))
                    lastRow.Cells(4).Range.Text = pageInName
                    '����ļ������ҳ������ʵҳ����һ��
                    If pageInName <> realPage Then
                        lastRow.Cells(4).Range.Text = lastRow.Cells(4).Range.Text & "���ļ���ҳ��������"
                        errNameNum = errNameNum + 1
                    End If
                Else
                    lastRow.Cells(4).Range.Text = "δ�ҵ�ҳ��..."
                End If
            End If
            
            '4.2.1.4 ͳ�Ƹ��ļ�ҳ���������ļ��ֵ���
            Dim fileInfoDic As Object: Set fileInfoDic = CreateObject("scripting.dictionary") '�����ļ�����Ϣ
            fileInfoDic.Add "filename", fileName
            fileInfoDic.Add "realpage", realPage
            fileInfoDic.Add "coverpage", coverPage
            fileInfoDic.Add "pageinname", pageInName
            '�ֵ��ֵ���ֵ䣬�����ֵ�
            filesDic.Add foldersDic(subFolder) & "\" & fileName, fileInfoDic
            
            '4.2.1.5 ��ȡ������¼Ŀ¼��ҳ��
            If InStr(1, fileName, "������¼") <> 0 And realPage = 2 Then
                docReading.Activate
                ' ������λ��ҳ��Ŀ¼
                With Selection.Find
                    .Forward = True
                    .ClearFormatting
                    .MatchWholeWord = True
                    .MatchCase = False
                    .Wrap = wdFindContinue
                    .Execute FindText:="������������"
                End With
                startPos = Selection.Start
                Selection.EndKey wdStory
                endPos = Selection.Start
                '���ҳ���ַ���������
                Dim QAInfoArr() As String: QAInfoArr = Split(docReading.Range(startPos, endPos).Text, Chr(13))
                Call printArr(QAInfoArr)
                RE.Global = True
                RE.ignorecase = False
                '�������зֱ���ȡҳ����һһ��Ӧ�ķ����ֵ�
                For Each qainfo In QAInfoArr
                    RE.Pattern = "\d+(?=ҳ)"
                    Set QApage = RE.Execute(qainfo)
                    If QApage.count > 0 Then
                        quaDic.Add qainfo, Val(QApage(0))
                    End If
                Next
            End If
            
            '4.2.1.6 ��ȡ�ƽ��嵥���ҳ��
            If InStr(1, fileName, "�ƽ��嵥") <> 0 And realPage = 1 Then
                'TransDic
                docReading.Activate
                Dim rowCount As Integer
                Dim transInfoArr(0 To 1) As String
                '�кϲ��ĵ�Ԫ���������õ�Ԫ��������������
                For rowCount = 1 To docReading.Tables(1).Rows.count Step 1
                    docReading.Tables(1).cell(rowCount, 2).Select 'ÿһ�еĵ�2����Ԫ���ǷǺϲ���Ԫ��
                    With Selection
                        .SelectRow '���ù̶��ĵ�Ԫ��ѡ����һ��
                        If .Cells.count = 8 Then '�޺ϲ���Ԫ�����
                            If InStr(1, .Cells(7).Range.Text, "����") Then
                                
                                transInfoArr(0) = Split(.Cells(5).Range.Text, Chr(13))(0) 'ҳ��
                                transInfoArr(1) = Split(.Cells(8).Range.Text, Chr(13))(0)  '��ע��Ϣ
                                transDic.Add Split(.Cells(3).Range.Text, Chr(13))(0), transInfoArr
                            End If
                        ElseIf .Cells.count = 7 Then '�кϲ���Ԫ�����
                            If InStr(1, .Cells(7).Range.Text, "����") Then
                                '�кϲ���Ԫ��Ļ���transInfoArr(2)���䣬��Ϊ��ע��Ϣ�ظ�
                                transInfoArr(0) = Split(.Cells(5).Range.Text, Chr(13))(0) 'ҳ��
                                transDic.Add Split(.Cells(3).Range.Text, Chr(13))(0), transInfoArr
                            End If
                        End If
                    End With
                Next rowCount
            End If
            
            '4.2.1.7 �رն�ȡ��Ϣ���ļ�
            docReading.Close 0
            fileCount = fileCount + 1
            fileName = Dir
            
        Loop '��ȡ��һ���ļ���Ϣ��������ȡ��һ���ļ���Ϣ
        
        fileCount = 1 '��ԭ�ļ���������������ı�ͷ��Ϣ�޷�����д��
        
    Next subFolder '������һ�����ļ����е��ļ���������һ�����ļ���
    

    '5 ����������¼��ҳ��
    infoDoc.Activate
    Selection.EndKey wdStory
    Selection.TypeText "��������¼��Ƥ��ҳ������"
    Selection.Style = infoDoc.Styles("���� 4")
    Selection.TypeParagraph
    Selection.EndKey wdStory
    '����1��4�еı��д���ͷ��Ϣ
    Call insLastRow(infoDoc, "�ļ���", "��ʵҳ��", "Ŀ¼ҳ��", "�ԱȽ��", True)
    For Each file In filesDic.keys
        Dim nameStr1() As String: nameStr1 = Split(file, "\")
        Dim tempName1: tempName1 = nameStr1(UBound(nameStr1))
        Dim QuaRes
        '�������ļ�����Ϣ�ֵ�ȥ�Ա�������¼�ļ���ҳ���ֵ�
        For Each tempName2 In quaDic.keys
            
            If InStr(1, tempName1, "��������") <> 0 And InStr(1, tempName2, "��������") <> 0 Then
                '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(quaDic(tempName2)) Then 'ʵ��ҳ����Ŀ¼ҳ����һ��
                    QuaRes = "������"
                    Call insLastRow(infoDoc, tempName1, filesDic(file)("realpage"), quaDic(tempName2), QuaRes, False)
                Else
                    QuaRes = "ok"
                    Call insLastRow(infoDoc, tempName1, filesDic(file)("realpage"), quaDic(tempName2), QuaRes, False)
                End If
            ElseIf InStr(1, tempName1, "�������") <> 0 And InStr(1, tempName2, "�������") <> 0 Then
                '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(quaDic(tempName2)) Then 'ʵ��ҳ����Ŀ¼ҳ����һ��
                    QuaRes = "������"
                    Call insLastRow(infoDoc, tempName1, filesDic(file)("realpage"), quaDic(tempName2), QuaRes, False)
                Else
                    QuaRes = "ok"
                    Call insLastRow(infoDoc, tempName1, filesDic(file)("realpage"), quaDic(tempName2), QuaRes, False)
                
                End If
            ElseIf InStr(1, tempName1, "XX�ļ����ؼ���") <> 0 And InStr(1, tempName2, "XX�ļ����ؼ���") <> 0 Then
            '''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''
            ''''''ͳ�Ƹ����ļ���ҳ�����
            '''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''
            End If
        Next tempName2
    Next file
    
    '6 �����ƽ��嵥��ҳ��
    infoDoc.Activate
    Selection.EndKey wdStory
    Selection.TypeText "���ƽ��嵥��ҳ������"
    Selection.Style = infoDoc.Styles("���� 4")
    Selection.TypeParagraph
    Selection.EndKey wdStory
    '����1��4�еı��д���ͷ��Ϣ
    Call insLastRow(infoDoc, "�ļ���", "��ʵҳ��", "ҳ����", "��ע�ͽ��", True)
    For Each file In filesDic.keys
        nameStr1 = Split(file, "\")
        tempName1 = nameStr1(UBound(nameStr1))
        Dim TransRes
        Dim tmpPage As Integer
        For Each tempName2 In transDic
            If InStr(1, tempName1, "�鵵˵��") <> 0 And InStr(1, tempName2, "�鵵˵��") <> 0 Then
            
                '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then 'ʵ��ҳ����ҳ�����Ĳ�һ��
                    TransRes = "��ҳ�������嵥ҳ����ʵ��ҳ����ͬ��"
                Else
                    QuaRes = "ҳ������ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            '��������Ҫ�ų���������ظ�������
            ElseIf InStr(1, tempName1, "���") <> 0 And InStr(1, tempName2, "���") <> 0 And InStr(1, tempName1, "����") = 0 And InStr(1, tempName1, "���") = 0 Then
                tmpPage = 0
                RE.Pattern = "\d+(?=ҳ)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count > 0 Then
                    TransRes = page1(0)
                Else
                    TransRes = "��������ʽͳ�Ʊ�ע��ҳ������"
                End If
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "��1����ע����ʵ��ҳ���ͱ�ע��ҳ��(" & TransRes & ")��ͬ��"
                Else
                    TransRes = "1����ע����ok"
                End If
                
                '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then 'ʵ��ҳ����ҳ�����Ĳ�һ��
                    TransRes = TransRes & Chr(13) & "��2��ҳ������ʵ��ҳ����ҳ�����Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "2��ҳ������ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "����˵��") <> 0 And InStr(1, tempName2, "����˵��") <> 0 And InStr(1, tempName1, "����") = 0 And InStr(1, tempName1, "���") = 0 Then
            
                 '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then 'ʵ��ҳ����Ŀ¼ҳ����һ��
                    TransRes = "��1��ҳ������ʵ��ҳ����ҳ�����Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "1��ҳ������ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "���Լ�¼") <> 0 And InStr(1, tempName2, "���Լ�¼") <> 0 And InStr(1, tempName1, "���") = 0 Then
            
                '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then 'ʵ��ҳ����Ŀ¼ҳ����һ��
                    TransRes = "��1��ҳ������ʵ��ҳ����ҳ�����Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "1��ҳ������ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "���ⱨ��") <> 0 And InStr(1, tempName2, "���ⱨ��") <> 0 And InStr(1, tempName1, "����") = 0 And InStr(1, tempName1, "���") = 0 Then
            
                tmpPage = 0
                RE.Pattern = "\d+(?=ҳ)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count = 2 Then
                    TransRes = page1(0)
                Else
                    TransRes = "��������ʽͳ�Ʊ�ע��ҳ������"
                End If
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "��1����ע����ʵ��ҳ���ͱ�ע��ҳ��(" & TransRes & ")��ͬ��"
                Else
                    TransRes = "1����ע����ok"
                End If
                
                '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then 'ʵ��ҳ����ҳ�����Ĳ�һ��
                    TransRes = TransRes & Chr(13) & "��2��ҳ������ʵ��ҳ����ҳ�����Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "2��ҳ������ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "��������") <> 0 And InStr(1, tempName2, "��������") <> 0 And InStr(1, tempName1, "����") = 0 And InStr(1, tempName1, "���") = 0 Then
                
                tmpPage = 0
                RE.Pattern = "\d+(?=ҳ)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count > 0 Then
                    TransRes = page1(0)
                Else
                    TransRes = "��������ʽͳ�Ʊ�ע��ҳ������"
                End If
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "��1����ע����ʵ��ҳ���ͱ�ע��ҳ��(" & TransRes & ")��ͬ��"
                Else
                    TransRes = "1����ע����ok"
                End If
                
                '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then 'ʵ��ҳ����ҳ�����Ĳ�һ��
                    TransRes = TransRes & Chr(13) & "��2��ҳ������ʵ��ҳ����ҳ�����Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "2��ҳ������ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "���Ա���") <> 0 And InStr(1, tempName2, "���Ա���") <> 0 And InStr(1, tempName1, "����") = 0 And InStr(1, tempName1, "���") = 0 Then
            
                tmpPage = 0
                RE.Pattern = "\d+(?=ҳ)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count > 0 Then
                    TransRes = page1(0)
                Else
                    TransRes = "��������ʽͳ�Ʊ�ע��ҳ������"
                End If
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "��1����ע����ʵ��ҳ���ͱ�ע��ҳ��(" & TransRes & ")��ͬ��"
                Else
                    TransRes = "1����ע����ok"
                End If
                
                '�����±�񣬴��������¼�ļ����������
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then 'ʵ��ҳ����ҳ�����Ĳ�һ��
                    TransRes = TransRes & Chr(13) & "��2��ҳ������ʵ��ҳ����ҳ�����Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "2��ҳ������ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "������¼") <> 0 And InStr(1, tempName2, "������¼") Then
            
                tmpPage = 0
                RE.Pattern = "\d+(?=ҳ)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count = 3 Then
                    TransRes = page1(0)
                Else
                    TransRes = "��������ʽͳ�Ʊ�ע��ҳ������"
                End If
                '---ͳ��������¼�ļ������ҳ������ȷ����2ҳ
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "��1����ע����������¼�����ʵ��ҳ���ͱ�ע��ҳ��(" & TransRes & ")��ͬ��"
                Else
                    TransRes = "1����ע����������¼����ҳ��ok"
                End If
                '---ͳ��2������֤���ҳ��֮�͡������ļ�ҳ��֮��
                '���ݹؼ��ʡ�֤�顱����quaDic�����ҳ����
                Dim certificatesPage As Integer, noSecretsPage As Integer, quasPage As Integer
                For Each temp1 In quaDic
                    If InStr(1, temp1, "֤��") Then '����֤���ҳ����
                        certificatesPage = certificatesPage + Val(quaDic(temp1))
                    End If
                    If InStr(1, temp1, "֤��") = 0 Then '��֤�鼴���ܵ�ҳ����
                        noSecretsPage = noSecretsPage + Val(quaDic(temp1))
                    End If
                    quasPage = quasPage + Val(quaDic(temp1))
                Next temp1
                '��֤һ��ҳ��
                If certificatesPage + noSecretsPage <> quasPage Then
                    MsgBox "ͳ�������ļ����ܺ�ҳ������"
                    certificatesPage = "��error��"
                    noSecretsPage = "��error��"
                    quasPage = "��error��"
                    TransRes = TransRes & quasPage
                End If
                quasPage = quasPage + Val(filesDic(file)("realpage")) '���յ���ҳ����Ҫ�����Լ������ҳ��
                '�ȶ�������¼��Ƥ�ļ���Ŀ¼ҳ�����ƽ��嵥��ҳ��
                '�����±�񣬴��������¼�ļ����������
                If quasPage <> Val(transDic(tempName2)(0)) Then 'ʵ��ҳ����ҳ�����Ĳ�һ��
                    TransRes = TransRes & Chr(13) & "��2��ҳ������������¼�����ļ�����ҳ����ҳ�����Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "2��ҳ������������¼�����ļ�����ҳ��ok"
                End If
                If Val(page1(1)) <> certificatesPage Then 'ҳ������֤���ҳ������
                    TransRes = TransRes & Chr(13) & "��3����ע����������¼�е�֤��ҳ���ͱ�ע���Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "3����ע����������¼�е�֤��ҳ��ok"
                End If
                If Val(page1(2)) <> noSecretsPage Then 'ҳ�����ķ����ļ���ҳ���д�
                    TransRes = TransRes & Chr(13) & "��4����ע����������¼�з����ļ���ҳ���ͱ�ע���Ĳ�һ�¡�"
                Else
                    TransRes = TransRes & Chr(13) & "4����ע����������¼�з����ļ���ҳ��ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            End If
        Next tempName2
    Next file
    
    '7 ͳ�ƴ���������������ʱ��
    infoDoc.Activate
    Dim endTime: endTime = Now
    With Selection
        .EndKey wdStory
        .TypeText "��Ƥҳ��������ĵ�������" & errFileNum
        .Style = infoDoc.Styles("���� 6")
        .TypeParagraph
        .TypeText "�ļ���ҳ��������ĵ�������" & errNameNum
        .Style = infoDoc.Styles("���� 6")
        .TypeParagraph
        .TypeText "�ű�����ʱ��" & endTime & Chr(13)
        .TypeText "�ű����к�ʱ" & Abs(DateDiff("s", startTime, endTime)) & "��" & Chr(13)
    End With
    infoDoc.Save
    
End Sub

Function insLastRow(ByRef infoDoc As Document, Optional str1 = "", Optional str2 = "", Optional str3 = "", Optional str4 = "", Optional newTable As Boolean = False)
    infoDoc.Activate
    If newTable = True Then
        Selection.EndKey wdStory
'        Selection.TypeParagraph
        '����1��4�еı��д���ͷ��Ϣ
        infoDoc.Tables.Add Selection.Range, 1, 4, wdWord9TableBehavior, wdAutoFitFixed
        With infoDoc.Tables(infoDoc.Tables.count).Rows.First
            .Cells(1).Range.Text = str1
            .Cells(2).Range.Text = str2
            .Cells(3).Range.Text = str3
            .Cells(4).Range.Text = str4
        End With
        infoDoc.Save '����
        Exit Function
    ElseIf infoDoc.Tables.count = 0 Then
        Exit Function
    ElseIf infoDoc.Tables(infoDoc.Tables.count).Rows.Last.Range.Cells.count <> 4 Then
        MsgBox "insLastRow()��������������ĵ������һ��������һ�еĵ�Ԫ�������Ϊ4��"
        Exit Function
    End If
    
    infoDoc.Tables(infoDoc.Tables.count).Rows.Last.Select
    Selection.InsertRowsBelow
'    Set lastRow = infoDoc.Tables(infoDoc.Tables.count).Rows.Last
    With infoDoc.Tables(infoDoc.Tables.count).Rows.Last.Range
        .Cells(1).Range.Text = str1
        .Cells(2).Range.Text = str2
        .Cells(3).Range.Text = str3
        .Cells(4).Range.Text = str4
    End With
    infoDoc.Save '����
End Function
