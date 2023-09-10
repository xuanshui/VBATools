Sub 统计文件页数_最终版()

    Dim workPath As String: workPath = ThisDocument.Path & "\"
    Dim fileType As String: fileType = "*.doc?"
    Dim infoFileName As String: infoFileName = "页数统计.doc"
    Dim wholeInfoFileName As String: wholeInfoFileName = ThisDocument.Path & "\" & infoFileName
    
    Dim startTime: startTime = Now
    
    '1、获取工作目录下的第一层子目录名称，存放进字典中
    Dim FSO As Object: Set FSO = CreateObject("scripting.filesystemobject")
    Dim rootFolder As Object: Set rootFolder = FSO.GetFolder(workPath)
    Dim subFolders As Object: Set subFolders = rootFolder.subFolders
    Dim tmpFolder
    '子文件夹信息
    Dim foldersDic As Object: Set foldersDic = CreateObject("scripting.dictionary")
    For Each tmpFolder In subFolders
        foldersDic.Add tmpFolder.Name, tmpFolder.Path
    Next tmpFolder

    
    '2、判断保存信息的文件infoDoc是否存在
    Dim infoDoc As Document
    If FSO.FileExists(wholeInfoFileName) Then '存在就打开
        Set infoDoc = Documents.Open(wholeInfoFileName)
    Else
        Set infoDoc = Documents.Add '不存在就新建
        ActiveDocument.SaveAs wholeInfoFileName
    End If
    
    '3、在infoDoc中写入本次统计的项目的路径、开始时间
    With Selection
        .EndKey wdStory
        .TypeText Chr(13) & rootFolder.Name
        .Style = infoDoc.Styles("标题 1")
        .TypeParagraph
        .TypeText "脚本开始时间：" & startTime
        .TypeParagraph
        .TypeText "项目路径" & workPath
        .TypeParagraph
    End With
    
    '4、统计工作目录下所有子文件夹中的Word文件的页数，每个子文件夹的信息都用一张表存放
    '4.1---定义一系列变量，在遍历子文件夹时会使用到
    
    Dim filesDic As Object: Set filesDic = CreateObject("scripting.dictionary") '所有文件的信息字典
    Dim quaDic As Object: Set quaDic = CreateObject("scripting.dictionary") '存放质量记录文件的页数信息
    Dim transDic As Object: Set transDic = CreateObject("scripting.dictionary") '存放移交清单文件的页数信息
    
    Dim subFolder '遍历子文件夹字典时使用
    Dim fileCount As Integer: fileCount = 1 '判断循环次数，新建表格需要写入表头信息，每次读取完一个文件都需要复原
    
    
    '操作需要用到的变量
    Dim lastRow As row 'infoDoc的最后一个表的最后一行
    Dim fileName As String '使用dir从子文件夹获取的文件名
    Dim docReading As Document '需要读取信息的文档
    Dim startPos As Integer '页数信息的起点
    Dim endPos As Integer '页数信息的终点
    Dim coverPage As Integer '封面的页数
    Dim pageInName As Integer '文件名里的页数
    Dim realPage As Integer '实际页数
    Dim errFileNum As Integer '封面页数出错的文件个数
    Dim errNameNum As Integer '文件名页数出错的文件个数
    
    Dim RE As Object: Set RE = CreateObject("vbscript.regexp") '正则表达式
    
    '4.2---for循环遍历字典中的键值对，每个键值对即表示一个子文件夹
    For Each subFolder In foldersDic
        
        '4.2.1 遍历子文件夹下所有指定类型的文件，获取页数，写入infoDoc
        fileName = Dir(foldersDic(subFolder) & "\" & fileType)
        Do While fileName <> ""
            If fileName = ThisDocument.Name Then
                MsgBox ("请勿将脚本文件放在子文件夹中")
                Exit Do
            End If
            
            '4.2.1.1 如果子文件夹中有指定类型的文件，才统计页数
            If fileCount = 1 Then
                '第一次循环的时候加上表头
                ' 打开infoDoc，创建表格
                infoDoc.Activate
                Selection.EndKey wdStory
                Selection.TypeText subFolder
                Selection.Style = infoDoc.Styles("标题 6")
                Selection.TypeParagraph
                Selection.EndKey wdStory
                '创建1行4列的表格，写入表头信息
                Call insLastRow(infoDoc, "文件名", "真实页数", "封皮页数", "文件名的页数", True)
                '打开需要统计信息的文件
                Set docReading = Documents.Open(foldersDic(subFolder) & "\" & fileName)
                realPage = docReading.ComputeStatistics(wdStatisticPages)
                Call insLastRow(infoDoc, docReading.Name, realPage)
                Set lastRow = infoDoc.Tables(infoDoc.Tables.count).Rows.Last
            '不是第一次循环，正常写入信息到infoDoc
            Else
                '打开需要统计信息的文件
                Set docReading = Documents.Open(foldersDic(subFolder) & "\" & fileName)
                realPage = docReading.ComputeStatistics(wdStatisticPages)
                
                Call insLastRow(infoDoc, docReading.Name, realPage)
                Set lastRow = infoDoc.Tables(infoDoc.Tables.count).Rows.Last
            End If
            
            '4.2.1.2 获取封皮上面的页数，并判断它和真实页数是否相等
            If InStr(1, subFolder, "测试文档") Then '子文件夹名包含“测试文档”，检查页数
                docReading.Activate
                ' 获取页数
                With Selection.Find
                    .Forward = True
                    .ClearFormatting
                    .MatchWholeWord = True
                    .MatchCase = False
                    .Wrap = wdFindContinue
                    .Execute FindText:="页数"
                End With
                Selection.MoveRight wdCharacter, 2
                startPos = Selection.Start
                Selection.EndKey wdLine
                endPos = Selection.Start
                
                coverPage = Val(docReading.Range(startPos, endPos).Text)
'                lastRow.Range.Cells(3).Range.Text = coverPage
                '如果封面页数和真实页数不一致
                If coverPage > 0 And coverPage <> realPage Then
                    lastRow.Range.Cells(3).Range.Text = coverPage
                    lastRow.Range.Cells(3).Range.Text = lastRow.Range.Cells(3).Range.Text & "【封面页数错误！】"
                    errFileNum = errFileNum + 1
                Else
                    lastRow.Range.Cells(3).Range.Text = "未找到页数..."
                End If
            End If
            
            
            '4.2.1.3 获取文件名“【】”里面的页数，并判断它和真实页数是否相等
            If InStr(1, fileName, "【") <> 0 And InStr(1, fileName, "】") <> 0 Then
                RE.Global = True
                RE.ignorecase = False
                RE.Pattern = "\d+(?=页)"
                Set result = RE.Execute(fileName)
                '如果找到页数，才有后续步骤
                If result.count > 0 Then
                    pageInName = Val(result(0))
                    lastRow.Cells(4).Range.Text = pageInName
                    '如果文件名里的页数和真实页数不一致
                    If pageInName <> realPage Then
                        lastRow.Cells(4).Range.Text = lastRow.Cells(4).Range.Text & "【文件名页数出错！】"
                        errNameNum = errNameNum + 1
                    End If
                Else
                    lastRow.Cells(4).Range.Text = "未找到页数..."
                End If
            End If
            
            '4.2.1.4 统计各文件页数，放入文件字典中
            Dim fileInfoDic As Object: Set fileInfoDic = CreateObject("scripting.dictionary") '单个文件的信息
            fileInfoDic.Add "filename", fileName
            fileInfoDic.Add "realpage", realPage
            fileInfoDic.Add "coverpage", coverPage
            fileInfoDic.Add "pageinname", pageInName
            '字典的值是字典，二重字典
            filesDic.Add foldersDic(subFolder) & "\" & fileName, fileInfoDic
            
            '4.2.1.5 获取质量记录目录的页数
            If InStr(1, fileName, "质量记录") <> 0 And realPage = 2 Then
                docReading.Activate
                ' 搜索定位到页数目录
                With Selection.Find
                    .Forward = True
                    .ClearFormatting
                    .MatchWholeWord = True
                    .MatchCase = False
                    .Wrap = wdFindContinue
                    .Execute FindText:="测试任务工作单"
                End With
                startPos = Selection.Start
                Selection.EndKey wdStory
                endPos = Selection.Start
                '拆分页数字符串到数组
                Dim QAInfoArr() As String: QAInfoArr = Split(docReading.Range(startPos, endPos).Text, Chr(13))
                Call printArr(QAInfoArr)
                RE.Global = True
                RE.ignorecase = False
                '从数组中分别提取页数，一一对应的放入字典
                For Each qainfo In QAInfoArr
                    RE.Pattern = "\d+(?=页)"
                    Set QApage = RE.Execute(qainfo)
                    If QApage.count > 0 Then
                        quaDic.Add qainfo, Val(QApage(0))
                    End If
                Next
            End If
            
            '4.2.1.6 获取移交清单里的页数
            If InStr(1, fileName, "移交清单") <> 0 And realPage = 1 Then
                'TransDic
                docReading.Activate
                Dim rowCount As Integer
                Dim transInfoArr(0 To 1) As String
                '有合并的单元格，所以利用单元格来遍历表格的行
                For rowCount = 1 To docReading.Tables(1).Rows.count Step 1
                    docReading.Tables(1).cell(rowCount, 2).Select '每一行的第2个单元格都是非合并单元格
                    With Selection
                        .SelectRow '利用固定的单元格选中这一行
                        If .Cells.count = 8 Then '无合并单元格的行
                            If InStr(1, .Cells(7).Range.Text, "短期") Then
                                
                                transInfoArr(0) = Split(.Cells(5).Range.Text, Chr(13))(0) '页数
                                transInfoArr(1) = Split(.Cells(8).Range.Text, Chr(13))(0)  '备注信息
                                transDic.Add Split(.Cells(3).Range.Text, Chr(13))(0), transInfoArr
                            End If
                        ElseIf .Cells.count = 7 Then '有合并单元格的行
                            If InStr(1, .Cells(7).Range.Text, "短期") Then
                                '有合并单元格的话，transInfoArr(2)不变，因为备注信息重复
                                transInfoArr(0) = Split(.Cells(5).Range.Text, Chr(13))(0) '页数
                                transDic.Add Split(.Cells(3).Range.Text, Chr(13))(0), transInfoArr
                            End If
                        End If
                    End With
                Next rowCount
            End If
            
            '4.2.1.7 关闭读取信息的文件
            docReading.Close 0
            fileCount = fileCount + 1
            fileName = Dir
            
        Loop '读取完一个文件信息，继续读取下一个文件信息
        
        fileCount = 1 '还原文件计数器，否则表格的表头信息无法正常写入
        
    Next subFolder '遍历完一个子文件夹中的文件，遍历下一个子文件夹
    

    '5 计算质量记录的页数
    infoDoc.Activate
    Selection.EndKey wdStory
    Selection.TypeText "《质量记录封皮》页数检查表"
    Selection.Style = infoDoc.Styles("标题 4")
    Selection.TypeParagraph
    Selection.EndKey wdStory
    '创建1行4列的表格，写入表头信息
    Call insLastRow(infoDoc, "文件名", "真实页数", "目录页数", "对比结果", True)
    For Each file In filesDic.keys
        Dim nameStr1() As String: nameStr1 = Split(file, "\")
        Dim tempName1: tempName1 = nameStr1(UBound(nameStr1))
        Dim QuaRes
        '用所有文件的信息字典去对比质量记录文件的页数字典
        For Each tempName2 In quaDic.keys
            
            If InStr(1, tempName1, "任务工作单") <> 0 And InStr(1, tempName2, "任务工作单") <> 0 Then
                '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(quaDic(tempName2)) Then '实际页数和目录页数不一致
                    QuaRes = "【错误】"
                    Call insLastRow(infoDoc, tempName1, filesDic(file)("realpage"), quaDic(tempName2), QuaRes, False)
                Else
                    QuaRes = "ok"
                    Call insLastRow(infoDoc, tempName1, filesDic(file)("realpage"), quaDic(tempName2), QuaRes, False)
                End If
            ElseIf InStr(1, tempName1, "大纲评审单") <> 0 And InStr(1, tempName2, "大纲评审单") <> 0 Then
                '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(quaDic(tempName2)) Then '实际页数和目录页数不一致
                    QuaRes = "【错误】"
                    Call insLastRow(infoDoc, tempName1, filesDic(file)("realpage"), quaDic(tempName2), QuaRes, False)
                Else
                    QuaRes = "ok"
                    Call insLastRow(infoDoc, tempName1, filesDic(file)("realpage"), quaDic(tempName2), QuaRes, False)
                
                End If
            ElseIf InStr(1, tempName1, "XX文件名关键词") <> 0 And InStr(1, tempName2, "XX文件名关键词") <> 0 Then
            '''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''
            ''''''统计各个文件的页数情况
            '''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''
            End If
        Next tempName2
    Next file
    
    '6 计算移交清单的页数
    infoDoc.Activate
    Selection.EndKey wdStory
    Selection.TypeText "《移交清单》页数检查表"
    Selection.Style = infoDoc.Styles("标题 4")
    Selection.TypeParagraph
    Selection.EndKey wdStory
    '创建1行4列的表格，写入表头信息
    Call insLastRow(infoDoc, "文件名", "真实页数", "页数栏", "备注和结果", True)
    For Each file In filesDic.keys
        nameStr1 = Split(file, "\")
        tempName1 = nameStr1(UBound(nameStr1))
        Dim TransRes
        Dim tmpPage As Integer
        For Each tempName2 In transDic
            If InStr(1, tempName1, "归档说明") <> 0 And InStr(1, tempName2, "归档说明") <> 0 Then
            
                '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then '实际页数和页数栏的不一致
                    TransRes = "【页数栏：清单页数和实际页数不同】"
                Else
                    QuaRes = "页数栏：ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            '！！！还要排除出入库表的重复！！！
            ElseIf InStr(1, tempName1, "大纲") <> 0 And InStr(1, tempName2, "大纲") <> 0 And InStr(1, tempName1, "评审") = 0 And InStr(1, tempName1, "库表") = 0 Then
                tmpPage = 0
                RE.Pattern = "\d+(?=页)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count > 0 Then
                    TransRes = page1(0)
                Else
                    TransRes = "【正则表达式统计备注栏页数出错】"
                End If
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "【1、备注栏：实际页数和备注栏页数(" & TransRes & ")不同】"
                Else
                    TransRes = "1、备注栏：ok"
                End If
                
                '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then '实际页数和页数栏的不一致
                    TransRes = TransRes & Chr(13) & "【2、页数栏：实际页数和页数栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "2、页数栏：ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "测试说明") <> 0 And InStr(1, tempName2, "测试说明") <> 0 And InStr(1, tempName1, "评审") = 0 And InStr(1, tempName1, "库表") = 0 Then
            
                 '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then '实际页数和目录页数不一致
                    TransRes = "【1、页数栏：实际页数和页数栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "1、页数栏：ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "测试记录") <> 0 And InStr(1, tempName2, "测试记录") <> 0 And InStr(1, tempName1, "库表") = 0 Then
            
                '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then '实际页数和目录页数不一致
                    TransRes = "【1、页数栏：实际页数和页数栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "1、页数栏：ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "问题报告") <> 0 And InStr(1, tempName2, "问题报告") <> 0 And InStr(1, tempName1, "评审") = 0 And InStr(1, tempName1, "库表") = 0 Then
            
                tmpPage = 0
                RE.Pattern = "\d+(?=页)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count = 2 Then
                    TransRes = page1(0)
                Else
                    TransRes = "【正则表达式统计备注栏页数出错】"
                End If
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "【1、备注栏：实际页数和备注栏页数(" & TransRes & ")不同】"
                Else
                    TransRes = "1、备注栏：ok"
                End If
                
                '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then '实际页数和页数栏的不一致
                    TransRes = TransRes & Chr(13) & "【2、页数栏：实际页数和页数栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "2、页数栏：ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "测评报告") <> 0 And InStr(1, tempName2, "测评报告") <> 0 And InStr(1, tempName1, "评审") = 0 And InStr(1, tempName1, "库表") = 0 Then
                
                tmpPage = 0
                RE.Pattern = "\d+(?=页)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count > 0 Then
                    TransRes = page1(0)
                Else
                    TransRes = "【正则表达式统计备注栏页数出错】"
                End If
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "【1、备注栏：实际页数和备注栏页数(" & TransRes & ")不同】"
                Else
                    TransRes = "1、备注栏：ok"
                End If
                
                '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then '实际页数和页数栏的不一致
                    TransRes = TransRes & Chr(13) & "【2、页数栏：实际页数和页数栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "2、页数栏：ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "测试报告") <> 0 And InStr(1, tempName2, "测试报告") <> 0 And InStr(1, tempName1, "评审") = 0 And InStr(1, tempName1, "库表") = 0 Then
            
                tmpPage = 0
                RE.Pattern = "\d+(?=页)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count > 0 Then
                    TransRes = page1(0)
                Else
                    TransRes = "【正则表达式统计备注栏页数出错】"
                End If
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "【1、备注栏：实际页数和备注栏页数(" & TransRes & ")不同】"
                Else
                    TransRes = "1、备注栏：ok"
                End If
                
                '插入新表格，存放质量记录文件的正误情况
                If Val(filesDic(file)("realpage")) <> Val(transDic(tempName2)(0)) Then '实际页数和页数栏的不一致
                    TransRes = TransRes & Chr(13) & "【2、页数栏：实际页数和页数栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "2、页数栏：ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            ElseIf InStr(1, tempName1, "质量记录") <> 0 And InStr(1, tempName2, "质量记录") Then
            
                tmpPage = 0
                RE.Pattern = "\d+(?=页)"
                Set page1 = RE.Execute(transDic(tempName2)(1))
                TransRes = ""
                If page1.count = 3 Then
                    TransRes = page1(0)
                Else
                    TransRes = "【正则表达式统计备注栏页数出错】"
                End If
                '---统计质量记录文件本身的页数，正确的是2页
                If Val(TransRes) <> filesDic(file)("realpage") Then
                    TransRes = "【1、备注栏：质量记录封面的实际页数和备注栏页数(" & TransRes & ")不同】"
                Else
                    TransRes = "1、备注栏：质量记录封面页数ok"
                End If
                '---统计2个评审证书的页数之和、非密文件页数之和
                '根据关键词“证书”计算quaDic里面的页数，
                Dim certificatesPage As Integer, noSecretsPage As Integer, quasPage As Integer
                For Each temp1 In quaDic
                    If InStr(1, temp1, "证书") Then '两个证书的页数和
                        certificatesPage = certificatesPage + Val(quaDic(temp1))
                    End If
                    If InStr(1, temp1, "证书") = 0 Then '非证书即非密的页数和
                        noSecretsPage = noSecretsPage + Val(quaDic(temp1))
                    End If
                    quasPage = quasPage + Val(quaDic(temp1))
                Next temp1
                '验证一下页数
                If certificatesPage + noSecretsPage <> quasPage Then
                    MsgBox "统计质量文件的总和页数有误"
                    certificatesPage = "【error】"
                    noSecretsPage = "【error】"
                    quasPage = "【error】"
                    TransRes = TransRes & quasPage
                End If
                quasPage = quasPage + Val(filesDic(file)("realpage")) '最终的总页数还要加上自己本身的页数
                '比对质量记录封皮文件的目录页数和移交清单的页数
                '插入新表格，存放质量记录文件的正误情况
                If quasPage <> Val(transDic(tempName2)(0)) Then '实际页数和页数栏的不一致
                    TransRes = TransRes & Chr(13) & "【2、页数栏：质量记录所有文件的总页数和页数栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "2、页数栏：质量记录所有文件的总页数ok"
                End If
                If Val(page1(1)) <> certificatesPage Then '页数栏的证书的页数出错
                    TransRes = TransRes & Chr(13) & "【3、备注栏：质量记录中的证书页数和备注栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "3、备注栏：质量记录中的证书页数ok"
                End If
                If Val(page1(2)) <> noSecretsPage Then '页数栏的非密文件的页数有错
                    TransRes = TransRes & Chr(13) & "【4、备注栏：质量记录中非密文件总页数和备注栏的不一致】"
                Else
                    TransRes = TransRes & Chr(13) & "4、备注栏：质量记录中非密文件总页数ok"
                End If
                Call insLastRow(infoDoc, tempName2, filesDic(file)("realpage"), Val(transDic(tempName2)(0)), TransRes, False)
            End If
        Next tempName2
    Next file
    
    '7 统计错误情况，保存结束时间
    infoDoc.Activate
    Dim endTime: endTime = Now
    With Selection
        .EndKey wdStory
        .TypeText "封皮页数出错的文档个数：" & errFileNum
        .Style = infoDoc.Styles("标题 6")
        .TypeParagraph
        .TypeText "文件名页数出错的文档个数：" & errNameNum
        .Style = infoDoc.Styles("标题 6")
        .TypeParagraph
        .TypeText "脚本结束时间" & endTime & Chr(13)
        .TypeText "脚本运行耗时" & Abs(DateDiff("s", startTime, endTime)) & "秒" & Chr(13)
    End With
    infoDoc.Save
    
End Sub

Function insLastRow(ByRef infoDoc As Document, Optional str1 = "", Optional str2 = "", Optional str3 = "", Optional str4 = "", Optional newTable As Boolean = False)
    infoDoc.Activate
    If newTable = True Then
        Selection.EndKey wdStory
'        Selection.TypeParagraph
        '创建1行4列的表格，写入表头信息
        infoDoc.Tables.Add Selection.Range, 1, 4, wdWord9TableBehavior, wdAutoFitFixed
        With infoDoc.Tables(infoDoc.Tables.count).Rows.First
            .Cells(1).Range.Text = str1
            .Cells(2).Range.Text = str2
            .Cells(3).Range.Text = str3
            .Cells(4).Range.Text = str4
        End With
        infoDoc.Save '保存
        Exit Function
    ElseIf infoDoc.Tables.count = 0 Then
        Exit Function
    ElseIf infoDoc.Tables(infoDoc.Tables.count).Rows.Last.Range.Cells.count <> 4 Then
        MsgBox "insLastRow()函数出错：传入的文档的最后一个表格最后一行的单元格个数不为4！"
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
    infoDoc.Save '保存
End Function
