Attribute VB_Name = "JL_EXCL_WD"
'-----------------------------------------------------------------------------------------------------------------------------------------
'version:       1.0.0
'author:        lihong
'update date:   20230902
'function��
'1����Excel����Լ�¼��ȡ���еĲ���������Ϣ
'2�����Excel����Լ�¼�еĲ��ֵͼ�����
'3���Ѳ���������Ϣд��Word��Ĳ��Լ�¼�������ͨ����һ���ű��Զ����ɵĲ��Լ�¼��
'
'
'
'-----------------------------------------------------------------------------------------------------------------------------------------

Option Base 1
Option Explicit 'ǿ�Ƴ����еı�������Ϊ��ʽ������������ʹ��Dim����ReDimŵ����������

'-----------------------------------------------------------------------------------------------------------------------------------------
'           ���ļ�����
'-----------------------------------------------------------------------------------------------------------------------------------------
Dim g_JL_Excel As Object
Dim g_JL_Word As Document
Public ExcelApp As Object
Public WordApp As Object
Dim workPath As String
'-----------------------------------------------------------------------------------------------------------------------------------------
'           ��Excel������
'-----------------------------------------------------------------------------------------------------------------------------------------
Dim g_arr_num_excel_rowInSheets() As Long
Dim g_arr_num_excel_colInSheets() As Long
Dim g_num_excel_actual_row As Long
Dim g_num_excel_actual_column As Integer
'-----------------------------------------------------------------------------------------------------------------------------------------
'           �����ݽṹ����Ų���������������Ϣ
'-----------------------------------------------------------------------------------------------------------------------------------------
Public g_arr_testCase() As String
Public g_dic_subItem_testCase As Object
Public g_arr_testCase_key() As String                 '�����飺ÿ��sheet�ĵ�һ�е�ÿ����Ԫ�����ݾ���ÿ�������ֵ�ļ�
'-----------------------------------------------------------------------------------------------------------------------------------------
'           ������ͳ��
'-----------------------------------------------------------------------------------------------------------------------------------------
Dim g_num_err_excel As Integer
Dim g_arr_err_excel() As String
'-----------------------------------------------------------------------------------------------------------------------------------------
'           ����ѡ����
'-----------------------------------------------------------------------------------------------------------------------------------------
Const C_excel_start_row As Integer = 2      'Excel���Լ�¼�Ĳ���������ʵ����ʼ��
Const C_opt As Boolean = False              '�ű��ٶ��Ż�����

Const C_word_rowNum_caseName As Integer = 1         'Word����еġ������������ơ��ڵ�һ��
Const C_word_rowNum_caseCode As Integer = 2         'Word����еġ�����������ʶ���ڵڶ���
Const C_word_rowNum_caseInit As Integer = 3         'Word����еġ�����������ʼ�����ڵڶ���
Const C_word_rowNum_premise  As Integer = 6         'Word����еġ�ǰ����Լ�����ڵ�����
Const C_word_rowNum_caseStart As Integer = 10       'Word����еĲ�������ʵ�����ݴӵ�ʮ�п�ʼ
Const C_word_rowNum_designer As Integer = 11        'Word����еġ������Ա���ڵ�ʮһ��
Const C_word_rowNum_tester As Integer = 12          'Word����еġ�������Ա���ڵ�ʮ����

Const C_excel_fileName As String = "���Լ�¼.xlsx"
Const C_word_fileName As String = "���Լ�¼.doc"
Const C_word_fileName_saved As String = "���Լ�¼-�ű���д.doc"
'Excel���Լ�¼���������Ӧ�������
Public Enum E_colNum_JLExcel
    subItemCode = 2                     '��2�У���������ı�ʶ��
    caseNumInSubItem = 3                '��3�У�ÿ�����������µĲ����������
    caseName = 4                        '��4�У�������������
    casePass = 8                        '��8�У����������Ƿ�ͨ��
End Enum
'Excel���Լ�¼����Ĵ������
Private Enum E_errCode_JL
    'Excel���Լ�¼����Ĵ������
    subItemCode_Null             '"��������Ϊ��"
    subItemCode_Repeat           '"���������ظ�"
    caseNumInSubItem_NotNum      '"������Ų�������"
    caseNumInSubItem_Diff        '"���������ʵ�ʲ�һ��"
    caseName_Null                '"��������Ϊ��"
    caseName_SeqNull             '"�������Ƶ����Ϊ��"
    caseName_SeqNotNum           '"�������Ƶ���Ų�������"
    caseName_SeqDiff             '"�������Ƶ������ʵ�ʲ�һ��"
    casePass_Null                '"���������Ľ���Ϊ��"
    casePass_Wrong               '"���������Ľ��۲���"ͨ��"��"δͨ��"��"��ͨ��"��"����Ľ�"
    
    '��Excel���Լ�¼��ȡ��Ϣ������Word���Լ�¼����ʱ�Ĵ������
    subItemNum_Diff              'Excel��Word�����汾�Ĳ��Լ�¼�еĲ���������������ͬ
End Enum

'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------
'           ������
'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------
Sub JL_EXCL_WD()
    '0�������쳣��������
    'On Error GoTo errhandlemain

    '1����ʼ��������������ʾ���жϰ汾�����ù�����ͼ��
    init

    '2�����Excel����Լ�¼�еĴ���д��һ��txt�ļ���Ȼ����ʾ������Ա�޸�Excel���
    checkErr

    '3����ȡexcel����еĲ���������Ϣ����������g_dic_subItem
    readDataFromExcel
    
    '4���ر��Ҳ�����Excel��Ĳ��Լ�¼
    'Call printDic2(g_dic_subItem_testCase)
    '5����дword��Ĳ��Լ�¼
    writeDataToWord

    '6���رպͱ�������ļ������word��Ĳ��Լ�¼
    saveJL

    Exit Sub

errhandlemain:
    Debug.Print Chr(13) & "----------- Error Happend! ----------- " & Chr(13)
    Debug.Print "���� " & Err.Number & " : " & Err.Description & Chr(13) & "����λ : " & Err.Source
    If Err.Number = 4608 Then
    Else
        'err.raise err.number
    End If
    Debug.Print Chr(13) & "-----------   Error  End   ----------- " & Chr(13)
    Resume Next
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
'           1��Init����ʼ��
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub init()

    workPath = ThisDocument.Path
    
    Set ExcelApp = CreateObject("excel.application")
    'Set WordApp = CreateObject("word.application")
    
    '3���򿪲��Լ�¼
    Set g_JL_Excel = ExcelApp.workbooks.Open(workPath & "\" & C_excel_fileName)
    ExcelApp.Visible = True
    Set g_JL_Word = Application.Documents.Open(workPath & "\" & C_word_fileName)
    
    
    g_JL_Word.Activate
    Selection.HomeKey wdStory
    
    '5����ѡ�Ż�����Word���¼תΪ��ͨ��ͼ��ʡȥWord��ҳ����������ٶ�
    If C_opt Then
        g_JL_Word.ActiveWindow.View.Type = wdNormalView
    End If
    
    '6����ȡExcel���������������Ϣ��Ȼ�����Excelʵ��������ʼ���������������顢ȫ���ֵ�ڶ����ֵ�ļ�����
    getExcelLineInfo
    ReDim g_arr_testCase(g_num_excel_actual_row, g_num_excel_actual_column)
    ReDim g_arr_testCase_key(g_JL_Excel.worksheets.count, g_num_excel_actual_column)
    
    '7����ʼ��ȫ���ֵ�
    Set g_dic_subItem_testCase = CreateObject("scripting.dictionary")
    
    '8����ʼ���������
    g_num_err_excel = 0
    
End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
'           2��checkErr�����Excel����Լ�¼�еĴ���д��һ��txt�ļ���Ȼ����ʾ������Ա�޸�Excel���
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub checkErr()



End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
'           3��readDataFromExcel����ȡexcel����еĲ���������Ϣ����������g_dic_subItem
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub readDataFromExcel()

    '1���Ż�ѡ��
    If C_opt Then
        ExcelApp.ScreenUpdating = False          '�ű�����Exce1ʱ����ˢ����Ļ
        Application.ScreenUpdating = False       '�ű�����wordʱ����ˢ����Ļ
    End If
    
    
    '2���������
    Dim sheetNum, rowNum, colNum As Long         '��������õ��������
    Dim caseNum As Long: caseNum = 0             '���������ܼ���
    Dim cellStr As String                        '�ݴ汣�浥Ԫ�������
    Dim subItemCode As String                    '��һ�����������Ĳ��������ʶ��
    'Dim tmp_dic_subItemCode_repeat As Object              '�ظ��Ĳ��������Ĳ��������ʶ��
    'Dim subItemCodeRepeatRow As Long             '�ڵڼ����ظ�
    'Dim caseNumInSubItem As String              'Excel����в��������ļ����������Ǵ����
    Dim realCaseNumInSubItem As Integer          '��ȷ�Ĳ��������еĲ��������ļ�
    Dim tmp_arr_caseNumInSubItem() As String     '����ַ���-���ָ��������ƺ������
    Dim tmp_numInCaseName As String              '����������ơ�-��������
    '�ֵ����
    Dim tmpKey, tmpVal As String                 '�����ʱ�ֵ�ļ���ֵ
   
    
    
    
    '3������Excel����sheet������row������column
'----------------------��һ��ѭ��------------------------------------------------------------------------------------------
    For sheetNum = 1 To g_JL_Excel.worksheets.count
        '3-1 ��ȡ��һ�е����ݣ����������g_arr_testCase_key
        'ReDim g_arr_testCase_key(g_JL_Excel.worksheets(sheetNum).usedrange.Columns.count)
        For colNum = 1 To g_arr_num_excel_colInSheets(sheetNum)
            cellStr = Escape_Text(g_JL_Excel.worksheets(sheetNum).Cells(1, colNum).Value)
            '��ֵΪ��ʱ��ʹ���������Ϊ��
            If cellStr = "" Then
                g_arr_testCase_key(sheetNum, colNum) = colNum
            Else
                g_arr_testCase_key(sheetNum, colNum) = cellStr
            End If
        Next colNum
        '�����ظ��������ʱ�ֵ�
        'Set tmp_dic_subItemCode_repeat = CreateObject("scripting.dictionary")
'----------------------�ڶ���ѭ��------------------------------------------------------------------------------------------
        For rowNum = C_excel_start_row To g_arr_num_excel_rowInSheets(sheetNum) '��ÿ��sheet��ʵ����ʼ�п�ʼ '�ӱ���������ʼ�п�ʼ���ֵ��вű�����������Ϣ

            caseNum = caseNum + 1 'ͳ������������Ϊ������������ĵ�һ���±�
            Dim tmp_dic_testCase As Object
            Set tmp_dic_testCase = CreateObject("scripting.dictionary")  '���Ҫ�ظ�����ĳһ����ʱ���ֵ�������Ǿͱ�����ÿ��ʹ��ǰsetһ�£���������ֵ�����ݣ���������
'----------------------������ѭ��------------------------------------------------------------------------------------------
            For colNum = 1 To g_arr_num_excel_colInSheets(sheetNum)
                '3-2 ��ȡ��Ԫ������
                cellStr = g_JL_Excel.worksheets(sheetNum).Cells(rowNum, colNum).Value '�洢ԭʼ�ַ������������ݲ�ͬ���н��д���
                
                '3-3 ��ȡ��ʱ�ֵ�tmp_dic_testCase�ļ�
                tmpKey = g_arr_testCase_key(sheetNum, colNum)
                
                '3-4 ���ݲ�ͬ���У����в�ͬ�Ĵ�����ȡ��Ϣ���жϴ���
                Select Case colNum
                
                    '1���������ţ����������ʶ�ţ�
                    Case E_colNum_JLExcel.subItemCode
                        tmpVal = Escape_Text(cellStr)
                        '1-1 ---------��������1---------����������Ϊ��
                        If tmpVal = "" Then
                            tmpVal = subItemCode    '�ڲ�������Ϊ�յ�����£�����ʹ����һ�����������Ĳ�������
                            realCaseNumInSubItem = realCaseNumInSubItem + 1 '�ڲ�������Ϊ�յ�����£��趨ʵ���������+1
                            recErrInfo E_errCode_JL.subItemCode_Null & "��������Ϊ��", sheetNum, rowNum, colNum        '��¼������Ϣ
                            
                        '1-2 ��ǰ������������һ�������Ĳ���������ͬ������ǰ��������һ��������ͬһ���������������
                        ElseIf cellStr = subItemCode Then
                            realCaseNumInSubItem = realCaseNumInSubItem + 1 'ʵ���������+1

                        '1-4 ��ǰ������������һ�������Ĳ������Ų�ͬ����һ���������������������ͳ���꣬������һ�����������������д����Ϣ��ȫ���ֵ�
                        Else
                            'ÿ�ν�����һ����������ʱ�������³�ʼ�������������ʱ�ֵ�
                            Dim tmp_dic_subItem As Object
                            '����ʵ���������
                            realCaseNumInSubItem = 1
                            If rowNum > C_excel_start_row Then  'Ҫ�Ƚű����е�������ʼ�е���һ�У�������ʱ�ֵ��вŻ������ݣ����ܽ���д��
                                '--------------------------------------------
                                '���������Ĳ��������ظ���Ӱ����һ������������֮ǰ�Ĳ������������������Ϣ�Ѿ�ͳ����ɣ�����д��ȫ���ֵ�
                                '--------------------------------------------
                                g_dic_subItem_testCase.Add subItemCode, tmp_dic_subItem
                            End If
                            
                            ' ---------��������2---------�����������ظ�����ǰ���������Ĳ��������ظ��ˣ���subItemCode���ϡ��ظ�����ǣ���������Ĳ���������û����ģ�����д��
                            If g_dic_subItem_testCase.Exists(tmpVal) Then
                               
                                recErrInfo E_errCode_JL.subItemCode_Repeat & "���������ظ�", sheetNum, rowNum, colNum        '��¼������Ϣ
                                '�ű�������������������Ա�����޸�
                                
'                                '���ظ��������ŵĴ���-----
'                                If tmp_dic_subItemCode_repeat.Exists(tmpVal) Then
'                                    'tmp_dic_subItemCode_repeat(tmpVal) = tmp_dic_subItemCode_repeat(tmpVal) + 1
'                                    lastRepeatRowNum = rowNum
'                                Else
'                                    tmp_dic_subItemCode_repeat.Add tmpVal, rowNum
'                                    lastRepeatRowNum = rowNum
'                                End If
'                                If rowNum = lastRepeatRowNum Then
'                                    subItemCode = tmpVal & "-�ظ�-" & tmp_dic_subItemCode_repeat(tmpVal)
'                                End If
                                subItemCode = cellStr & "-�ظ�-" & g_num_err_excel
                                
                            '���ظ��Ĳ������ţ���������ķ�֧
                            Else

                                'д��ȫ���ֵ�󣬸���subItemCodeΪ��ǰ�Ĳ������ţ������ж���һ�������Ƿ�����ͬһ����������
                                subItemCode = tmpVal
                            End If
                            
                            '1-3-2 ÿ�ν�����һ����������ʱ�������³�ʼ�������������ʱ�ֵ�
                            Set tmp_dic_subItem = CreateObject("scripting.dictionary")
                        End If
                        
                    '2���������
                    Case E_colNum_JLExcel.caseNumInSubItem
                        tmpVal = Escape_Text(cellStr)
                        ' ---------��������3---------��������Ų�������
                        If Not IsNumeric(tmpVal) Then
                            recErrInfo E_errCode_JL.caseNumInSubItem_NotNum & "������Ų�������", sheetNum, rowNum, colNum        '��¼������Ϣ
                            '�ű�������������������Ա�����޸�
                        ' ---------��������4---------�����������ʵ����Ų�һ��
                        ElseIf realCaseNumInSubItem <> CInt(tmpVal) Then
                            recErrInfo E_errCode_JL.caseNumInSubItem_Diff & "���������ʵ�ʲ�һ��", sheetNum, rowNum, colNum         '��¼������Ϣ
                            '�ű�������������������Ա�����޸�
                        End If
                        
                    '3����������
                    Case E_colNum_JLExcel.caseName
                        tmpVal = Escape_Text(cellStr)
                         ' ---------��������5---------��'��������Ϊ��
                        If tmpVal = "" Then
                            recErrInfo E_errCode_JL.caseName_Null & "��������Ϊ��", sheetNum, rowNum, colNum           '��¼������Ϣ
                            '�ű�������������������Ա�����޸�
                        Else
                            tmp_arr_caseNumInSubItem = Split(tmpVal, "-")
                            tmp_numInCaseName = tmp_arr_caseNumInSubItem(UBound(tmp_arr_caseNumInSubItem))  '��ȡ���������-����������
                            ' ---------��������5---------���������ƺ���û�мӡ�-�������������ƺ�û�����
'                            If UBound(tmp_arr_caseNumInSubItem) = -1 Then
                            If InStr(tmpVal, "-") = 0 Then
                                recErrInfo E_errCode_JL.caseName_SeqNull & "�������Ƶ����Ϊ��", sheetNum, rowNum, colNum          '��¼������Ϣ
                                '��������
                                tmpVal = tmpVal & "-" & realCaseNumInSubItem
                            ' ---------��������6---------���������ơ�-������Ĳ�������
                            ElseIf Not IsNumeric(tmp_numInCaseName) Then
                                recErrInfo E_errCode_JL.caseName_SeqNotNum & "�������Ƶ���Ų�������", sheetNum, rowNum, colNum           '��¼������Ϣ
                                tmpVal = getSubItemName(tmpVal) & "-" & realCaseNumInSubItem '�ű�����
                                
                            ' ---------��������7---------���������ơ�-����������ʵ�ʲ�һ��
                            ElseIf realCaseNumInSubItem <> CInt(tmp_numInCaseName) Then
                                recErrInfo E_errCode_JL.caseName_SeqDiff & "�������Ƶ������ʵ�ʲ�һ��", sheetNum, rowNum, colNum            '��¼������Ϣ
                                tmpVal = getSubItemName(tmpVal) & "-" & realCaseNumInSubItem '�ű�����
                                
                            End If
                        End If
                        
                    '4���Ƿ�ͨ��
                    Case E_colNum_JLExcel.casePass
                        tmpVal = Escape_Text(cellStr)
                         ' ---------��������8---------�����Խ���Ϊ��
                        If tmpVal = "" Then
                            recErrInfo E_errCode_JL.casePass_Null, sheetNum & "���������Ľ���Ϊ��", rowNum, colNum             '��¼������Ϣ
                            '�ű�������������������Ա�����޸�
                         ' ---------��������9---------�����������Ľ��۱�����"ͨ��"��"δͨ��"��"����Ľ�"
                        ElseIf tmpVal <> "ͨ��" And tmpVal <> "δͨ��" And tmpVal <> "��ͨ��" And tmpVal <> "����Ľ�" Then
                            recErrInfo E_errCode_JL.casePass_Wrong & "���������Ľ��۲���""ͨ��""��""δͨ��""��""��ͨ��""��""����Ľ�""", sheetNum, rowNum, colNum             '��¼������Ϣ
                            '�ű�������������������Ա�����޸�
                        End If
'
                    '5����ȥ�����4�У������в��жϴ���ֱ��ʹ��ԭʼ����
                    Case Else
                        tmpVal = cellStr
                        
                End Select
                '--------------------------------------------
                '����������һ����Ϣ�Ѵ�������д����ʱ�ֵ�tmp_dic_testCase
                '--------------------------------------------
                tmp_dic_testCase.Add tmpKey, tmpVal
            Next colNum
'----------------------������ѭ��------------------------------------------------------------------------------------------

        '--------------------------------------------
        'һ������������������Ϣ�Ѿ�д����ʱ�ֵ�tmp_dic_testCase����tmp_dic_testCaseд�����һ�����ʱ�ֵ�tmp_dic_testCase
        '--------------------------------------------
        'Call printDic1(tmp_dic_testCase)
        tmp_dic_subItem.Add realCaseNumInSubItem, tmp_dic_testCase
        '���ű��Ѿ�������ÿ��sheet�����һ��ʱ����ʱ�ű��͵��������һ������������ֱ�Ӱѵ�ǰ����ʱ�ֵ�tmp_dic_subItem����ȫ���ֵ�
'        If sheetNum = g_JL_Excel.worksheets.count And rowNum = g_arr_num_excel_rowInSheets(g_JL_Excel.worksheets.count) Then
        If rowNum = g_arr_num_excel_rowInSheets(g_JL_Excel.worksheets.count) Then
            '--------------------------------------------
            '���������������������һ��������ֱ�ӽ���ʱ�ֵ�tmp_dic_subItemд��ȫ���ֵ�
            '--------------------------------------------
            g_dic_subItem_testCase.Add subItemCode, tmp_dic_subItem
        End If
        Next rowNum
        'Call printDic2(tmp_dic_subItem)
'----------------------�ڶ���ѭ��------------------------------------------------------------------------------------------
    Next sheetNum
'----------------------��һ��ѭ��------------------------------------------------------------------------------------------
    'Call printDic3(g_dic_subItem_testCase)
    '��������ֵ��Ƿ������Ԥ�ڵ�����
    
End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
'           4��writeDataToWord����дword��Ĳ��Լ�¼
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub writeDataToWord()

'    init
'    readDataFromExcel

    ' ---------��������---------�����Լ�¼��Excel(�ű�ͳ��-ȫ���ֵ�)��Word(�������)�����汾�еĲ���������������ͬ
    If ActiveDocument.Tables.count <> g_dic_subItem_testCase.count Then
        recErrInfo E_errCode_JL.subItemNum_Diff & "���Լ�¼��Excel(�ű�ͳ��-ȫ���ֵ�)��Word(�������)�����汾�еĲ���������������ͬ"         '��¼������Ϣ
    End If
    
    '˼·��
    '1������ʵ�ʲ������ݣ���Excel������ݣ�ȥ��дWord
    '2�����ܻ���������������⣺һ��Excel�ж�Word�ޣ�������Excel�Ĳ�������д��������Word��Excel�ޣ�Excel©��©д��
    '����ʵ�֣�
    '1������ȫ���ֵ䣬��Word�ı�����Դ���������������ʶ��1���������������ʶһ�£��������д��2��������ΪExcel�Ĳ�������д������д���������������������¼������Ϣ��������Ա�ֶ���д
    '2��������ȫ���ֵ���ٱ���Word��ȫ��������������뼰����˵������һ���µĵ�Ԫ��1�����ȫ�������ݣ�˵��Excelû��©��©д��2�������пհ׵Ļ�������ΪExcel����©��©д����¼������Ϣ
    
    'Word����������ر���
    Dim tbl As table '��ǰ���������Word���
    Dim tblNum As Long 'Word������
    'Dim cellStr As String  'Word���ĵ�Ԫ������
    
    '��¼����
    Dim dic_errSubItem_Excel As Object: Set dic_errSubItem_Excel = CreateObject("scripting.dictionary")     '���Excel��û��д��Word�Ĳ�������������µĶ��������������Excel�Ĳ��������ʶ��д��
    Dim dic_errSubItem_Word As Object: Set dic_errSubItem_Word = CreateObject("scripting.dictionary")       '���Word�е���Excelû�еĲ������Excel����Լ�¼©��©д
    
    '
    Dim arr_subItemCode() As Variant: arr_subItemCode = g_dic_subItem_testCase.keys       '���еĲ��������ʶ��
    
'    Dim tmp_arr_caseName() As String  '��ŷָ�����������Ƶ�����
    Dim excelCaseName As String 'Excel���������еĲ�����������
    
    'Word�������
    Dim caseName As Range  '������������
    Dim caseCode As Range  '����������ʶ
    Dim caseInit As Range  '����������ʼ��
    Dim premise As Range  'ǰ����Լ��
    'Dim caseStart As Range  '��������ʵ������
    Dim caseInput As Range  '����
    Dim caseExpectedOutcome As Range    '�������
    Dim caseCert As Range       '����׼��
    Dim caseActualResult As Range   'ʵ�ʲ��Խ��
    Dim isPass As Range             'ִ�н��
    Dim designer As Range  '�����Ա
    Dim dsgnDate As Range   '�������
    Dim execCond As Range   'ִ�����
    Dim tester As Range  '������Ա
    Dim supervisor As Range '�ල��Ա
    Dim execDate As Range   'ִ������
    
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
        
        '�������������ʶһ��("��������"))
        If arr_subItemCode(tblNum - 1) & "-" = Escape_Text(caseCode.Text) Then
        
            Dim dic_subCodeItem As Object: Set dic_subCodeItem = CreateObject("scripting.dictionary")
            Dim dic_testCase As Object: Set dic_testCase = CreateObject("scripting.dictionary")
            Dim caseNumInSubItem As Integer
            
            '��ȡ���������µ����в��������ֵ�
            Set dic_subCodeItem = g_dic_subItem_testCase(arr_subItemCode(tblNum - 1))
            '��ȡkeys
            Dim arr_caseNumInSubItem() As Variant: arr_caseNumInSubItem = dic_subCodeItem.keys

            '�������������µ����в�������
            For caseNumInSubItem = 1 To dic_subCodeItem.count
                Set dic_testCase = dic_subCodeItem(arr_caseNumInSubItem(caseNumInSubItem - 1))
                Dim colName() As Variant: colName = dic_testCase.keys
                '������������Ĳ���������Word���ı��һֱ
                If dic_testCase(colName(subItemCode - 1)) & "-" = Escape_Text(caseCode.Text) Then
                    Debug.Print dic_testCase(colName(subItemCode - 1))
                    '���ݲ�ͬ���У���д���������Word�ı����
                End If
            Next caseNumInSubItem
        End If
        
    
    Next tblNum
    
End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
'           5��saveJL���رպͱ�������ļ������word��Ĳ��Լ�¼
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub saveJL()
    '1�����word��Ĳ��Լ�¼�����ر�ԭ�ļ��ղ�����
    g_JL_Word.SaveAs (workPath & "\" & C_word_fileName_saved)
    g_JL_Word.Close wdDoNotSaveChanges
    '2���ر�exce1��Ĳ��Լ�¼���Ҳ�����
    g_JL_Excel.Close wdDoNotSaveChanges
    '3���˳�������APP���ͷ�APP�ڴ�
    'WordApp.Quit
    ExcelApp.Quit
    Set g_JL_Word = Nothing
    Set g_JL_Excel = Nothing
End Sub


'-------------------------------------------------------------------------------------------
'     1������getExcelLineInfo
'-------------------------------------------------------------------------------------------
Private Sub getExcelLineInfo()

    Dim sheetNum, rowCount, colCountOne, colCountAll, tmp As Long
      
    sheetNum = g_JL_Excel.worksheets.count
    ReDim g_arr_num_excel_rowInSheets(sheetNum)    '����ʵ�ʵ�sheetֵ�ض���������
    ReDim g_arr_num_excel_colInSheets(sheetNum)     '����ʵ�ʵ�sheetֵ�ض���������

    '��ȡÿ��sheet������������sheet���������ܺ�
    For tmp = 1 To sheetNum
        g_arr_num_excel_rowInSheets(tmp) = g_JL_Excel.worksheets(tmp).usedrange.Rows.count '���ȫ�ֱ�������
        rowCount = rowCount + g_JL_Excel.worksheets(tmp).usedrange.Rows.count
    Next tmp
    
    g_num_excel_actual_row = rowCount 'ȫ�ֱ�������ȡ���Լ�¼����sheetʹ�õ�����
    
    
    '��ȡÿ��sheet������������sheet�����������ֵ
    colCountOne = g_JL_Excel.worksheets(1).usedrange.Columns.count
    colCountAll = g_JL_Excel.worksheets(1).usedrange.Columns.count
    For tmp = 1 To sheetNum
        'Ŀǰ��Ϊsheet����������С��8��E_colNum_JLExcel.casePass�������С��8������Ϊ8
        If g_JL_Excel.worksheets(tmp).usedrange.Columns.count < E_colNum_JLExcel.casePass Then
            colCountOne = E_colNum_JLExcel.casePass
        Else
            colCountOne = g_JL_Excel.worksheets(tmp).usedrange.Columns.count
        End If
        g_arr_num_excel_colInSheets(tmp) = colCountOne   '����-��ȡÿ��sheet������
        
        If g_JL_Excel.worksheets(tmp).usedrange.Columns.count > colCountAll Then
            colCountAll = g_JL_Excel.worksheets(tmp).usedrange.Columns.count
        End If
    Next tmp
    
    g_num_excel_actual_column = colCountAll 'ȫ�ֱ�������ȡ���Լ�¼����sheet���������
    
End Sub
'-------------------------------------------------------------------------------------------
'     2������Escape_Text
'-------------------------------------------------------------------------------------------
Private Function Escape_Text(ByVal originStr As String)
    Dim tmp As String
    tmp = Replace(originStr, Chr(13), "")   '���з�
    tmp = Replace(tmp, Chr(10), "")         '���з�
    tmp = Replace(tmp, Chr(7), "")          'word����е�С�ڵ�
    tmp = Replace(tmp, Chr(9), "")          'Tab-ˮƽ�Ʊ��
    tmp = Replace(tmp, Chr(11), "")         '��ֱ�Ʊ��
    tmp = Replace(tmp, " ", "") '           ��Chr(32) = һ���ո��
    Escape_Text = tmp
End Function
'-------------------------------------------------------------------------------------------
'     3������recErrInfo
'-------------------------------------------------------------------------------------------
Private Sub recErrInfo(ByVal errMsg As String, Optional ByVal sheetNum As String = "", Optional ByVal rowNum As String = "", Optional ByVal colNum As String = "")
    g_num_err_excel = g_num_err_excel + 1
    ReDim Preserve g_arr_err_excel(g_num_err_excel)
    If sheetNum <> "" Then
        g_arr_err_excel(g_num_err_excel) = g_num_err_excel & " --" & errMsg & "--λ�ã�sheet=" & sheetNum & " ��=" & rowNum & " ��=" & colNum
    Else
        g_arr_err_excel(g_num_err_excel) = g_num_err_excel & " --" & errMsg
    End If
    
End Sub
'-------------------------------------------------------------------------------------------
'     4������getSubItemName���������������ҵ������������ƣ�����ӡ���ʼ������-01���еõ�����ʼ�����ܡ�
'-------------------------------------------------------------------------------------------
Private Function getSubItemName(ByVal str As String)
    
    Dim charPos As Integer
    charPos = InStrRev(str, "-")                '����InStr()�ͺ���InStrRev()��������ͬ�������෴��ǰ�ߴ������Ҳ���ָ���ַ����ֵ�λ�ã����ߴ����������
    getSubItemName = Mid(str, 1, charPos - 1)
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ӡ�ֵ䣬printDic
'���������ô��ݣ����ڽϴ���ֵ���Խ�ʡʱ��
'���õ�ʱ����Ҫ�ں���ǰ����� ��Call��
'input���ֵ�
'output����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function printDic1(ByRef dic As Object)
    'Debug.Print "printDic()"
    Dim count As Integer, keys, items
    
    keys = dic.keys
    items = dic.items
    
    Debug.Print "һά�ֵ乲��" & dic.count & "��Ԫ��"
    For count = 1 To dic.count
        Debug.Print "No." & count & ": " & "(" & keys(count - 1) & ", " & items(count - 1) & ")"
        'Debug.Print "��" & count & "��Ԫ�����ݣ�" & items(count - 1) & Chr(13)
    Next count
    Debug.Print "��ӡ��ɡ�" & Chr(13)
End Function
Private Function printDic2(ByRef dic As Object)
    'Debug.Print "printDic()"
    Dim count As Integer, keys, items
    Dim dic2 As Object
    
    Set dic2 = CreateObject("scripting.dictionary")
    
    keys = dic.keys
    items = dic.items
    
    Debug.Print "��ά�ֵ乲��" & dic.count & "��Ԫ��" & Chr(13)
    For count = 1 To dic.count
        Set dic2 = CreateObject("scripting.dictionary")
        Set dic2 = dic(keys(count - 1))
        Debug.Print "��ά�ֵ�ĵ�" & count & "�������������Ϊ��" & keys(count - 1)
        '�ڶ����ֵ伴dic2
        Call printDic1(dic2)
    Next count
    
    Debug.Print "��ӡ��ɡ�" & Chr(13)
End Function
Private Function printDic3(ByRef dic As Object)
    Dim count As Integer, keys, items
    Dim dic2 As Object
    
    Set dic2 = CreateObject("scripting.dictionary")
    
    keys = dic.keys
    items = dic.items
    
    Debug.Print "��ά�ֵ乲��" & dic.count & "��Ԫ��" & Chr(13)
    For count = 1 To dic.count
        Set dic2 = CreateObject("scripting.dictionary")
        Set dic2 = dic(keys(count - 1))
        Debug.Print "��ά�ֵ�ĵ�" & count & "�����������ʶ��Ϊ��" & keys(count - 1) & Chr(13)
        '�ڶ����ֵ伴dic2
        Call printDic2(dic2)
    Next count
    
    Debug.Print "��ӡ��ɡ�"
End Function

Private Sub test()
    getSubItemName ("01-��ʼ��-����-01")
End Sub

    


