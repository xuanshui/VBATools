' option base 1
option explicit

'使用方法：将被污染的百度网盘的链接放进来就行了，最好是https://pan.baidu.com/s/后面的，这个程序能保证99.99%的成功率。
'如果包含前面的部分（https://pan.baidu.com/s/这部分），也可以，但是成功率可能只有90%甚至80%

'测试字符串：     1cLYoq5删除b-u我[wuyongzifu ]yL-uQap啊啊pZ2vkQ 
'测试字符串：     pan.baidu.com/s/1cLYoq5删除b-u我[wuyongzifu ]yL-uQap啊啊pZ2vkQ 
'测试字符串：     htt(shanchu)ps://pan.（删）baidu.com/s/1cLYoq5删除b-u我[wuyongzifu ]yL-uQap啊啊pZ2vkQ 

main

Public Sub main()
    Dim parseStr
    
    '1、获取输入，并处理输入字符串中的杂乱字符,拼接出用户需要的字符串
    parseStr = parseInputStr()

    '2、输出最终的字符串
    If parseStr <> "" Then
        outputStr (parseStr)
    End If
End Sub

Private Function parseInputStr()
    Dim str         '获取用户输入
    Dim set_str     '存放正则表达式匹配到的结果，是一个集合类型
    
    Dim re          '正则表达式
    Set re = New RegExp
    
    re.Global = True
    re.IgnoreCase = False
    
    '1、用户取消输入/空字符串：直接退出
    str = InputBox("Input Baidu netdisk link...")

    If str = "" Then            '用户未输入，则直接返回空字符串
        parseInputStr = ""
        Exit Function
    End If
    
    '2、去除杂乱字符
    str = Replace(str, " ", "")         '去掉空格
    str = Replace(str, Chr(13), "")     '去掉换行符
    str = Replace(str, Chr(10), "")
    str = Replace(str, Chr(7), "")
    
    re.Pattern = "\[[0-9a-zA-Z\-\_\:\/\.]*\]+"    '去掉[love]这样的，在方括号中的字符，一般是表情
    str = re.Replace(str, "")

    re.Pattern = "\([0-9a-zA-Z\-\_\:\/\.]*\)+"    '去掉(love)这样的，在英文括号中的字符，一般是表情
    str = re.Replace(str, "")
    
    re.Pattern = "\（[0-9a-zA-Z\-\_\:\/\.]*\）+"    '去掉（love）这样的，在中文括号中的字符，一般是表情
    str = re.Replace(str, "")
    
    '3、获取有效字符
    re.Pattern = "[0-9a-zA-Z\-\_\:\/\.]+"         '只获取数字、英文字母、“-”、“_”、“://”、“.”，由于原始字符串中可能有中文，所以匹配结果是多个集合
    Set set_str = re.Execute(str)
    
    '4、遍历匹配到的结果，拼接字符串
    Dim count
    str = ""
    For count = 1 To set_str.count
        str = str & set_str.Item(count - 1).Value
    Next
    
    '5、如果是http://开始就不再进行拼接
    If InStr(LCase(str), "http://") = 1 Or InStr(LCase(str), "https://") = 1 Then
    
    '如果是以pan.baidu.com开始
    ElseIf InStr(LCase(str), "pan.baidu.com") = 1 Then
        str = "https://" & str
    '不是以https://开头，拼接上头部
    Else
        str = "https://pan.baidu.com/s/" & str
    End If
    '2、处理pwd后面的密码部分


    parseInputStr = str
End Function

Private Function outputStr(resultStr)
    Dim objShell
    Set objShell = CreateObject("wscript.shell")
    objShell.Run (resultStr)    '直接打开传入的链接
End Function
