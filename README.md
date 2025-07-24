### Hi there 👋
Sub ExportRecentEmails()
    Dim olApp As Object
    Dim olNs As Object
    Dim olFolder As Object
    Dim olItems As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Dim filterDate As Date
    Dim strFilter As String
    Dim i As Long
    Dim lastRow As Long
    
    ' 设置日期范围（当前时间 - 24小时）
    filterDate = Now - 1
    
    ' 创建Outlook应用对象
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number <> 0 Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    If olApp Is Nothing Then
        MsgBox "无法启动Outlook.", vbCritical
        Exit Sub
    End If
    
    ' 获取收件箱
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFolder = olNs.GetDefaultFolder(6) ' 6 = 收件箱
    Set olItems = olFolder.Items
    
    ' 设置时间过滤器
    strFilter = "[ReceivedTime] >= '" & Format(filterDate, "ddddd h:nn AMPM") & "'"
    
    ' 应用过滤器并排序（最新邮件在前）
    olItems.Sort "[ReceivedTime]", True
    Set olItems = olItems.Restrict(strFilter)
    
    ' 设置工作表
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear
    With ws
        .Range("A1:D1") = Array("发件人", "主题", "接收时间", "正文预览")
        .Rows(1).Font.Bold = True
    End With
    
    ' 导出邮件
    i = 2
    For Each olMail In olItems
        If TypeName(olMail) = "MailItem" Then
            ws.Cells(i, 1) = olMail.SenderName
            ws.Cells(i, 2) = olMail.Subject
            ws.Cells(i, 3) = olMail.ReceivedTime
            ws.Cells(i, 4) = Left(olMail.Body, 100) ' 截取前100个字符
            i = i + 1
        End If
    Next olMail
    
    ' 自动调整列宽
    ws.Columns("A:D").AutoFit
    
    ' 添加时间戳
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    ws.Cells(lastRow, 1) = "导出时间："
    ws.Cells(lastRow, 2) = Now
    
    ' 清理对象
    Set olMail = Nothing
    Set olItems = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    
    MsgBox "成功导出 " & i - 2 & " 封邮件！", vbInformation
End Sub
<!--
**xiaoka00xiaoy/xiaoka00xiaoy** is a ✨ _special_ ✨ repository because its `README.md` (this file) appears on your GitHub profile.

Here are some ideas to get you started:

- 🔭 I’m currently working on ...
- 🌱 I’m currently learning ...
- 👯 I’m looking to collaborate on ...
- 🤔 I’m looking for help with ...
- 💬 Ask me about ...
- 📫 How to reach me: ...
- 😄 Pronouns: ...
- ⚡ Fun fact: ...
-->
