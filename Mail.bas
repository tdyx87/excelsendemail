Attribute VB_Name = "Mail"

Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "U\n14"

sendMail (2)

End Sub

Sub sendMail(startrow As Integer)
On Error GoTo errHandle
Dim outlookApp As Outlook.Application
Dim outlookItem As Outlook.MailItem
Set outlookApp = New Outlook.Application
Dim rows As Integer
Dim columns As Integer
rows = ActiveSheet.UsedRange.rows.Count
columns = ActiveSheet.UsedRange.columns.Count
Dim row As Integer
Dim column As Integer
Dim tableHead As String
tableHead = createTableHead(startrow)

For row = startrow + 1 To startrow + rows - 1
Set outlookItem = outlookApp.CreateItem(olMailItem)

Dim email As String
Dim htmlbody As String
Dim subject As String
email = Cells(row, 1)
subject = CStr(Cells(row, 2)) + "工资单"
htmlbody = tableHead + createTableContent(row)
With outlookItem
.To = email
.subject = subject
.htmlbody = htmlbody
.Send
End With
Set outlookItem = Nothing
Next row
Set outlookApp = Nothing
MsgBox "发送完成"
Exit Sub
errHandle:
    MsgBox Err.Description
Exit Sub
End Sub

Function createTableHead(headrow As Integer) As String

Dim tableHead As String
Dim tableStyle As String

tableHead = "<table style=""background:#FFF""><tbody>"
tableHead = tableHead + "<tr>"
'标题从第二列开始
For column = 2 To ActiveSheet.UsedRange.columns.Count
tableHead = tableHead + createtdContentWithBlackBorder(Cells(headrow, column))
Next column
tableHead = tableHead + "</tr>"

createTableHead = tableStyle + tableHead
End Function

Function createTableContent(contentrow As Integer) As String
Dim tableContent As String
tableContent = "<tr>"
For column = 2 To ActiveSheet.UsedRange.columns.Count
tableContent = tableContent + createtdContentWithBlackBorder(Cells(contentrow, column))
Next column
tableContent = tableContent + "</tr></tbody></table>"

createTableContent = tableContent

End Function

Function createtdContentWithBlackBorder(content As Variant) As String

createtdContentWithBlackBorder = createtdContent(content, "background:#FFF")

End Function


Function createtdContent(content As Variant, style As String) As String

createtdContent = "<td nowrap=""nowrap"" style=" + style + ">" + CStr(content) + "</td>"

End Function
