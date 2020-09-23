<div align="center">

## Creating email and attaching a file in lotus notes


</div>

### Description

This code is to show how to create an email in lotus notes with an attached file
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stan Allan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stan-allan.md)
**Level**          |Unknown
**User Rating**    |4.6 (46 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stan-allan-creating-email-and-attaching-a-file-in-lotus-notes__1-3256/archive/master.zip)





### Source Code

```
Private Sub CmdSend_Click()
Dim oSess As Object
Dim oDB As Object
Dim oDoc As Object
Dim oItem As Object
Dim direct As Object
Dim Var As Variant
Dim flag As Boolean
Form1.MousePointer = 11
Form1.StatusBar1.SimpleText = "Opening Lotus Notes..."
Set oSess = CreateObject("Notes.NotesSession")
Set oDB = oSess.GETDATABASE("", "")
Call oDB.OPENMAIL
flag = True
If Not (oDB.ISOPEN) Then flag = oDB.OPEN("", "")
If Not flag Then
MsgBox "Can't open mail file: " & oDB.SERVER & " " & oDB.FILEPATH
GoTo exit_SendAttachment
End If
On Error GoTo err_handler
Form1.StatusBar1.SimpleText = "Building Message"
Set oDoc = oDB.CREATEDOCUMENT
Set oItem = oDoc.CREATERICHTEXTITEM("BODY")
oDoc.Form = "Memo"
oDoc.subject = Form1.TxtSubject.Text
oDoc.sendto = Form1.TxtSendTo.Text
oDoc.body = Form1.TxtMessage.Text
oDoc.postdate = Date
Form1.StatusBar1.SimpleText = "Attaching Database " & Form1.TxtFilePath
Call oItem.EMBEDOBJECT(1454, "", Form1.TxtFilePath)
oDoc.visable = True
Form1.StatusBar1.SimpleText = "Sending Message"
oDoc.SEND False
exit_SendAttachment:
On Error Resume Next
Set oSess = Nothing
Set oDB = Nothing
Set oDoc = Nothing
Set oItem = Nothing
Form1.StatusBar1.SimpleText = "Done!"
Form1.MousePointer = 1
Exit Sub
err_handler:
If Err.Number = 7225 Then
MsgBox "File doesn't exist"
Else
MsgBox Err.Number & " " & Err.Description
End If
On Error GoTo exit_SendAttachment
Form1.MousePointer = 1
End Sub
```

