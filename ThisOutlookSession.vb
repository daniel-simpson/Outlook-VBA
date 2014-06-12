Private WithEvents Items As Outlook.Items
Private Sub Application_Startup()
  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  ' default local Inbox
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub
Private Sub Items_ItemAdd(ByVal item As Object)

  On Error GoTo ErrorHandler
  Dim msg As Outlook.MailItem
  If TypeName(item) = "MailItem" Then
    Set msg = item
        
    Dim txt As String
    Send_Notification "From: " + msg.Sender, msg.Subject
  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & " - " & Err.description
  Resume ProgramExit
End Sub

Private Sub Test_Notifications()
    Dim msg As String
    Dim timeStr As String
    timeStr = time
    
    Send_Notification "From: SENDER _TimeAndEn#oding test2", "DESC"
End Sub

Private Sub Send_Notification(ByVal eventStr As String, ByVal descriptionStr As String)
    Dim result As String
    Dim httpReq As Object: Set httpReq = CreateObject("MSXML2.XMLHTTP")
    Dim timeStr As String: timeStr = time
    
    Dim url As String
    
    'Note, need to add API key on next line
    url = "https://www.notifymyandroid.com/publicapi/notify?apikey=API KEY HERE" _
            + "&application=" + URLEncode("Outlook (New Mail)") _
            + "&event=" + URLEncode(eventStr + " (" + timeStr + ")") _
            + "&description=" + URLEncode(descriptionStr)
        
    httpReq.Open "GET", url, False
    httpReq.Send
     
    result = httpReq.responseText
    Set httpReq = Nothing
End Sub

Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function
