Private WithEvents Items As Outlook.Items


Private Sub Application_Startup()
  Dim olApp As Outlook.Application : Set olApp = Outlook.Application
  Dim objNS As Outlook.NameSpace : Set objNS = olApp.GetNamespace("MAPI")
    
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub


Private Sub Items_ItemAdd(ByVal item As Object)
  On Error GoTo ErrorHandler
  
  If TypeName(item) = "MailItem" And Check_If_Locked = "Locked" Then
    Dim msg As Outlook.MailItem: Set msg = item
    Send_Notification "From: " + msg.Sender, msg.Subject
  End If

ProgramExit:
  Exit Sub
  
ErrorHandler:
  MsgBox Err.Number & " - " & Err.description
  Resume ProgramExit
End Sub


