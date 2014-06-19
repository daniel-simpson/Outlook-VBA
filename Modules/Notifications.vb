
Private Sub Send_Notification(ByVal senderStr As String, ByVal subjectStr As String)
    Dim httpReq As Object: Set httpReq = CreateObject("MSXML2.XMLHTTP")
    Dim result, Url, timeStr As String: timeStr = time
    
    'Note, need to add API key on next line
    Url = "https://www.notifymyandroid.com/publicapi/notify?apikey=__MYAPIKEY__" _
            + "&application=" + URLEncode("Outlook (New Mail)") _
            + "&event=" + URLEncode("From: " + senderStr + " (" + timeStr + ")") _
            + "&description=" + URLEncode(subjectStr)
        
    httpReq.Open "GET", Url, False
    httpReq.Send
     
    result = httpReq.responseText
    'TODO: check for 200 success code
    Set httpReq = Nothing
End Sub

Private Sub Test_Notifications()
    Send_Notification "From: TEST SENDER", "DESC"
End Sub

