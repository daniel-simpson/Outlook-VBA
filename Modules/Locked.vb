Private Declare Function SwitchDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Declare Function OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CloseDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Const DESKTOP_SWITCHDESKTOP As Long = &H100
 
 Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)
 
' http://www.mrexcel.com/forum/microsoft-access/646623-check-if-system-locked-unlocked-using-visual-basic-applications.html
 
Function Check_If_Locked() As String
    Dim p_lngHwnd As Long
    Dim p_lngRtn As Long
    Dim p_lngErr As Long
    p_lngHwnd = OpenDesktop(lpszDesktop:="Default", dwFlags:=0, fInherit:=False, dwDesiredAccess:=DESKTOP_SWITCHDESKTOP)
    If p_lngHwnd = 0 Then
        system = "Error"
    Else
        p_lngRtn = SwitchDesktop(hDesktop:=p_lngHwnd)
        p_lngErr = Err.LastDllError
         
        If p_lngRtn = 0 Then
            If p_lngErr = 0 Then
                system = "Locked"
            Else
                system = "Error"
            End If
        Else
            system = "Unlocked"
        End If
        p_lngHwnd = CloseDesktop(p_lngHwnd)
    End If
    Check_If_Locked = system
End Function

Private Sub Form_Timer()
    Sleep 2000
    MsgBox Check_If_Locked
End Sub
