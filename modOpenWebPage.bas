Attribute VB_Name = "modOpenWebPage"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Public Sub LoadWebPage(psWebPage As String, psForm As Form)
    If Len(psWebPage) <> 0 Then
        ShellExecute psForm.hwnd, "open", psWebPage, "", "", SW_SHOW
    Else
        Exit Sub
    End If
End Sub
