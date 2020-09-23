Attribute VB_Name = "Modol"
'Display browse for folder
Public Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
  ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" _
   (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'--//

'Always on top
Private Declare Function SetWindowPos Lib "user32.dll" _
(ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, _
ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const hwnd_TopMost = -1
Public Const hwnd_noTopMost = -2
Public Const hwnd_Top = 0
Public Const swp_noSize = &H1
Public Const swp_noMove = &H2

Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'/

'App Prev Instance
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function OpenIcon Lib "User32" (ByVal hWnd As Long) As Long
Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Public Const GW_HWNDPREV = 3
'/

'Always on top
Public Sub FormTopMost(hWnd As Long)
SetWindowPos hWnd, hwnd_TopMost, 0, 0, 0, 0, swp_noSize + swp_noMove
End Sub

Public Sub FormNoTopMost(hWnd As Long)
SetWindowPos hWnd, hwnd_noTopMost, 0, 0, 0, 0, swp_noMove + swp_noSize
End Sub
'/

Public Function FileExists(strPath As String) As Integer
On Error Resume Next
Dim lngRetVal As Long
lngRetVal = Len(Dir$(strPath))
If err Or lngRetVal = 0 Then FileExists = False Else FileExists = True
End Function

Public Function StripPath(T$) As String
Dim err As ErrObject
Dim ERROR_CHARS As String: ERROR_CHARS = "/" & Chr(34)
Dim X%, ct%, L
  StripPath$ = T$
  X% = InStr(T$, "\")
  Do While X%
     ct% = X%
     X% = InStr(ct% + 1, T$, "\")
  Loop
  If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
  For i = 1 To Len(StripPath$)
    If InStr(1, ERROR_CHARS, Mid(StripPath$, i, 1), vbTextCompare) Then
      'If err <> "" Then
        L = StripPath$
        StripPath$ = L
        X% = InStr(L, "/")
        Do While X%
          ct% = X%
          X% = InStr(ct% + 1, L, "/")
        Loop
        If ct% > 0 Then StripPath$ = Mid$(L, ct% + 1)
        Exit Function
      'End If
    End If
  Next
End Function

Public Sub ShowPrevInstance()
Dim OldTitle As String
Dim ll_WindowHandle As Long

OldTitle = App.Title
App.Title = "abcde - Aplikasi ini akan ditutup!"
ll_WindowHandle = FindWindow("ThunderRT6Main", OldTitle)
If ll_WindowHandle = 0 Then Exit Sub
ll_WindowHandle = GetWindow(ll_WindowHandle, GW_HWNDPREV)
Call OpenIcon(ll_WindowHandle)
Call SetForegroundWindow(ll_WindowHandle)
End
End Sub
