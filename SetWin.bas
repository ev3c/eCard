Attribute VB_Name = "Win"
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1


Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
       hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
       ByVal cy As Long, ByVal wFlags As Long)


