Attribute VB_Name = "CloseApp"
Option Explicit

Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long

Private Declare Function FindWindow Lib "user32" _
   Alias "FindWindowA" _
   (ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long

Private Declare Function PostMessage Lib "user32" _
   Alias "PostMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Private Declare Function IsWindow Lib "user32" _
   (ByVal hWnd As Long) As Long

'Constants used by the API functions
Const WM_CLOSE = &H10
Const INFINITE = &HFFFFFFFF

Private Sub Form_Load()
   Command1.Caption = "Start the Calculator"
   Command2.Caption = "Close the Calculator"
End Sub

Private Sub Command1_Click()
'Starts the Windows Calculator
   Shell "calc.exe", vbNormalNoFocus
End Sub

Private Sub Command2_Click()
'Closes the Windows Calculator
   Dim hWindow As Long
   Dim lngResult As Long
   Dim lngReturnValue As Long

   hWindow = FindWindow(vbNullString, "Calculator")
   lngReturnValue = PostMessage(hWindow, WM_CLOSE, vbNull, vbNull)
   lngResult = WaitForSingleObject(hWindow, INFINITE)

   'Does the handle still exist?
   DoEvents
   hWindow = FindWindow(vbNullString, "Calculator")
   If IsWindow(hWindow) = 1 Then
      'The handle still exists. Use the TerminateProcess function
      'to close all related processes to this handle. See the
      'article for more information.
      MsgBox "Handle still exists."
   Else
      'Handle does not exist.
      MsgBox "Program closed."
   End If
End Sub
