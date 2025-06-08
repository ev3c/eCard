Attribute VB_Name = "HotKey"
Public Const GWL_WNDPROC = (-4)
Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
 
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
 
Public Proc_Original As Long
 
Public Sub Subclasifica_Ventana(hWnd As Long)
  If Proc_Original = 0 Then
    Proc_Original = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf Mi_Proc)
  End If
End Sub
 
Public Sub Ventana_Normal(hWnd As Long)
  If Proc_Original <> 0 Then
     Call SetWindowLong(hWnd, GWL_WNDPROC, Proc_Original)
     Proc_Original = 0
  End If
End Sub
 
 
Public Function Mi_Proc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
On Error GoTo errHandler
  
  If uMsg = WM_HOTKEY Then
  
    frmPrograma.Show
    
    If frmPrograma.Contraseña_Entrar Then
      frmPrograma.Show
      Call frmPrograma.Campos_Ver
      Call frmPrograma.Command_Mostrar
      Call frmPrograma.Evento_Ver
'      frmPrograma.tmrPrograma.Interval = 0
    Else
      frmPrograma.Hide
      App.TaskVisible = False
    End If
  
  Else
    Mi_Proc = CallWindowProc(Proc_Original, hw, uMsg, wParam, lParam)
  End If

  Exit Function

errHandler:
  strMsg = "Había una subventana abierta"
  MsgBox strMsg, vbInformation, gstrPrograma

End Function

