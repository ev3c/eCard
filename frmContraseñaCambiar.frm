VERSION 5.00
Begin VB.Form frmContraseñaCambiar 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtNuevaContraseña 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtContraseña 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtNuevaContraseña2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   3975
      Begin VB.Label lblNuevaContraseña2 
         Caption         =   "Reentrar Nueva Contraseña"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblNuevaContraseña 
         Caption         =   "Nueva Contraseña"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label lblContraseña 
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmContraseñaCambiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  If txtContraseña.Text = gstrContraseña Then
    If txtNuevaContraseña = txtNuevaContraseña2 Then
        
      gdb.NewPassword gstrContraseña, txtNuevaContraseña
      
      gstrContraseña = txtNuevaContraseña
      SaveSetting "eCard", "Inicio", _
          "gstrContraseña", _
          Contraseña_Encriptar(gstrContraseña)
    
      Unload frmContraseñaCambiar
    Else
      MsgBox "La nueva contraseña no coincide", _
      vbInformation, "Cambiar Contraseña - Error"
      txtNuevaContraseña.SetFocus
    End If
  Else
    
    MsgBox "Contraseña Erronea", vbInformation, _
           "Cambiar Contraseña - Error"
    txtContraseña.SetFocus

  End If
End Sub

Private Sub cmdCancelar_Click()
  Unload frmContraseñaCambiar
End Sub

Private Sub Form_Load()
    frmContraseñaCambiar.Caption = "Cambiar Contraseña"
    lblContraseña = "Contraseña"
    lblNuevaContraseña.Caption = "Nueva Contraseña"
    lblNuevaContraseña2.Caption = "Reentrar nueva Contraseña"
    cmdAceptar.Caption = "&Aceptar"
    cmdCancelar.Caption = "&Cancelar"
End Sub

Private Sub txtContraseña_KeyPress(KeyAscii As Integer)
  If InStr("0123456789" & _
           "abcdefghijklmnopqrstuvwxyzñç" & _
           "ABCDEFGHIJKLMNOPQRSTUVWXYZÑÇ", _
           Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
  End If
End Sub
Private Sub txtNuevaContraseña_KeyPress(KeyAscii As Integer)
  If InStr("0123456789" & _
           "abcdefghijklmnopqrstuvwxyzñç" & _
           "ABCDEFGHIJKLMNOPQRSTUVWXYZÑÇ", _
           Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
  End If
End Sub
Private Sub txtNuevaContraseña2_KeyPress(KeyAscii As Integer)
  If InStr("0123456789" & _
           "abcdefghijklmnopqrstuvwxyzñç" & _
           "ABCDEFGHIJKLMNOPQRSTUVWXYZÑÇ", _
           Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
  End If
End Sub

