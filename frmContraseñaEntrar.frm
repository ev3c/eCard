VERSION 5.00
Begin VB.Form frmContraseñaEntrar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entrar Contraseña"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContraseña 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   3975
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
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmContraseñaEntrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  If txtContraseña.Text = gstrContraseña Then
    gblnContraseña = True
    Unload frmContraseñaEntrar
  Else
    MsgBox "Contraseña Erronea", vbInformation, "Entrar Contraseña - Error"
    txtContraseña.SetFocus
  End If
End Sub

Private Sub cmdCancelar_Click()
  gblnContraseña = False
  Unload frmContraseñaEntrar
End Sub

Private Sub Form_Load()
  gblnContraseña = False
  frmContraseñaEntrar.Caption = "Entrar Contraseña"
  lblContraseña = "Contraseña"
  cmdAceptar.Caption = "&Aceptar"
  cmdCancelar.Caption = "&Cancelar"
End Sub


