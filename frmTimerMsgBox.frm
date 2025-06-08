VERSION 5.00
Begin VB.Form frmTimerMsgBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRetraso 
      Left            =   240
      Top             =   960
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblMsg 
      Caption         =   "Label1"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmTimerMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
  Unload frmTimerMsgBox
End Sub

Private Sub Form_Unload(Cancel As Integer)
  tmrRetraso.Enabled = False
End Sub

Private Sub tmrRetraso_Timer()
  Unload frmTimerMsgBox
End Sub
