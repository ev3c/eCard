VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmPrograma 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9780
   Icon            =   "frmPrograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   4920
      TabIndex        =   35
      Top             =   0
      Width           =   2415
      Begin VB.TextBox txtPvpPag 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,000""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         ToolTipText     =   "Quantums por página impresa"
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Qua Página:"
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
         Left            =   240
         TabIndex        =   36
         ToolTipText     =   "Quantums por página impresa"
         Top             =   200
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   7920
      TabIndex        =   34
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox cboComm 
      Height          =   315
      ItemData        =   "frmPrograma.frx":030A
      Left            =   8160
      List            =   "frmPrograma.frx":0320
      Style           =   2  'Dropdown List
      TabIndex        =   23
      ToolTipText     =   "Especificar el puerto serie de la tarjeta"
      Top             =   240
      Width           =   1095
   End
   Begin MSCommLib.MSComm mscCard 
      Left            =   7320
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdEsconder 
      Caption         =   "&Esconder"
      Default         =   -1  'True
      Height          =   375
      Left            =   7920
      TabIndex        =   21
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCambiarContraseña 
      Caption         =   "&Cambiar Contraseña"
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdMover 
      Height          =   375
      Index           =   3
      Left            =   1440
      Picture         =   "frmPrograma.frx":0356
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdMover 
      Height          =   375
      Index           =   2
      Left            =   1080
      Picture         =   "frmPrograma.frx":04A0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdMover 
      Height          =   375
      Index           =   1
      Left            =   720
      Picture         =   "frmPrograma.frx":05EA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdMover 
      Height          =   375
      Index           =   0
      Left            =   360
      Picture         =   "frmPrograma.frx":0734
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Frame fraPrograma 
      Caption         =   "Programas"
      Height          =   1815
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   7335
      Begin MSComCtl2.DTPicker dtpPrgSesion 
         Height          =   300
         Left            =   5640
         TabIndex        =   33
         Top             =   800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24510467
         CurrentDate     =   36545
      End
      Begin VB.Frame Frame1 
         Caption         =   "Eventos"
         Height          =   1095
         Left            =   3480
         TabIndex        =   25
         Top             =   120
         Width           =   3615
         Begin VB.CommandButton cmdBorrarEventos 
            Caption         =   "Borrar Eventos"
            Height          =   315
            Left            =   2160
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
         Begin MSComCtl2.UpDown updPrgIdx 
            Height          =   255
            Left            =   1080
            TabIndex        =   27
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "lblPrgIdx"
            BuddyDispid     =   196628
            OrigLeft        =   4200
            OrigTop         =   840
            OrigRight       =   4395
            OrigBottom      =   1125
            Max             =   1
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblCard 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            Height          =   255
            Left            =   1560
            TabIndex        =   31
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Tarjeta número:"
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
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblSesion 
            Caption         =   "Sesión:"
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
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblPrgSesion 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   28
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblPrgIdx 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   840
            TabIndex        =   26
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   2
         ToolTipText     =   "Nivel de Acceso a los programas (0 mayor - 9 menor)"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtPvp 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,000""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Precio en Quantums por minuto"
         Top             =   960
         Width           =   855
      End
      Begin VB.Timer tmrPrograma 
         Left            =   7080
         Top             =   720
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         ToolTipText     =   "Nombre con el que identificamos al programa"
         Top             =   600
         Width           =   2175
      End
      Begin MSComDlg.CommonDialog dlgNombre 
         Left            =   7080
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar HD"
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   1275
         Width           =   1335
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         ToolTipText     =   "Dirección del programa .exe en el Disco Duro"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label lblLevel 
         Caption         =   "Nivel:"
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
         Left            =   2280
         TabIndex        =   24
         ToolTipText     =   "Nivel de Acceso a los programas (0 mayor - 9 menor)"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblPvp 
         Caption         =   "Qua min:"
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
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Precio en Quantums por minuto"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Path:"
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
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Dirección del programa .exe en el Disco Duro"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblNombre 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre:"
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
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Nombre con el que identificamos al programa"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
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
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdAñadir 
      Caption         =   "&Añadir"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblPaginas 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   41
      ToolTipText     =   "Número de páginas impresas"
      Top             =   195
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Páginas:"
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
      Left            =   3120
      TabIndex        =   40
      ToolTipText     =   "Número de páginas impresas"
      Top             =   195
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Consumo:"
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
      Left            =   1320
      TabIndex        =   39
      ToolTipText     =   "Consumo en Quantums por minuto"
      Top             =   195
      Width           =   855
   End
   Begin VB.Label lblConsumo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   38
      ToolTipText     =   "Consumo en Quantums por minuto"
      Top             =   195
      Width           =   735
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const EC_SAVESETTINGS = True
#Const EC_EXE = True

Private Sub cboComm_Click()
  Call Comm_Cerrar
  giComm = cboComm.ListIndex + 1
  Call Comm_Abrir
End Sub

Private Sub cmdEsconder_Click()
  On Error Resume Next
  
  frmPrograma.Hide
  App.TaskVisible = False
'  tmrPrograma.Interval = 1000
End Sub
Private Sub cmdCambiarContraseña_Click()
  If Not frmContraseñaCambiar.Visible Then
    Load frmContraseñaCambiar
    frmContraseñaCambiar.Show vbModal
  End If
End Sub

Public Function Contraseña_Entrar() As Boolean

  Load frmContraseñaEntrar
  frmContraseñaEntrar.Show vbModal
  If gblnContraseña = True Then
    Contraseña_Entrar = True
  Else
    Contraseña_Entrar = False
  End If

End Function




Private Sub cmdImprimir_Click()
  Load frmImprimir
  frmImprimir.Show vbModal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Select Case UnloadMode
    Case vbFormControlMenu, vbAppTaskManager
      If Salir_SiNo = vbNo Then
        Cancel = True
      End If
    Case vbAppWindows   ', vbAppTaskManager
  End Select
End Sub

Private Sub cmdAceptar_Click()
Dim strMsg0, strMsg1, strMsg2 As String
Dim lPos As Long
  strMsg0 = "Debe especificar el Nombre, Precio, Nivel y Path del Programa"
  strMsg1 = "El Path del Programa a Añadir no existe"
  strMsg2 = "Añadir Programa"
  
  If txtNombre = "" Or txtPath = "" Or txtPvp = "" Or txtLevel = "" Then
    MsgBox strMsg0, vbInformation, strMsg2
  Else
    If Dir(txtPath) = "" Then
      MsgBox strMsg1, vbInformation, strMsg2
    Else
      Call Campos_Grabar
      grsPrograma.Update
      Call Command_Mostrar
      lPos = lblID - 1
      Call Programas_Leer
      grsPrograma.AbsolutePosition = lPos
    End If
  End If
End Sub

Private Sub cmdAñadir_Click()
  Call Command_Ocultar
  
  If Not grsPrograma.EOF Then
    grsPrograma.MoveLast
    lblID = grsPrograma.Fields("ID") + 1
  Else
    lblID = "1"
  End If
  
  grsPrograma.AddNew
  txtNombre = ""
  txtPath = ""
  txtPvp = ""
  txtLevel = ""
  txtNombre.SetFocus

End Sub

Private Sub cmdBorrar_Click()
  Dim intBorrar, x As Integer
  Dim intID, intBorradoID As Integer
  
  strMsgPrg = "Borrar Programa"
  strMsgBorrar = "Seguro que desea Borrar el Programa " & grsPrograma.Fields("Nombre")
  strMsgNoBorrar = "No puede Borrar un Programa que está siendo Cronometrado"
  
  intBorradoID = grsPrograma.Fields("ID")
  
  If gaPrg(intBorradoID).on <> CDate("0") Then
    MsgBox strMsgNoBorrar, vbInformation, strMsgPrg
  Else
      
    intBorrar = MsgBox(strMsgBorrar, vbYesNo + vbCritical, strMsgPrg)
    If intBorrar = vbYes Then
      grsPrograma.Delete
      grsPrograma.MoveFirst
      Do While Not grsPrograma.EOF
        intID = grsPrograma.Fields("ID")
        If intID > intBorradoID Then
          grsPrograma.Edit
          grsPrograma.Fields("ID") = intID - 1
          grsPrograma.Update
        End If
        grsPrograma.MoveNext
      Loop
       
      If grsEvento.RecordCount > 0 Then
        grsEvento.MoveFirst
        Do While Not grsEvento.EOF
          intID = grsEvento.Fields("ProgramaID")
          If intID > intBorradoID Then
            grsEvento.Edit
            grsEvento.Fields("ProgramaID") = intID - 1
            grsEvento.Update
          End If
          grsEvento.MoveNext
        Loop
      End If
      x = intBorradoID
      Do While gaPrg(x).Path <> ""
        gaPrg(x).on = gaPrg(x + 1).on
        gaPrg(x).Path = gaPrg(x + 1).Path
        gaPrg(x).pvp = gaPrg(x + 1).pvp
        gaPrg(x).level = gaPrg(x + 1).level
        x = x + 1
      Loop
      
      gaPrg(x).on = CDate("0")
      gaPrg(x).Path = ""
      gaPrg(x).pvp = 0
      gaPrg(x).level = 0
          
      Call Campos_Ver
      Call Command_Mostrar
      Call Evento_Ver
      
    End If
  
  End If

End Sub

Private Sub cmdCancelar_Click()
  If Not grsPrograma.EOF Then
    grsPrograma.CancelUpdate
  End If
  Call Campos_Ver
  Call Command_Mostrar
End Sub

Private Sub cmdExaminar_Click()
       
  On Error GoTo errHandler
  
  dlgNombre.DialogTitle = "Examinar Programas"
  dlgNombre.Filter = "Archivos de Programa (*.exe) |*.exe"
  
  dlgNombre.CancelError = True
  dlgNombre.InitDir = "C:\"
  dlgNombre.ShowOpen
  
  txtPath = dlgNombre.FileName
      
  Exit Sub
    
errHandler:
txtFichero = Error
End Sub

Private Sub cmdModificar_Click()
  strMsgPrg = "Modificar Programa"
  strMsgNoBorrar = "No puede Modificar un Programa que está siendo Cronometrado"
  
  intBorradoID = grsPrograma.Fields("ID")
  
  If gaPrg(intBorradoID).on <> CDate("0") Then
    MsgBox strMsgNoBorrar, vbInformation, strMsgPrg
  Else
    Call Command_Ocultar
    grsPrograma.Edit
    txtNombre.SetFocus
  End If
End Sub

Private Sub cmdSalir_Click()
  
  If Salir_SiNo = vbYes Then
    Call Form_Unload(False)
  End If
  
End Sub
  
Private Sub Form_Load()
   
  On Error Resume Next
  
  'Establece MyDate
  MyDate = Date
  gstrFormatoFecha = "dd/MM/yyyy"
  
  'Comprueba si ya está activo
  If App.PrevInstance Then
    End
  End If
    
#If EC_EXE Then
  'Subclassifica la finestra per captura Ctrl+Alt+Mays+C
  Dim lRet As Long
  lRet = RegisterHotKey(frmPrograma.hwnd, &HB000&, _
         MOD_ALT Or MOD_CONTROL Or MOD_SHIFT, vbKeyC)
  Call Subclasifica_Ventana(frmPrograma.hwnd)

  If getVersion = 1 Then
    'Oculta eCard de Ctrl+Alt+Supr
    RegisterServiceProcess GetCurrentProcessId, 1
    'Hide app
  End If

#End If
  
  gstrPrograma = "eCard v" & _
    App.Major & "." & App.Minor & "." & App.Revision
  frmPrograma.Caption = gstrPrograma

  ReDim gaPrg(1 To 99) As ecOnPathPvp
  ReDim gaPrgAct(1 To 99) As ecPath
  
  Call Evento_Abrir
  
  Call Programas_Leer

'Esconde el formulario.
#If EC_EXE Then
  frmPrograma.Hide
#End If
  
  Call Configuracion_Settings_Leer

  gdFechaOn = MyDate
  
  'Comprueba eventos cada segundo
  tmrPrograma.Interval = 1000
    
  dtpPrgSesion.CustomFormat = gstrFormatoFecha
  dtpPrgSesion.Value = MyDate
    
    
  Call Campos_Ver
  Call Command_Mostrar
  
  Call Evento_Ver
  
  lblConsumo = Format(0, "#,##0")
 
  If cboComm.ListIndex = -1 Then
    giComm = 1
  Else
    giComm = cboComm.ListIndex + 1
  End If

End Sub
Private Sub cmdMover_Click(Index As Integer)
  
  If Not grsPrograma.EOF Then
    Select Case Index
      Case 0
        grsPrograma.MoveFirst
      Case 1
        grsPrograma.MovePrevious
      Case 2
        grsPrograma.MoveNext
      Case 3
        grsPrograma.MoveLast
    End Select
    If grsPrograma.BOF Then grsPrograma.MoveFirst
    If grsPrograma.EOF Then grsPrograma.MoveLast
  
  End If

  Call Campos_Ver
  Call Evento_Ver
  
End Sub

Public Sub Campos_Ver()
  lblID = ""
  txtNombre = ""
  txtPath = ""
  txtPvp = ""
  txtLevel = ""
    
  If grsPrograma.RecordCount >= 1 Then
    If grsPrograma.AbsolutePosition < 0 Then
       grsPrograma.MoveFirst
    End If
    lblID = grsPrograma.Fields("ID")
    txtNombre = grsPrograma.Fields("Nombre")
    txtPath = grsPrograma.Fields("Path")
    txtPvp = Format(grsPrograma.Fields("pvp_hora"), "#,##0")
    txtLevel = grsPrograma.Fields("Level")
  End If
  

End Sub

  


Private Sub Campos_Grabar()
  grsPrograma.Fields("ID") = lblID
  grsPrograma.Fields("Nombre") = txtNombre
  grsPrograma.Fields("Path") = txtPath
  grsPrograma.Fields("pvp_hora") = CDbl(txtPvp)
  grsPrograma.Fields("Level") = CInt(txtLevel)
End Sub

Private Sub Command_Ocultar()
  cmdAñadir.Visible = False
  cmdModificar.Visible = False
  cmdBorrar.Visible = False
  
  cmdAceptar.Visible = True
  cmdCancelar.Visible = True
  cmdExaminar.Enabled = True
   
  cmdMover.Item(0).Enabled = False
  cmdMover.Item(1).Enabled = False
  cmdMover.Item(2).Enabled = False
  cmdMover.Item(3).Enabled = False
 
  cmdSalir.Enabled = False
  cmdEsconder.Enabled = False
  cmdCambiarContraseña.Enabled = False
  
  txtNombre.Enabled = True
  txtPath.Enabled = True
  txtPvp.Enabled = True
  txtLevel.Enabled = True
  
End Sub

Public Sub Command_Mostrar()
  cmdAñadir.Visible = True
  cmdModificar.Visible = True
  cmdBorrar.Visible = True
  
  cmdAceptar.Visible = False
  cmdCancelar.Visible = False
  cmdExaminar.Enabled = False
  
  cmdMover.Item(0).Enabled = True
  cmdMover.Item(1).Enabled = True
  cmdMover.Item(2).Enabled = True
  cmdMover.Item(3).Enabled = True
  
  cmdSalir.Enabled = True
  cmdEsconder.Enabled = True
  cmdCambiarContraseña.Enabled = True
  
  txtNombre.Enabled = False
  txtPath.Enabled = False
  txtPvp.Enabled = False
  txtLevel.Enabled = False
  
  If lblID = "" Then
    cmdModificar.Enabled = False
    cmdBorrar.Enabled = False
  Else
    cmdModificar.Enabled = True
    cmdBorrar.Enabled = True
  End If
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If Cancel = False Then
   
   tmrPrograma.Enabled = False
   Call Configuracion_Settings_Grabar
   Call Evento_Cerrar
   
#If EC_EXE Then
   'Recupera la HotKey y desclasifica ventana
   Call UnregisterHotKey(frmPrograma.hwnd, &HB000&)
   Call Ventana_Normal(frmPrograma.hwnd)

 ' Vuelve a mostrar Programa en Ctrl+Alt+Supr
   If getVersion = 1 Then
     'Put the following code in Form_Unload()
     RegisterServiceProcess GetCurrentProcessId, 0
     'Remove service flag
   End If

#End If

   '***************
   'End the program
   End
  
  End If
  
End Sub



Private Sub tmrPrograma_Timer()
  Dim iPrg As Integer
  Dim IsPrgOn(0 To 99) As Boolean
  Dim dFecha As Date
  Dim x, iFile As Integer
  
  Dim iMinNum, iQuaMin As Integer
  Dim dMinPvp, dTotal As Double
  Static dSaldo As Double
  Static dTimeUltimo As Date
    
  Static bApagar, bPrgOn, bPocoSaldo As Boolean
  Static iApagarMin, iApagarMinNum As Integer
  
  Static iCardID, iLevel, iStat As Integer  'lectura comm
  
  Call GetIsPrgOn(IsPrgOn())
  Call Comm_Leer(iCardID, iLevel, iStat)

  'Controla Apagar los programas
  If bApagar Then
    iApagarMin = iApagarMin + 1
    If iApagarMin >= iApagarMinNum Then
      iPrg = 1
      Do While gaPrg(iPrg).Path <> ""
        If gaPrg(iPrg).on <> CDate("0") Then
          iFile = 1
          Do
            If UCase(gaPrg(iPrg).Path) = _
               UCase(gaPrgAct(iFile).Path) Then
              Call Programa_Apagar(gaPrgAct(iFile).Path)
            End If
            iFile = iFile + 1
          Loop While gaPrgAct(iFile).Path <> ""
          If gaPrg(iPrg).card <> 0 Then
            Call Evento_Grabar(iPrg, gaPrg(iPrg).card, _
                MyDate, gaPrg(iPrg).on, Time)
          End If
          gaPrg(iPrg).on = CDate("0")
          gaPrg(iPrg).card = 0
        End If
        iPrg = iPrg + 1
      Loop
      dSaldo = 0
      iApagarMin = 0
      bApagar = False
      bPocoSaldo = False
      dTimeUltimo = Time
    End If
  
  Else
    
    'Controla Programa y gasto
    bPrgOn = False
    iPrg = 1
    Do While gaPrg(iPrg).Path <> ""
      If IsPrgOn(iPrg) Then
        bPrgOn = True
        If gaPrg(iPrg).on = CDate("0") Then
          If iLevel <= gaPrg(iPrg).level Then
            gaPrg(iPrg).on = Time
            gaPrg(iPrg).card = iCardID
          Else
            Call TimerMsgBox("Su tarjeta no tiene suficiente" & vbCrLf & _
                             "nivel para utilizar este programa, " & vbCrLf & _
                             "por lo que va a ser apagado", gstrPrograma, 10)
            
            Call Programa_Apagar(gaPrg(iPrg).Path)
          End If
        Else
          iMinNum = DateDiff("n", gaPrg(iPrg).on, Time) + 1
          dMinPvp = gaPrg(iPrg).pvp
          dTotal = dTotal + (iMinNum * dMinPvp)
          iQuaMin = iQuaMin + dMinPvp
        End If
      Else
        If gaPrg(iPrg).on <> CDate("0") Then
          iMinNum = DateDiff("n", gaPrg(iPrg).on, Time) + 1
          dMinPvp = gaPrg(iPrg).pvp
          dSaldo = dSaldo + (iMinNum * dMinPvp)
          
          Call Evento_Grabar(iPrg, gaPrg(iPrg).card, _
                MyDate, gaPrg(iPrg).on, Time)

          gaPrg(iPrg).on = CDate("0")
          gaPrg(iPrg).card = 0
        End If
      End If
      iPrg = iPrg + 1
    Loop
    
    'Controla 12PM
    If Time >= #11:59:59 PM# Then
      MyDate = DateAdd("d", 1, MyDate)
      Exit Sub
    End If
    
    'Controla Medianoche
    If gdFechaOn < MyDate Then
      gdFechaOn = MyDate
      dTimeUltimo = #12:00:00 AM#
      
      iPrg = 1
      Do While gaPrg(iPrg).Path <> ""
        If gaPrg(iPrg).on <> CDate("0") Then
          iMinNum = DateDiff("n", gaPrg(iPrg).on, #11:59:59 PM#) + 1
          dMinPvp = gaPrg(iPrg).pvp
          dSaldo = dSaldo + (iMinNum * dMinPvp)
          
          Call Evento_Grabar(iPrg, gaPrg(iPrg).card, _
                MyDate, gaPrg(iPrg).on, #11:59:59 PM#)
          
          gaPrg(iPrg).on = CDate("0")
          gaPrg(iPrg).card = 0
        End If
        iPrg = iPrg + 1
      Loop
    End If
     
     'Controla Retraso Hora
    If dTimeUltimo > Time And _
       dTimeUltimo <> CDate("0") Then
      Call TimerMsgBox("Se ha atrasado el RELOJ por lo que los" & vbCrLf & _
                  "programas utilizados van a ser apagados" & vbCrLf & _
                  "dentro de tres minutos", gstrPrograma, 10)
      bApagar = True
      iApagarMinNum = 180
    'No tarjeta y programa
    ElseIf iStat = 0 And bPrgOn Then
      Call TimerMsgBox("No hay TARJETA insertada por lo que los" & vbCrLf & _
                  "programas utilizados van a ser apagados", gstrPrograma, 10)
      bApagar = True
      iApagarMinNum = 0
    'No Saldo y programa
    ElseIf iStat = 2 And bPrgOn Then
      Call TimerMsgBox("No hay SALDO en la tarjeta por lo que los" & vbCrLf & _
                  "programas utilizados van a ser apagados", gstrPrograma, 10)
      iApagarMinNum = 0
      bApagar = True
    'Poco Saldo y programa
    ElseIf iStat = 3 And bPrgOn And Not bPocoSaldo Then
      Call TimerMsgBox("Queda poco SALDO en la tarjeta." & vbCrLf & _
                       "Guarde los trabajos y cierre los programas" & vbCrLf & _
                       "antes de que se cierren automaticamente", gstrPrograma, 10)
      bPocoSaldo = True
    'Ni tarjeta ni Programa
    ElseIf iStat = 0 And Not bPrgOn Then
      dSaldo = 0
      bPocoSaldo = False
    End If
      
    dTimeUltimo = Time
    If dSaldo + dTotal >= 0 Then
      lblConsumo = Format(dSaldo + dTotal, "#,##0")
      Call Comm_Grabar(dSaldo + dTotal, iQuaMin)
    End If
  
  End If
End Sub

Public Sub Evento_Abrir()
  Dim strMsg As String
  
  On Error GoTo Evento_Error
  
  gstrContraseña = Contraseña_DesEncriptar( _
      GetSetting("eCard", "Inicio", _
      "gstrContraseña", ""))
  
  
  Set gdb = DBEngine.Workspaces(0).OpenDatabase( _
       App.Path & "\eCard.mdb", True, False, _
       ";PWD=" & gstrContraseña)

  Set grsPrograma = gdb.OpenRecordset("SELECT * " & _
                    "FROM tblPrograma")
  
  Set grsEvento = gdb.OpenRecordset("SELECT * " & _
                  "FROM tblEvento " & _
                  "ORDER BY ProgramaID, Fecha, HoraOn")
 
  Exit Sub
  
Evento_Error:

  strMsg = "Error al Abrir la Base de Datos" & vbCrLf & _
            Err.Number & " " & Err.Description
  MsgBox strMsg, vbCritical, gstrPrograma
  End
  
End Sub


Public Sub Evento_Cerrar()
  
  iPrg = 1
  Do While gaPrg(iPrg).Path <> ""
    If gaPrg(iPrg).on <> CDate("0") Then
      Call Evento_Grabar(iPrg, gaPrg(iPrg).card, _
          MyDate, gaPrg(iPrg).on, Time)
    End If
    iPrg = iPrg + 1
  Loop
  
  grsPrograma.Close
  Set grsPrograma = Nothing
  grsEvento.Close
  Set grsEvento = Nothing
  gdb.Close
  Set gdb = Nothing
  
End Sub

Public Function Salir_SiNo()
  If Contraseña_Entrar Then
    strMsg = "Seguro que desea salir de " & gstrPrograma
    Salir_SiNo = MsgBox(strMsg, vbYesNo + vbCritical, gstrPrograma)
  Else
    Salir_SiNo = vbNo
  End If
End Function


Public Sub Configuracion_Settings_Leer()
  On Error Resume Next
  
  gstrContraseña = Contraseña_DesEncriptar( _
      GetSetting("eCard", "Inicio", _
      "gstrContraseña", ""))
   
  cboComm.ListIndex = GetSetting("eCard", _
      "Inicio", "cboComm", 0)

  txtPvpPag = GetSetting("eCard", _
      "Inicio", "txtPvpPag", 0)

End Sub

Public Sub Configuracion_Settings_Grabar()

#If EC_SAVESETTINGS Then        'No ho entenc
  
  SaveSetting "eCard", "Inicio", _
      "gstrContraseña", Contraseña_Encriptar(gstrContraseña)
  
  SaveSetting "eCard", "Inicio", _
      "cboComm", cboComm.ListIndex
  
  res = SetRegValue(HKEY_LOCAL_MACHINE, _
  "Software\Microsoft\Windows\CurrentVersion\Run", _
  "eCard", App.Path & "\eCard.exe")
   
  SaveSetting "eCard", "Inicio", _
      "txtPvpPag", txtPvpPag
   
#End If

End Sub

Private Sub txtPvp_KeyPress(KeyAscii As Integer)
  If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
    End If
'  Else
'    If Chr(KeyAscii) = "," And InStr(txtPvp, ",") <> 0 Then
'      KeyAscii = 0
'      Beep
'    End If
  End If
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
  If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
    End If
  End If
End Sub

Public Sub GetIsPrgOn(IsPrgOn() As Boolean)
  Dim iPrg, iFile As Integer
  
  Call GetWinFiles
  
  iPrg = 1
  Do While gaPrg(iPrg).Path <> ""
    
    iFile = 1
    Do
      If UCase(gaPrg(iPrg).Path) = _
         UCase(gaPrgAct(iFile).Path) Then
        IsPrgOn(iPrg) = True
        Exit Do
      End If
      iFile = iFile + 1
    Loop While gaPrgAct(iFile).Path <> ""
    
    iPrg = iPrg + 1
  Loop
  
End Sub

Private Sub Programas_Leer()
  Dim iPrg As Integer
  
  grsPrograma.MoveFirst
  iPrg = 1
  Do While Not grsPrograma.EOF
    gaPrg(iPrg).Path = grsPrograma.Fields("Path")
    gaPrg(iPrg).pvp = grsPrograma.Fields("pvp_hora")
    gaPrg(iPrg).level = grsPrograma.Fields("Level")
    iPrg = iPrg + 1
    grsPrograma.MoveNext
  Loop

End Sub

Private Sub Comm_Leer(iCardID, iLevel, iStat)
  Dim sStat As String
  Static iNoInput As Integer
  
  On Error GoTo Comm_Leer_Err:
  
  If mscCard.InBufferCount <> 0 Then
    mscCard.InputLen = 0        'Per recuperar tot el contingut del buffer.
    sStat = mscCard.Input       'LLegueix pel port
    iNoInput = 0
  Else

#If EC_EXE Then
    iNoInput = iNoInput + 1
    If iNoInput = 60 Then
      iNoInput = 0
      iStat = 0
    End If
#End If

#If Not EC_EXE Then
    sStat = "0010L9S1"
#End If
    
  End If
  
  If Mid$(sStat, 5, 1) = "L" And _
     Mid$(sStat, 7, 1) = "S" Then     'No error en la comunicació
    iCardID = CInt(Left$(sStat, 4))   'Número tarjeta
    iLevel = CInt(Mid$(sStat, 6, 1))  'Level
    iStat = CInt(Mid$(sStat, 8, 1))   'Stat
  End If
  
   
Exit Sub
Comm_Leer_Err:
  Call MsgBox("Input Error", vbInformation, "Input Error " & Err.Number & Err.Description)

End Sub

Private Sub Comm_Grabar(dCantidad As Double, iQuaMin As Integer)
  Dim sOutput As String

  On Error GoTo Comm_Grabar_Err:

  If mscCard.OutBufferCount = 0 Then
    sOutput = Trim(Str(CInt(dCantidad))) & _
    ";" & Trim(Str(iQuaMin)) & "X" 'Escriu pel port
    mscCard.Output = sOutput
  End If
  
  Exit Sub

Comm_Grabar_Err:


End Sub

Private Sub Comm_Abrir()
On Error GoTo Comm_Abrir_Err:
  
  mscCard.CommPort = giComm         'Estableix Com1
  mscCard.Settings = "9600,N,8,1"   'Estableix els paràmetres de comunicació
                                    '9600 Bauds, No paritat, 8 bits de dades, 1 bit de parada.
  mscCard.PortOpen = True
  
  Exit Sub
  
Comm_Abrir_Err:

  MsgBox "El puerto no existe o ya está abierto : " & Err.Number, vbInformation, "Abrir Puerto Serie"
  Call Comm_Cerrar
  
End Sub

Private Sub Comm_Cerrar()
On Error Resume Next
  mscCard.PortOpen = False
End Sub

Public Sub TimerMsgBox(Texto As String, Titulo As String, Tiempo As Integer)
  
  frmTimerMsgBox.lblMsg = Texto
  frmTimerMsgBox.Caption = Titulo
  frmTimerMsgBox.tmrRetraso.Interval = Tiempo * 1000
  Beep
  
  Load frmTimerMsgBox
  SetWindowPos frmTimerMsgBox.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
               SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
  frmTimerMsgBox.Show vbModal

End Sub

Public Sub Evento_Grabar(id, CardID, Fecha, HoraOn, HoraOff)

  grsEvento.AddNew
  grsEvento.Fields("ProgramaID") = id
  grsEvento.Fields("CardID") = CardID
  grsEvento.Fields("Fecha") = Fecha
  grsEvento.Fields("HoraOn") = HoraOn
  grsEvento.Fields("HoraOff") = HoraOff
  grsEvento.Update
  
End Sub

Private Sub cmdBorrarEventos_Click()
  
  strMsgPrg = "Borrar Eventos"
  strMsgBorrar = "Seguro que desea Borrar Todos los Eventos "
    
  intBorrar = MsgBox(strMsgBorrar, vbYesNo + vbCritical, strMsgPrg)
  If intBorrar = vbYes Then
    If grsEvento.RecordCount > 0 Then
      grsEvento.MoveFirst
    End If
    Do While Not grsEvento.EOF
      grsEvento.Delete
      grsEvento.MoveNext
    Loop
     
    Call Evento_Ver
 
  End If
End Sub

Public Sub Evento_Ver()
Dim strMsg As Integer

On Error GoTo Evento_Ver_Err:

If grsPrograma.RecordCount > 0 Then
  Set grsPrg = gdb.OpenRecordset("SELECT * " & _
               "FROM tblEvento " & _
               "WHERE ProgramaID = " & lblID & _
               " AND Fecha = #" & Format(dtpPrgSesion, "MM/dd/yy") & _
               "# ORDER BY ProgramaID, Fecha, HoraOn")


  If grsPrg.RecordCount > 0 Then
    grsPrg.MoveFirst
    updPrgIdx.Min = 1
    updPrgIdx.Max = grsPrg.RecordCount
    updPrgIdx.Value = 1
  Else
    updPrgIdx.Min = 1
    updPrgIdx.Max = 1
    updPrgIdx.Value = 1
  End If
                
End If
  
  
Exit Sub
Evento_Ver_Err:
  strMsg = "Error al Abrir la Base de Datos" & vbCrLf & _
            Err.Number & " " & Err.Description
  MsgBox strMsg, vbCritical, gstrPrograma
                        
End Sub

Private Sub updPrgIdx_Change()
  
  If grsPrg.RecordCount > 0 Then
    grsPrg.AbsolutePosition = updPrgIdx.Value - 1
    lblPrgSesion = Format(grsPrg.Fields("HoraOff") - _
                          grsPrg.Fields("HoraOn"), "h:mm:ss")
    lblPrgSesion.ToolTipText = grsPrg.Fields("HoraOn") & _
                          " - " & grsPrg.Fields("HoraOff")
    lblCard = grsPrg.Fields("CardID")
  Else
    lblPrgSesion = Format(CDate(0), "h:mm:ss")
    lblPrgSesion.ToolTipText = CDate(0) & " - " & CDate(0)
    lblCard = "0000"
  End If
  
End Sub

Private Sub dtpPrgSesion_Change()

  Call Evento_Ver

End Sub

Private Sub txtPvpPag_KeyPress(KeyAscii As Integer)
  If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
    End If
'  Else
'    If Chr(KeyAscii) = "," And InStr(txtPvp, ",") <> 0 Then
'      KeyAscii = 0
'      Beep
'    End If
  End If
End Sub

