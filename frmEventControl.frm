VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEventControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EventControl 2000 v1.0"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInternetHistorial 
      Caption         =   "Internet &Historial"
      Height          =   375
      Left            =   6840
      TabIndex        =   59
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5400
      TabIndex        =   56
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdPausar 
      Caption         =   "Pa&usar"
      Height          =   375
      Left            =   6840
      TabIndex        =   45
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdInformacion 
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdAnalizarFecha 
      Cancel          =   -1  'True
      Caption         =   "Analizar &Fecha"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   1560
      Width           =   1695
   End
   Begin VB.PictureBox picIconoBandeja 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3720
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer tmrPrograma 
      Left            =   3960
      Top             =   840
   End
   Begin VB.Frame fraControlTiempo 
      Caption         =   "Control de Tiempo"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   8055
      Begin VB.CommandButton cmdCrono 
         Caption         =   "On Crono"
         Height          =   300
         Left            =   5280
         TabIndex        =   53
         Top             =   250
         Width           =   1215
      End
      Begin VB.ComboBox cboPrg 
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.UpDown updWinIdx 
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblWinIdx"
         BuddyDispid     =   196644
         OrigLeft        =   1200
         OrigTop         =   480
         OrigRight       =   1395
         OrigBottom      =   735
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updScrIdx 
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblScrIdx"
         BuddyDispid     =   196643
         OrigLeft        =   2640
         OrigTop         =   480
         OrigRight       =   2835
         OrigBottom      =   735
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updMdmIdx 
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblMdmIdx"
         BuddyDispid     =   196642
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
      Begin MSComCtl2.UpDown updPrgIdx 
         Height          =   255
         Left            =   6960
         TabIndex        =   40
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblPrgIdx"
         BuddyDispid     =   196629
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
      Begin MSComCtl2.UpDown updTmrIdx 
         Height          =   255
         Left            =   5520
         TabIndex        =   47
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblTmrIdx"
         BuddyDispid     =   196619
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
      Begin VB.Label lblTmrIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   52
         Top             =   600
         Width           =   300
      End
      Begin VB.Label lblTmrSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   51
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblTmrDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblTmrMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   49
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblTmrAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   48
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPrgAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   44
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblPrgMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   43
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblPrgDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   42
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPrgSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   41
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblPrgIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6720
         TabIndex        =   39
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblMdmAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   37
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblMdmMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblMdmDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   35
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblMdmSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblScrAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblScrMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblScrDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblScrSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblWinAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblWinMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblWinSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblMdmIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   600
         Width           =   450
      End
      Begin VB.Label lblScrIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblWinIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblMdm 
         Caption         =   "Modem"
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
         Left            =   4320
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblScr 
         Caption         =   "ScrSaver"
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
         Left            =   2880
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblWindows 
         Caption         =   "Windows"
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
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblAño 
         Caption         =   "Año"
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
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblMes 
         Caption         =   "Mes"
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDia 
         Caption         =   "Día"
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
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblSesion 
         Caption         =   "Sesión"
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
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdProgramas 
      Caption         =   "&Programas"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCambiarContraseña 
      Caption         =   "&Contraseña"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cboIdioma 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox chkAutoArranque 
      Caption         =   "Arrancar EventControl al iniciar Windows"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.CheckBox chkVerPantalla 
      Caption         =   "Ver esta pantalla antes de salir de Windows"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.CheckBox chkMostrarIcono 
      Caption         =   "Mostrar Icono en Barra de Tareas"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Frame fraConfiguracion 
      Caption         =   "Configuración"
      Height          =   1155
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      Begin VB.PictureBox picIconoBandejaPausa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   4320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   58
         Top             =   600
         Visible         =   0   'False
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   54
      Top             =   1320
      Width           =   5175
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   1800
         TabIndex        =   57
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24510467
         CurrentDate     =   36545
      End
      Begin VB.Label lblIntroduceFecha 
         Caption         =   "Introduce fecha:"
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
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Menu mnuIconoBandeja 
      Caption         =   "IconoBandeja"
      Visible         =   0   'False
      Begin VB.Menu mnuVerControlTiempo 
         Caption         =   "&Ver Control de Tiempo"
      End
      Begin VB.Menu mnuNull0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCambiarContraseña 
         Caption         =   "&Cambiar Contraseña"
      End
      Begin VB.Menu mnuProgramas 
         Caption         =   "&Programas"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuInternetHistorial 
         Caption         =   "Internet &Historial"
      End
      Begin VB.Menu mnuNull1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPausar 
         Caption         =   "Pa&usar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmEventControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
