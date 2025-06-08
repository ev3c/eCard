VERSION 5.00
Object = "{7562DC0C-75BD-41F9-AE70-5C4A30EEF693}#1.0#0"; "gsCtlPrinterJobs.ocx"
Begin VB.Form fPrinterJobsTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prueba de Printer Jobs"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "fPrinterJobsTestAX.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin gsCtlPrinterJobs.ucPrinterJobs ucPrinterJobs1 
      Left            =   2280
      Top             =   1020
      _ExtentX        =   900
      _ExtentY        =   900
      Interval        =   700
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Estado"
      Height          =   405
      Left            =   7380
      TabIndex        =   3
      ToolTipText     =   " Informar del estado de la impresora (no operativo) "
      Top             =   360
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Detener"
      Height          =   405
      Left            =   6060
      TabIndex        =   2
      ToolTipText     =   " Terminar la captura de la información "
      Top             =   360
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   150
      TabIndex        =   4
      ToolTipText     =   " Haz doble-click para copiar el contenido en el protapapeles "
      Top             =   840
      Width           =   8475
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   405
      Left            =   4740
      TabIndex        =   1
      ToolTipText     =   " Iniciar la captura de la impresora indicada "
      Top             =   360
      Width           =   1245
   End
   Begin VB.ComboBox cboPrinters 
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   390
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Impresoras:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   " Lista de impresoras disponibles "
      Top             =   150
      Width           =   1965
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   2940
      Width           =   8445
   End
End
Attribute VB_Name = "fPrinterJobsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Prueba para saber los trabajos pendientes de una impresora        (08/Jun/01)
' Usando una DLL escrita en CBuilder
'
' Revisión usando el control ucPrinterJobs1                         (14/Dic/01)
'
' ©Guillermo 'guille' Som, 2001
'------------------------------------------------------------------------------
Option Explicit
'

Private Sub cmdCancel_Click()
    ucPrinterJobs1.Cancelar = True
End Sub

Private Sub cmdInfo_Click()
    Dim s As String
    Dim n As Long
    '
    Dim iStatus As ePrinterStatus
    '
    s = cboPrinters.Text
    ucPrinterJobs1.NombreImpresora = cboPrinters.Text
    List1.AddItem s
    '
    'n = ucPrinterJobs1.Version(True)
    '
    ' Ejemplo de PrinterStatus:
    iStatus = ucPrinterJobs1.PrinterStatus
    '
    '
    ' Número de trabajos pendientes y número de páginas             (11/Jun/01)
    ' Para saber el número de páginas impresas                      (13/Dic/01)
    '
    ucPrinterJobs1.Iniciar
    '
End Sub

Private Sub cmdStatus_Click()
    Call ucPrinterJobs1.PrinterStatus(cboPrinters.Text)
End Sub

Private Sub ucPrinterJobs1_PrinterJobsEvent(Cancel As Boolean, ByVal nJobs As Long, ByVal nTotalPages As Long, ByVal nPrintedPages As Long)
    ' Si se asigna el valor True a Cancel, se cancela la información
    lblStatus = "Trabajos: " & CStr(nJobs) & ", Páginas: " & CStr(nTotalPages) & ", impresas: " & CStr(nPrintedPages)
End Sub

Private Sub ucPrinterJobs1_PrinterJobsInfoEvent(ByVal HayError As ePrinterJobsInfo, ByVal sDescription As String)
    List1.AddItem sDescription
    lblStatus = sDescription
End Sub

Private Sub ucPrinterJobs1_PrinterStatusEvent(ByVal sDescription As String)
    List1.AddItem sDescription
    lblStatus = sDescription
End Sub

Private Sub Form_Load()
    ' Enumerar las impresoras disponibles
    Dim tPrinter As Printer
    '
    ' Añadir las impresoras disponibles
    For Each tPrinter In Printers
        cboPrinters.AddItem tPrinter.DeviceName
    Next
    ' Asignar la variable de la impresora seleccionada
    Set tPrinter = Printer
    If cboPrinters.ListCount > 0 Then
        cboPrinters.Text = tPrinter.DeviceName
    End If
    '
    With ucPrinterJobs1
        .NombreImpresora = cboPrinters.Text
        ' Está asignado a 700 en las propiedades
        '.Interval = 700
    End With
    '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ucPrinterJobs1.Cancelar = True
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Si se está usando el formulario, hay que descargarlo
    '
    Set fPrinterJobsTest = Nothing
End Sub

Private Sub List1_DblClick()
    ' Copiar en el portapapeles
    Dim i As Long
    Dim s As String
    '
    s = ""
    With List1
        For i = 0 To .ListCount - 1
            s = s & .List(i) & vbCrLf
        Next
    End With
    '
    Clipboard.Clear
    Clipboard.SetText s
End Sub


