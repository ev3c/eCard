VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prueba de Printer Jobs"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   150
      TabIndex        =   2
      Top             =   780
      Width           =   6885
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   405
      Left            =   5790
      TabIndex        =   1
      Top             =   300
      Width           =   1245
   End
   Begin VB.ComboBox cboPrinters 
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   330
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Prueba para saber los trabajos pendientes de una impresora        (08/Jun/01)
' Usando una DLL escrita en CBuilder
'
' ©Guillermo 'guille' Som, 2001
'------------------------------------------------------------------------------
Option Explicit
'
' Constantes para saber el estado de la impresora (enumeración)
Private Enum ePrinterStatus
    PRSTAT_NOESTA = 0&
    PRSTAT_FALLO = -1&
    PRINTER_STATUS_PAUSED = &H1
    PRINTER_STATUS_ERROR = &H2
    PRINTER_STATUS_PENDING_DELETION = &H4
    PRINTER_STATUS_PAPER_JAM = &H8
    PRINTER_STATUS_PAPER_OUT = &H10
    PRINTER_STATUS_MANUAL_FEED = &H20
    PRINTER_STATUS_PAPER_PROBLEM = &H40
    PRINTER_STATUS_OFFLINE = &H80
    PRINTER_STATUS_IO_ACTIVE = &H100
    PRINTER_STATUS_BUSY = &H200
    PRINTER_STATUS_PRINTING = &H400
    PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
    PRINTER_STATUS_NOT_AVAILABLE = &H1000
    PRINTER_STATUS_WAITING = &H2000
    PRINTER_STATUS_PROCESSING = &H4000
    PRINTER_STATUS_INITIALIZING = &H8000
    PRINTER_STATUS_WARMING_UP = &H10000
    PRINTER_STATUS_TONER_LOW = &H20000
    PRINTER_STATUS_NO_TONER = &H40000
    PRINTER_STATUS_PAGE_PUNT = &H80000
    PRINTER_STATUS_USER_INTERVENTION = &H100000
    PRINTER_STATUS_OUT_OF_MEMORY = &H200000
    PRINTER_STATUS_DOOR_OPEN = &H400000
    PRINTER_STATUS_SERVER_UNKNOWN = &H800000
    PRINTER_STATUS_POWER_SAVE = &H1000000
End Enum
' Declaración de las funciones
'
' Funciones incluidas a partir del 13 de Diciembre 2001
Private Declare Function gsPJVersion Lib "gsPrinterJobs.dll" Alias "Version" () As Long
'
Private Declare Function GetPrinterStatus Lib "gsPrinterJobs.dll" _
    (ByVal sDeviceName As String) As Long
'
' Esta en particular, es por encargo de Esteve Valentí
Private Declare Function GetNumPagesEx Lib "gsPrinterJobs.dll" _
    (ByVal sDeviceName As String, _
     ByRef nJobs As Long, _
     ByRef nPagesPrinted As Long) As Long
'
' Funciones anteriores:
Private Declare Function GetPrinterJobs Lib "gsPrinterJobs.dll" _
    (ByVal sDeviceName As String) As Long
Private Declare Function GetNumPages Lib "gsPrinterJobs.dll" _
    (ByVal sDeviceName As String, ByRef NumJobs As Long) As Long

Private Sub cmdInfo_Click()
    Dim s As String
    Dim n As Long
    Const Fallo As Long = -1&
    Dim p As Long
    '
    Dim pp As Long
    '
    Dim iStatus As ePrinterStatus
    Dim sPS As String
    '
    s = cboPrinters.Text
    List1.AddItem s
    '
    '
    Dim nMajor As Long, nMinor As Long
    '
    n = gsPJVersion()
    nMajor = n \ 100
    nMinor = n - nMajor * 100
    List1.AddItem "Versión de gsPrinterJobs.dll: " & CStr(nMajor) & "." & Format$(nMinor, "00")
    '
    ' Ejemplo de GetPrinterStatus:
    '
    'iStatus = GetPrinterStatus(Printer.DeviceName)
    iStatus = GetPrinterStatus(s)
    If iStatus = PRSTAT_FALLO Then
        List1.AddItem "Fallo al llamar a la función GetPrinterStatus"
    Else
        List1.AddItem "iStatus = " & CStr(iStatus)
        sPS = "" '"iStatus = " & CStr(iStatus)
        If iStatus And PRINTER_STATUS_BUSY Then _
            sPS = sPS & " Ocupada"
        If iStatus And PRINTER_STATUS_OFFLINE Then _
            sPS = sPS & " OffLine"
        If iStatus And PRINTER_STATUS_NOT_AVAILABLE Then _
            sPS = sPS & " No disponible"
        If iStatus And PRINTER_STATUS_PAPER_JAM Then _
            sPS = sPS & " Papel atascado"
        If iStatus And PRINTER_STATUS_PRINTING Then _
            sPS = sPS & " Imprimiendo"
        If iStatus And PRINTER_STATUS_WAITING Then _
            sPS = sPS & " Esperando"
        '
        If Len(sPS) Then
            List1.AddItem "Estado de la impresora: " & sPS
        End If
    End If
    '
    '
'    n = GetPrinterJobs(s)
'    If n = Fallo Then
'        List1.AddItem "Fallo al llamar a la función GetPrinterJobs"
'    Else
'        List1.AddItem "Número de trabajos pendientes: " & n
'    End If
'    List1.AddItem ""
'    List1.ListIndex = List1.ListCount - 1
    '
    ' Número de trabajos pendientes y número de páginas             (11/Jun/01)
    ' (aunque el número de páginas no se muestran???)
    '
    n = GetNumPages(s, p)
    If n = Fallo Then
        List1.AddItem "Fallo al llamar a la función GetNumPages"
    Else
        List1.AddItem s
        List1.AddItem "Número de trabajos pendientes: " & p & ", páginas: " & n
    End If
    List1.AddItem ""
    '
    ' Para saber el número de páginas impresas                      (13/Dic/01)
    '
    n = GetNumPagesEx(s, p, pp)
    If n <= Fallo Then
        List1.AddItem "Fallo al llamar a la función GetNumPagesEx" ' (" & n & ")"
    Else
        List1.AddItem s
        List1.AddItem "Número de trabajos pendientes: " & p & ", páginas: " & n & ", impresas: " & pp
    End If
    '
    List1.AddItem ""
    ' Seleccionar el último elemento de la lista
    List1.ListIndex = List1.ListCount - 1
    '
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


