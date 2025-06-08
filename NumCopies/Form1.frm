VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   120
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      Option Explicit

      Private Const NULLPTR = 0&
      ' Constants for DEVMODE
      Private Const CCHDEVICENAME = 32
      Private Const CCHFORMNAME = 32
      ' Constants for DocumentProperties
      Private Const DM_MODIFY = 8
      Private Const DM_COPY = 2
      Private Const DM_IN_BUFFER = DM_MODIFY
      Private Const DM_OUT_BUFFER = DM_COPY
      ' Constants for dmOrientation
      Private Const DMORIENT_PORTRAIT = 1
      Private Const DMORIENT_LANDSCAPE = 2
      ' Constants for dmPrintQuality
      Private Const DMRES_DRAFT = (-1)
      Private Const DMRES_HIGH = (-4)
      Private Const DMRES_LOW = (-2)
      Private Const DMRES_MEDIUM = (-3)
      ' Constants for dmTTOption
      Private Const DMTT_BITMAP = 1
      Private Const DMTT_DOWNLOAD = 2
      Private Const DMTT_DOWNLOAD_OUTLINE = 4
      Private Const DMTT_SUBDEV = 3
      ' Constants for dmColor
      Private Const DMCOLOR_COLOR = 2
      Private Const DMCOLOR_MONOCHROME = 1
      ' Constants for dmCollate
      Private Const DMCOLLATE_FALSE = 0
      Private Const DMCOLLATE_TRUE = 1
      Private Const DM_COLLATE As Long = &H8000
      ' Constants for dmDuplex
      Private Const DM_DUPLEX = &H1000&
      Private Const DMDUP_HORIZONTAL = 3
      Private Const DMDUP_SIMPLEX = 1
      Private Const DMDUP_VERTICAL = 2

      Private Type DEVMODE
          dmDeviceName(1 To CCHDEVICENAME) As Byte
          dmSpecVersion As Integer
          dmDriverVersion As Integer
          dmSize As Integer
          dmDriverExtra As Integer
          dmFields As Long
          dmOrientation As Integer
          dmPaperSize As Integer
          dmPaperLength As Integer
          dmPaperWidth As Integer
          dmScale As Integer
          dmCopies As Integer
          dmDefaultSource As Integer
          dmPrintQuality As Integer
          dmColor As Integer
          dmDuplex As Integer
          dmYResolution As Integer
          dmTTOption As Integer
          dmCollate As Integer
          dmFormName(1 To CCHFORMNAME) As Byte
          dmUnusedPadding As Integer
          dmBitsPerPel As Integer
          dmPelsWidth As Long
          dmPelsHeight As Long
          dmDisplayFlags As Long
          dmDisplayFrequency As Long
                
      End Type

      Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
      "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
      ByVal pDefault As Long) As Long

      Private Declare Function DocumentProperties Lib "winspool.drv" _
      Alias "DocumentPropertiesA" (ByVal hwnd As Long, _
      ByVal hPrinter As Long, ByVal pDeviceName As String, _
      pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) _
      As Long

      Private Declare Function ClosePrinter Lib "winspool.drv" _
      (ByVal hPrinter As Long) As Long

      Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

      Function StripNulls(OriginalStr As String) As String
         If (InStr(OriginalStr, Chr(0)) > 0) Then
            OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
         End If
         StripNulls = Trim(OriginalStr)
      End Function

      Function ByteToString(ByteArray() As Byte) As String
        Dim TempStr As String
        Dim I As Integer

        For I = 1 To CCHDEVICENAME
            TempStr = TempStr & Chr(ByteArray(I))
        Next I
        ByteToString = StripNulls(TempStr)
      End Function

Function GetPrinterSettings(szPrinterName As String, hdc As Long) _
      As Boolean
      Dim hPrinter As Long
      Dim nSize As Long
      Dim pDevMode As DEVMODE
      Dim aDevMode() As Byte
      Dim TempStr As String

        If OpenPrinter(szPrinterName, hPrinter, NULLPTR) <> 0 Then
           nSize = DocumentProperties(NULLPTR, hPrinter, szPrinterName, _
           NULLPTR, NULLPTR, 0)
          If nSize < 1 Then
            GetPrinterSettings = False
            Exit Function
          End If
          ReDim aDevMode(1 To nSize)
          nSize = DocumentProperties(NULLPTR, hPrinter, szPrinterName, _
          aDevMode(1), NULLPTR, DM_OUT_BUFFER)
          If nSize < 0 Then
            GetPrinterSettings = False
            Exit Function
          End If
         Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))


        If pDevMode.dmCopies <> 1 Then
         List1.AddItem "Copies: " & CStr(pDevMode.dmCopies)
         ' Add any other items of interest ...
        End If
         
         Call ClosePrinter(hPrinter)
         GetPrinterSettings = True
      Else
         GetPrinterSettings = False
      End If
      End Function


Private Sub Form_Load()
  Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
      If GetPrinterSettings(Printer.DeviceName, Printer.hdc) = False Then
         List1.AddItem "No Settings Retrieved!"
         MsgBox "Unable to retrieve Printer settings.", , "Failure"
      End If
         List1.AddItem "printer.copies: " & Printer.Copies
         
End Sub
