Attribute VB_Name = "Declaraciones"

' Database y recorsets de eCard.mdb
Public gdb As Database
Public grsPrograma As Recordset
Public grsEvento As Recordset
Public grsPrg As Recordset
Public grsLst As Recordset

Public Type ecOnPathPvp
  on As Date        'Hora de arranque
  path  As String   'path del programa
  pvp As Integer    'Precio hora
  level As Integer  'Nivel de acceso del programa
  card As Integer   'N�mero de Tarjeta
End Type

Public Type ecPath
  path As String    'path del programa
End Type

Public gaPrg() As ecOnPathPvp     'programas grabados para control
Public gaPrgAct() As ecPath   'programas activos


Public gstrPrograma As String     'Nombre y Versi�n de eCard
Public gstrContrase�a As String   'Contrase�a
Public gblnContrase�a As Boolean  'Contrase�a correcta
Public gstrFormatoFecha As String 'Formato de la fecha

Public MyDate As Date         'Fecha interna
Public gdFechaOn As Date      'Fecha de arranque

Public giComm As Integer      'Puerto de Comunicaciones
