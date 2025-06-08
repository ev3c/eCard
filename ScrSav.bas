Attribute VB_Name = "ScrSav"

Public Declare Function TerminateProcess Lib "kernel32" _
(ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Declare Function Process32First Lib "kernel32" ( _
   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Declare Function Process32Next Lib "kernel32" ( _
   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Declare Function CloseHandle Lib "kernel32" _
   (ByVal Handle As Long) As Long

Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
      ByVal dwProcId As Long) As Long

Declare Function EnumProcesses Lib "psapi.dll" _
   (ByRef lpidProcess As Long, ByVal cb As Long, _
      ByRef cbNeeded As Long) As Long

Declare Function GetModuleFileNameExA Lib "psapi.dll" _
   (ByVal hProcess As Long, ByVal hModule As Long, _
      ByVal ModuleName As String, ByVal nSize As Long) As Long

Declare Function EnumProcessModules Lib "psapi.dll" _
   (ByVal hProcess As Long, ByRef lphModule As Long, _
      ByVal cb As Long, ByRef cbNeeded As Long) As Long

Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
   ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Declare Function GetVersionExA Lib "kernel32" _
   (lpVersionInformation As OSVERSIONINFO) As Integer

Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long           ' This process
   th32DefaultHeapID As Long
   th32ModuleID As Long            ' Associated exe
   cntThreads As Long
   th32ParentProcessID As Long     ' This process's parent process
   pcPriClassBase As Long          ' Base priority of process threads
   dwFlags As Long
   szExeFile As String * 260       ' MAX_PATH
End Type

Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long           '1 = Windows 95.
                                  '2 = Windows NT
   szCSDVersion As String * 128
End Type

Const PROCESS_QUERY_INFORMATION = 1024
Const PROCESS_VM_READ = 16
Const MAX_PATH = 260
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS = &H2&
Const hNull = 0

Function StrZToStr(s As String) As String
   StrZToStr = Left$(s, InStr(s, Chr(0)) - 1)
End Function

Public Function getVersion() As Long
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   getVersion = osinfo.dwPlatformId
End Function

Public Sub GetWinFiles()
  Dim iFichero As Integer

  On Error GoTo GetWinFiles_Err

  ReDim gaPrgAct(1 To 99) As ecPath
  iFichero = 1

  Select Case getVersion()
  
   Case 1 'Windows 95/98
  
     Dim f As Long, sname As String
     Dim hSnap As Long, proc As PROCESSENTRY32
     Dim hProc As Long
     Dim lRetX As Long
     hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
     If hSnap = hNull Then Exit Sub
     proc.dwSize = Len(proc)
     
     ' Iterate through the processes
     f = Process32First(hSnap, proc)
     
     Do While f
       sname = StrZToStr(proc.szExeFile)
       gaPrgAct(iFichero).Path = sname
       iFichero = iFichero + 1
       f = Process32Next(hSnap, proc)
     Loop
     
     lRetX = CloseHandle(hSnap)
    
   Case Else 'Windows NT
     
     Dim cb As Long
     Dim cbNeeded As Long
     Dim NumElements As Long
     Dim ProcessIDs() As Long
     Dim cbNeeded2 As Long
     Dim NumElements2 As Long
     Dim Modules(1 To 200) As Long
     Dim lRet As Long
     Dim ModuleName As String
     Dim nSize As Long
     Dim hProcess, hProcess2 As Long
     Dim i As Long
     'Get the array containing the process id's for each process  object
     cb = 8
     cbNeeded = 96
     Do While cb <= cbNeeded
        cb = cb * 2
        ReDim ProcessIDs(cb / 4) As Long
        lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
     Loop
     NumElements = cbNeeded / 4
  
     For i = 1 To NumElements
        'Get a handle to the Process
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
           Or PROCESS_VM_READ, 0, ProcessIDs(i))
        'Got a Process handle
        If hProcess <> 0 Then
            'Get an array of the module handles for the specified
            'process
            lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                         cbNeeded2)
            'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
               ModuleName = Space(MAX_PATH)
               nSize = 500
               lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                               ModuleName, nSize)
               sname = Left(ModuleName, lRet)
               'sname = StrZToStr(sname)
              gaPrgAct(iFichero).Path = sname
              iFichero = iFichero + 1
            End If
        End If
      'Close the handle to the process
     lRet = CloseHandle(hProcess)
     Next
  
  End Select
  
  gaPrgAct(iFichero).Path = ""
  
GetWinFiles_Err:

End Sub

Public Sub Programa_Apagar(sPath As String)

  On Error GoTo Programa_Apagar_Err


  Select Case getVersion()
  
   Case 1 'Windows 95/98
  
     Dim f As Long, sname As String
     Dim hSnap As Long, proc As PROCESSENTRY32
     Dim hProc As Long
     hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
     If hSnap = hNull Then Exit Sub
     proc.dwSize = Len(proc)
     
     ' Iterate through the processes
     f = Process32First(hSnap, proc)
     
     Do While f
       sname = StrZToStr(proc.szExeFile)
       hProc = OpenProcess(0, 0, proc.th32ProcessID)
       If UCase(sPath) = UCase(sname) Then
         lRetX = TerminateProcess(hProc, 0&)
       End If
       f = Process32Next(hSnap, proc)
     Loop
    
     lRetX = CloseHandle(hSnap)
     lRetX = CloseHandle(hProc)
    
   Case Else 'Windows NT
     
     Dim cb As Long
     Dim cbNeeded As Long
     Dim NumElements As Long
     Dim ProcessIDs() As Long
     Dim cbNeeded2 As Long
     Dim NumElements2 As Long
     Dim Modules(1 To 200) As Long
     Dim lRet As Long
     Dim ModuleName As String
     Dim nSize As Long
     Dim hProcess, hProcess2 As Long
     Dim i As Long
     'Get the array containing the process id's for each process  object
     cb = 8
     cbNeeded = 96
     Do While cb <= cbNeeded
        cb = cb * 2
        ReDim ProcessIDs(cb / 4) As Long
        lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
     Loop
     NumElements = cbNeeded / 4
  
     For i = 1 To NumElements
        'Get a handle to the Process
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
           Or PROCESS_VM_READ, 0, ProcessIDs(i))
        'Got a Process handle
        If hProcess <> 0 Then
            'Get an array of the module handles for the specified
            'process
            lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                         cbNeeded2)
            'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
               ModuleName = Space(MAX_PATH)
               nSize = 500
               lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                               ModuleName, nSize)
               sname = Left(ModuleName, lRet)
               'sname = StrZToStr(sname)
               hProcess2 = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessIDs(i))
              If UCase(sPath) = UCase(sname) Then
                lRet = TerminateProcess(hProcess2, 0&)
              End If
            End If
        End If
      'Close the handle to the process
     lRet = CloseHandle(hProcess)
     lRet = CloseHandle(hProcess2)
     Next
  
  End Select
  
Programa_Apagar_Err:

End Sub

