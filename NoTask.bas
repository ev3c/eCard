Attribute VB_Name = "NoTask"
'Author: Tony Jones
'Sometimes you need your app to be 'impossible' to shutdown, for
'example if you are making some kind of security program. Even if
'you don't provide a direct way to shutdown your program via any
'Exit commands, the user still can close the prog via Ctrl-Alt-Del.
'This code takes care of this by hiding your program from the so
'called Task List.
'NOTE: This code will *NOT* work in NT.

'Put the following code in a bas module
Public Declare Function RegisterServiceProcess _
       Lib "kernel32" (ByVal ProcessID As Long, _
       ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId _
       Lib "kernel32" () As Long




