Attribute VB_Name = "Contraseņa"
Public Function Contraseņa_Encriptar(sContraseņa As String) As String
  Dim sTemp, sChr As String
  
  For x = 1 To Len(sContraseņa)
     sChr = Mid$(sContraseņa, x, 1)
     sTemp = sTemp + Chr(Asc(sChr) + 5 + x)
  Next x
    
  Contraseņa_Encriptar = sTemp
End Function

Public Function Contraseņa_DesEncriptar(sContraseņa As String) As String
  Dim sTemp, sChr As String
  
  For x = 1 To Len(sContraseņa)
     sChr = Mid$(sContraseņa, x, 1)
     sTemp = sTemp + Chr(Asc(sChr) - 5 - x)
  Next x
    
  Contraseņa_DesEncriptar = sTemp
End Function

