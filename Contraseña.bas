Attribute VB_Name = "Contraseña"
Public Function Contraseña_Encriptar(sContraseña As String) As String
  Dim sTemp, sChr As String
  
  For x = 1 To Len(sContraseña)
     sChr = Mid$(sContraseña, x, 1)
     sTemp = sTemp + Chr(Asc(sChr) + 5 + x)
  Next x
    
  Contraseña_Encriptar = sTemp
End Function

Public Function Contraseña_DesEncriptar(sContraseña As String) As String
  Dim sTemp, sChr As String
  
  For x = 1 To Len(sContraseña)
     sChr = Mid$(sContraseña, x, 1)
     sTemp = sTemp + Chr(Asc(sChr) - 5 - x)
  Next x
    
  Contraseña_DesEncriptar = sTemp
End Function

