Attribute VB_Name = "Contrase�a"
Public Function Contrase�a_Encriptar(sContrase�a As String) As String
  Dim sTemp, sChr As String
  
  For x = 1 To Len(sContrase�a)
     sChr = Mid$(sContrase�a, x, 1)
     sTemp = sTemp + Chr(Asc(sChr) + 5 + x)
  Next x
    
  Contrase�a_Encriptar = sTemp
End Function

Public Function Contrase�a_DesEncriptar(sContrase�a As String) As String
  Dim sTemp, sChr As String
  
  For x = 1 To Len(sContrase�a)
     sChr = Mid$(sContrase�a, x, 1)
     sTemp = sTemp + Chr(Asc(sChr) - 5 - x)
  Next x
    
  Contrase�a_DesEncriptar = sTemp
End Function

