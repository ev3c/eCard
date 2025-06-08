Attribute VB_Name = "HoraSuma"
Function Hora_Suma(dtArg As Date) As String
Dim Horas As Double
Dim Minutos As Double
Dim RetBuffer As String

    Horas = dtArg * 24
    Minutos = (Horas - Int(Horas)) * 60
    Horas = Int(Horas)
    RetBuffer = Format(Horas, "0") & ":" & Format(Minutos, "00")
    Hora_Suma = RetBuffer
    ' Se puede asignar la expresion anterior a
    ' Hora_Suma.(en lo personal, no me gusta)
End Function
