Attribute VB_Name = "Matematicas"
'================================
'>>>>> WWW.FADICTOS.COM.AR <<<<<<
'================================
'Matute - matius_xd07@hotmail.com

Option Explicit
Public Function Porcentaje(Total As Variant, Porc As Variant) As Long

Porcentaje = Total * (Porc / 100)

End Function
Sub RestVar(Var As Variant, Take As Variant, MIN As Variant)

Var = Maximo(Var - Take, MIN)

End Sub
Sub AddtoVar(Var As Variant, Addon As Variant, MAX As Variant)

Var = Minimo(Var + Addon, MAX)

End Sub
Function Distancia(wp1 As WorldPos, wp2 As WorldPos)

Distancia = (Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100))

End Function
Function TipoClase(UserIndex As Integer) As Byte

Select Case UserList(UserIndex).Clase
    Case PALADIN, ASESINO, CAZADOR
        TipoClase = 2
    Case CLERIGO, BARDO, LADRON
        TipoClase = 3
    Case MAGO, NIGROMANTE, DRUIDA
        TipoClase = 4
    Case Else
        TipoClase = 1
End Select

End Function
Public Function TipoRaza(UserIndex As Integer) As Byte

If UserList(UserIndex).Raza = ENANO Or UserList(UserIndex).Raza = GNOMO Then
    TipoRaza = 2
Else: TipoRaza = 1
End If

End Function
Public Function RazaBaja(UserIndex As Integer) As Boolean

RazaBaja = (UserList(UserIndex).Raza = ENANO Or UserList(UserIndex).Raza = GNOMO)

End Function
Function Distance(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) As Double

Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function
