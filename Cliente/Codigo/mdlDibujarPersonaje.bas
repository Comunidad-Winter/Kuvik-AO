Attribute VB_Name = "mdlDibujarPersonaje"
Sub DibujaPJ(grh As grh, ByVal X As Integer, ByVal Y As Integer, Index As Integer)
On Error Resume Next
Dim iGrhIndex As Integer
If grh.grhindex <= 0 Then Exit Sub
iGrhIndex = GrhData(grh.grhindex).Frames(grh.FrameCounter)

Call GrhRenderToHdc(iGrhIndex, frmCuent.PJ(Index).hdc, X, Y, True)
frmCuent.PJ(Index).Refresh

End Sub

Sub dibujamuerto(Index As Integer)

End Sub

Sub DibujarTodo(ByVal Index As Integer, Body As Integer, Head As Integer, casco As Integer, Shield As Integer, Weapon As Integer, Baned As Integer, Nombre As String, LVL As Integer, Clase As String, Muerto As Integer)

Dim grh As grh
Dim Pos As Integer
Dim loopc As Integer

Dim YBody As Integer
Dim YYY As Integer
Dim XBody As Integer
Dim BBody As Integer

If Clase = 1 Then
Clase = "Ciudadano"
ElseIf Clase = 2 Then
Clase = "Trabajador"
ElseIf Clase = 3 Then
Clase = "Experto en minerales"
ElseIf Clase = 4 Then
Clase = "Minero"
ElseIf Clase = 8 Then
Clase = "Herrero"
ElseIf Clase = 13 Then
Clase = "Experto en uso de madera"
ElseIf Clase = 14 Then
Clase = "Leñador"
ElseIf Clase = 18 Then
Clase = "Carpintero"
ElseIf Clase = 23 Then
Clase = "Pescador"
ElseIf Clase = 27 Then
Clase = "Sastre"
ElseIf Clase = 31 Then
Clase = "Alquimista"
ElseIf Clase = 35 Then
Clase = "Luchador"
ElseIf Clase = 36 Then
Clase = "Con uso de mana"
ElseIf Clase = 37 Then
Clase = "Hechicero"
ElseIf Clase = 38 Then
Clase = "Mago"
ElseIf Clase = 39 Then
Clase = "Nigromante"
ElseIf Clase = 40 Then
Clase = "Orden sagrada"
ElseIf Clase = 41 Then
Clase = "Paladin"
ElseIf Clase = 42 Then
Clase = "Clerigo"
ElseIf Clase = 43 Then
Clase = "Naturalista"
ElseIf Clase = 44 Then
Clase = "Bardo"
ElseIf Clase = 45 Then
Clase = "Druida"
ElseIf Clase = 46 Then
Clase = "Sigiloso"
ElseIf Clase = 47 Then
Clase = "Asesino"
ElseIf Clase = 48 Then
Clase = "Cazador"
ElseIf Clase = 49 Then
Clase = "Sin uso de mana"
ElseIf Clase = 50 Then
Clase = "Arquero"
ElseIf Clase = 51 Then
Clase = "Guerrero"
ElseIf Clase = 52 Then
Clase = "Caballero"
ElseIf Clase = 53 Then
Clase = "Bandido"
ElseIf Clase = 55 Then
Clase = "Pirata"
ElseIf Clase = 56 Then
Clase = "Ladron"
End If

frmCuent.Nombre(Index).Caption = Nombre

frmCuent.Label1(Index).Font = frmMain.Font
frmCuent.Label1(Index).Font = frmMain.Font

frmCuent.Label1(Index).Caption = LVL
frmCuent.Label2(Index).Caption = Clase

XBody = 12
YBody = 15
BBody = 17

If Muerto = 1 Then
    Body = 8
    Head = 500
    arma = 2
    Shield = 2
    Weapon = 2
    XBody = 10
    YBody = 35
    BBody = 16
    Call dibujamuerto(Index)
End If

grh = BodyData(Body).Walk(3)
    
Call DibujaPJ(grh, XBody, YBody, Index)

If Muerto = 0 Then YYY = BodyData(Body).HeadOffset.Y
If Muerto = 1 Then YYY = -9

Pos = YYY + GrhData(GrhData(grh.grhindex).Frames(grh.FrameCounter)).pixelHeight
grh = HeadData(Head).Head(3)
    
Call DibujaPJ(grh, BBody, Pos, Index)

If casco <> 2 And casco > 0 Then
Call DibujaPJ(CascoAnimData(casco).Head(3), BBody, Pos, Index)

End If

If Weapon <> 2 And Weapon > 0 Then
Call DibujaPJ(WeaponAnimData(Weapon).WeaponWalk(3), XBody, BBody, Index)
End If

If Shield <> 2 And Shield > 0 Then
Call DibujaPJ(ShieldAnimData(Shield).ShieldWalk(3), XBody, BBody, Index)
End If

End Sub
