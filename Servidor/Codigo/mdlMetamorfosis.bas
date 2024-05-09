Attribute VB_Name = "modMetamorfosis"
'================================
'>>>>> WWW.FADICTOS.COM.AR <<<<<<
'================================
'Matute - matius_xd07@hotmail.com

Sub DoMetamorfosis(UserIndex As Integer)

metacuerpo = RandomNumber(1, 10)

Select Case (metacuerpo)
    Case 1
        metacuerpo = 9
    Case 2
        metacuerpo = 11
    Case 3
        metacuerpo = 42
    Case 4
        metacuerpo = 243
    Case 5
        metacuerpo = 149
    Case 6
        metacuerpo = 151
    Case 7
        metacuerpo = 155
    Case 8
        metacuerpo = 157
    Case 9
        metacuerpo = 159
    Case 10
        metacuerpo = 141
End Select

UserList(UserIndex).flags.Transformado = 1
UserList(UserIndex).Counters.Transformado = Timer

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, val(UserIndex), metacuerpo, 0, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
        
If UserList(UserIndex).flags.Desnudo Then UserList(UserIndex).flags.Desnudo = 0

Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_MORPH)
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARPMORPH & "," & 0)

End Sub
Sub DoTransformar(UserIndex As Integer, Optional ByVal FX As Boolean = True)

UserList(UserIndex).flags.Transformado = 0
UserList(UserIndex).Counters.Transformado = 0

If UserList(UserIndex).Invent.ArmourEqpObjIndex = 0 Then
    Call DarCuerpoDesnudo(UserIndex)
Else
    UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
End If

If UserList(UserIndex).Invent.CascoEqpObjIndex = 0 Then
   UserList(UserIndex).Char.CascoAnim = NingunCasco
Else
    UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
End If

If UserList(UserIndex).Invent.EscudoEqpObjIndex = 0 Then
   UserList(UserIndex).Char.ShieldAnim = NingunEscudo
Else
    UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
End If

If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
   UserList(UserIndex).Char.WeaponAnim = NingunArma
Else
    UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
End If

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

If FX Then
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_WARPMORPH)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARPMORPH & "," & 0)
End If

End Sub
