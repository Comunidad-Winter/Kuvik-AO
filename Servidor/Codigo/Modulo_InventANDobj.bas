Attribute VB_Name = "InvNpc"
'================================
'>>>>> WWW.FADICTOS.COM.AR <<<<<<
'================================
'Matute - matius_xd07@hotmail.com

Option Explicit
Public Function TirarItemAlPiso(POS As WorldPos, Obj As Obj) As WorldPos
On Error GoTo errhandler
Dim NuevaPos As WorldPos

Call Tilelibre(POS, NuevaPos)

If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
      Call MakeObj(ToMap, 0, POS.Map, _
      Obj, POS.Map, NuevaPos.X, NuevaPos.Y)
      TirarItemAlPiso = NuevaPos
End If

Exit Function
errhandler:

End Function
Public Sub NPC_TIRAR_ITEMS(MiNPC As Npc, UserIndex As Integer)

On Error Resume Next

If MiNPC.Invent.NroItems Then
    
    Dim i As Byte
    Dim MiObj As Obj
    Dim Prob As Integer
    
    For i = 1 To MAX_NPCINVENTORY_SLOTS
 If MiNPC.Probabilidad = 0 Then
        If MiNPC.Invent.Object(i).OBJIndex Then
              If val(MiNPC.MaxRecom) Then
              MiObj.Amount = RandomNumber(MiNPC.MinRecom, MiNPC.MaxRecom)
              Else
              MiObj.Amount = MiNPC.Invent.Object(i).Amount
              End If
              MiObj.OBJIndex = MiNPC.Invent.Object(i).OBJIndex
              
              If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)
              
              
        End If
        Else
        Prob = RandomNumber(0, 100)
        If Prob <= MiNPC.Probabilidad Then
                If MiNPC.Invent.Object(i).OBJIndex Then
              If MiNPC.MaxRecom Then
              MiObj.Amount = RandomNumber(MiNPC.MinRecom, MiNPC.MaxRecom)
              Else
              MiObj.Amount = MiNPC.Invent.Object(i).Amount
              End If
              MiObj.OBJIndex = MiNPC.Invent.Object(i).OBJIndex
              
              If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)
              End If
              Call UpdateUserInv(True, UserIndex, 0)
              
              
        End If
        End If
      End If
    Next

End If

End Sub
Function QuedanItems(NpcIndex As Integer, OBJIndex As Integer) As Boolean
On Error Resume Next
Dim i As Integer

If Npclist(NpcIndex).Invent.NroItems Then
    For i = 1 To MAX_NPCINVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(i).OBJIndex = OBJIndex Then
            QuedanItems = True
            Exit Function
        End If
    Next
End If

End Function
Sub ResetNpcInv(NpcIndex As Integer)
On Error Resume Next

Dim i As Integer

Npclist(NpcIndex).Invent.NroItems = 0

For i = 1 To MAX_NPCINVENTORY_SLOTS
   Npclist(NpcIndex).Invent.Object(i).OBJIndex = 0
   Npclist(NpcIndex).Invent.Object(i).Amount = 0
Next

Npclist(NpcIndex).InvReSpawn = 0

End Sub
Sub QuitarNpcInvItem(NpcIndex As Integer, Slot As Byte, Cantidad As Integer, UserIndex As Integer)
Dim OBJIndex As Integer

OBJIndex = Npclist(NpcIndex).Invent.Object(Slot).OBJIndex

If Npclist(NpcIndex).InvReSpawn = 1 Then
    Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
    If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
        Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
        Npclist(NpcIndex).Invent.Object(Slot).OBJIndex = 0
        Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
        If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
           Call CargarInvent(NpcIndex)
        End If
    End If
    Call UpdateNPCInv(False, UserIndex, NpcIndex, Slot)
End If

End Sub
Sub CargarInvent(NpcIndex As Integer)
Dim LoopC As Integer, ln As String, npcfile As String

If Npclist(NpcIndex).Numero >= 500 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else: npcfile = DatPath & "NPCs.dat"
End If

Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
Next

End Sub
