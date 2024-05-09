Attribute VB_Name = "Mod_TileEngine"
Option Explicit
Option Base 0

Public bRunning As Boolean

Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

Public Const SRCCOPY = &HCC0020

Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type GrhData
    sX          As Integer
    sY          As Integer
    FileNum     As Integer
    pixelWidth  As Integer
    pixelHeight As Integer
    TileWidth   As Single
    TileHeight  As Single
   
    NumFrames       As Integer
    Frames(1 To 25) As Integer
    speed           As Single
End Type
 
Public Type grh
    Loops        As Integer
    GrhIndex     As Integer
    FrameCounter As Single
    SpeedCounter As Single
    Started      As Byte
End Type

Public Type BodyData
    Walk(1 To 4) As grh
    HeadOffset As Position
End Type

Public Type HeadData
    Head(1 To 4) As grh
End Type

Type WeaponAnimData
    WeaponWalk(1 To 4) As grh
End Type

Type ShieldAnimData
    ShieldWalk(1 To 4) As grh
End Type


Public Type FxData
    FX As grh
    OffsetX As Long
    OffsetY As Long
End Type

Type position2
    X As Single
    Y As Single
End Type

Public Type Char
    active As Byte
    Heading As Byte
    Pos As Position

    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    FX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    Navegando As Byte
    
    Nombre As String
    GM As Integer
    
    haciendoataque As Byte
    Moving As Byte
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    MoveOffset As position2
    ServerIndex As Integer
    
    pie As Boolean
    Muerto As Boolean
    invisible As Boolean
    
End Type

Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public Type MapBlock
    Graphic(1 To 4) As grh
    CharIndex As Integer
    ObjGrh As grh

    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    light_value(3) As Long
   
    luz As Integer
    color(3) As Long
   
    particle_group As Integer
End Type

Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    Changed As Byte
End Type

Public IniPath As String
Public MapPath As String

Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

Public MainViewTop As Integer
Public MainViewLeft As Integer

Public TileBufferSize As Integer

Public DisplayFormhWnd As Long

Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Public LastTime As Long

Public MainDestRect   As RECT

Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer

Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public grh() As grh

Public MapData() As MapBlock
Public MapInfo As MapInfo

Public CharList(1 To 10000) As Char

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public bRain        As Boolean
Public bRainST      As Boolean
Public bTecho       As Boolean
Public brstTick     As Long

Private RLluvia(7)  As RECT
Private iFrameIndex As Byte
Private llTick      As Long
Private LTLluvia(4) As Integer
            
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

Public Enum PlayLoop
        plNone = 0
        plLluviain = 1
        plLluviaout = 2
        plFogata = 3
End Enum
    
'//////////VARIABLES DIRECTX8////////// 'THUSING
Dim bump_map_texture As Direct3DTexture8
Dim bump_map_texture_ex As Direct3DTexture8
Dim bump_map_supported As Boolean
Dim bump_map_powa As Boolean
 
Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2
 
Private Const MAX_DIALOGOS = 300
Private Const MAXLONG = 15
 
Public Const PI As Single = 3.14159265358979

'Cargas de texto desde GRH
Private Type tFont
    Caracteres(0 To 255) As Integer 'indice de cada letra
End Type
 
Private Fuentes(1) As tFont

Public font_list() As D3DXFont

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Const HASH_TABLE_SIZE As Long = 337
 
Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type
 
Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type
 
Private TexList(HASH_TABLE_SIZE - 1) As HashNode
 
Private lFrameLimiter As Long
Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public timerTicksPerFrame As Single
Public timerElapsedTime As Single
Public particletimer As Single
Public engineBaseSpeed As Single
 Public ScrollPixelsPerFrame As Single
 
'Describes a transformable lit vertex
Private Type TLVERTEX
  X As Single
  Y As Single
  Z As Single
  rhw As Single
  color As Long
  Specular As Long
  tu As Single
  tv As Single
End Type
 
'********** Direct X ***********
Private Type D3D8Textures
    texture As Direct3DTexture8
    texwidth As Integer
    texheight As Integer
End Type
 
Private Type D3D8Textures2
    texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type
 
'DirectX 8 Objects
Public Dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

Public D3DDeviceDam As Direct3DDevice8
'Font List
Public FontList As D3DXFont
 
Private Type tLight
    RGBcolor As D3DCOLORVALUE
    active As Boolean
    map_x As Byte
    map_y As Byte
    range As Byte
    id As Long
End Type
 
Private light_list() As tLight
Private NumLights As Byte
Dim light_count As Long
Dim light_last As Long
 
Public mFreeMemoryBytes As Long
 
Private pUdtMemStatus As MEMORYSTATUS
 
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
 
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
 
Public Base_Light As Long
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte
Type luzxhora
    R As Long
    G As Long
    b As Long
End Type
Public luz_dia(0 To 24) As luzxhora
 
Public Const ImgSize As Byte = 4
 
Private Type tCache
    Number        As Long
    SrcHeight     As Single
    SrcWidth      As Single
End Type: Private Cache As tCache
 
'BitBlt
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
 
'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Sub CargarCabezas()
On Error Resume Next
Dim n As Integer, i As Integer, Numheads As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , Numheads


ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #n, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub
Sub CargarCascos()
On Error Resume Next
Dim n As Integer, i As Integer, NumCascos As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCascos


ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #n, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

n = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCuerpos


ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #n, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
Next i

Close #n

End Sub
Sub CargarFxs()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

n = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumFxs


ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #n, , MisFxs(i)
    Call InitGrh(FxData(i).FX, MisFxs(i).Animacion, 1)
    FxData(i).OffsetX = MisFxs(i).OffsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #n

End Sub
Sub CargarArrayLluvia()
'On Error Resume Next
Dim n As Integer, i As Integer
Dim Nu As Integer
 
n = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #n
 
 
Get #n, , MiCabecera
 
 
Get #n, , Nu
 
 
ReDim bLluvia(1 To 230) As Byte
 
For i = 1 To 230
    Get #n, , bLluvia(i)
Next i
 
Close #n
 
End Sub
Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.renderer.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmMain.renderer.ScaleHeight \ 64
End Sub
Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer)
On Error Resume Next


If CharIndex > LastChar Then LastChar = CharIndex

NumChars = NumChars + 1

If arma = 0 Then arma = 2
If escudo = 0 Then escudo = 2
If casco = 0 Then casco = 2

CharList(CharIndex).Head = HeadData(Head)

CharList(CharIndex).Body = BodyData(Body)

If Body > 83 And Body < 88 Then
    CharList(CharIndex).Navegando = 1
Else: CharList(CharIndex).Navegando = 0
End If

CharList(CharIndex).arma = WeaponAnimData(arma)
    
CharList(CharIndex).escudo = ShieldAnimData(escudo)
CharList(CharIndex).casco = CascoAnimData(casco)

CharList(CharIndex).Heading = Heading


CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0


CharList(CharIndex).Pos.X = X
CharList(CharIndex).Pos.Y = Y


CharList(CharIndex).active = 1


MapData(X, Y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

CharList(CharIndex).active = 0
CharList(CharIndex).Criminal = 0
CharList(CharIndex).FX = 0
CharList(CharIndex).FxLoopTimes = 0
CharList(CharIndex).invisible = False
CharList(CharIndex).Moving = 0
CharList(CharIndex).Muerto = False
CharList(CharIndex).Nombre = ""
CharList(CharIndex).pie = False
CharList(CharIndex).Pos.X = 0
CharList(CharIndex).Pos.Y = 0
CharList(CharIndex).UsandoArma = False

End Sub

Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next





CharList(CharIndex).active = 0


If CharIndex = LastChar Then
    Do Until CharList(LastChar).active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

Call ResetCharInfo(CharIndex)


NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef grh As grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)



If GrhIndex = 0 Then Exit Sub
grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(grh.GrhIndex).NumFrames > 1 Then
        grh.Started = 1
    Else
        grh.Started = 0
    End If
Else
    grh.Started = Started
End If

grh.FrameCounter = 1

If grh.GrhIndex <> 0 Then grh.SpeedCounter = GrhData(grh.GrhIndex).speed

End Sub


Public Sub DoFogataFx()
If FX = 0 Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then Audio.StopWave
    Else
        bFogata = HayFogata()
        If bFogata Then Audio.PlayWave "fuego.wav", True
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim X As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next X
Next Y

EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, cont As Integer) As Boolean
Static TickCount(200) As Integer
If cont = 999 Then Exit Function
TickCount(Cual) = TickCount(Cual) + 1
If TickCount(Cual) < cont Then
    TickON = False
Else
    TickCount(Cual) = 0
    TickON = True
End If
End Function
Sub DoPasosFx(ByVal CharIndex As Integer)
Static pie As Boolean

If CharList(CharIndex).Navegando = 0 Then
    If UserMontando And EstaPCarea(CharIndex) And CharIndex = UserCharIndex Then
        If TickON(0, 4) Then Call Audio.PlayWave(SND_MONTANDO)
    Else
        If CharList(CharIndex).Criminal = 1 Then Exit Sub
        If Not CharList(CharIndex).Muerto And EstaPCarea(CharIndex) Then
            CharList(CharIndex).pie = Not CharList(CharIndex).pie
            If CharList(CharIndex).pie Then
                Call Audio.PlayWave(SND_PASOS1)
            Else
                Call Audio.PlayWave(SND_PASOS2)
            End If
        End If
    End If
Else: Call Audio.PlayWave(SND_NAVEGANDO)
End If

End Sub

Sub MoveCharByPosAndHead(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
 
On Error Resume Next
 
Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
 
 
 
X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y
 
MapData(X, Y).CharIndex = 0
 
addX = nX - X
addY = nY - Y
 
MapData(nX, nY).CharIndex = CharIndex
 
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
 
CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)
 
CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading
 
CharList(CharIndex).scrollDirectionX = Sgn(addX)
CharList(CharIndex).scrollDirectionY = Sgn(addY)
 
'.MoveOffset.X = -1 * (32 * addX)
'.MoveOffset.Y = -1 * (32 * addY)
 
'.Moving = 1
'.Heading = nheading
 
'.scrollDirectionX = addX
'.scrollDirectionY = addY
 
 
End Sub
Sub MoveCharByPos(CharIndex As Integer, nX As Integer, nY As Integer)
On Error Resume Next
 
Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nheading As Byte
 
X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y
 
'MapData(X, y).CharIndex = 0
 
addX = nX - X
addY = nY - Y
 
If Sgn(addX) = -1 Then nheading = WEST
If Sgn(addX) = 1 Then nheading = EAST
 
If Sgn(addY) = -1 Then nheading = NORTH
If Sgn(addY) = 1 Then nheading = SOUTH
 
'MapData(nX, nY).CharIndex = CharIndex
 
'CharList(CharIndex).POS.X = nX
'CharList(CharIndex).POS.y = nY
 
'CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
'CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)
 
'CharList(CharIndex).Moving = 1
'CharList(CharIndex).Heading = nheading
 
'CharList(CharIndex).scrollDirectionX = Sgn(addX)
'CharList(CharIndex).scrollDirectionY = Sgn(addY)
 
MoveCharByHead CharIndex, nheading
 
 
 
End Sub
Sub MoveCharByPosConHeading(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
On Error Resume Next
 
If InMapBounds(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y) Then MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0
 
MapData(nX, nY).CharIndex = CharIndex
 
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
 
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0
 
CharList(CharIndex).Heading = nheading
 
End Sub
 
Sub MoveCharByHead(CharIndex As Integer, nheading As Byte)
 
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer
 
With CharList(CharIndex)
X = .Pos.X
Y = .Pos.Y
 
 
Select Case nheading
 
    Case NORTH
        addY = -1
 
    Case EAST
        addX = 1
 
    Case SOUTH
        addY = 1
   
    Case WEST
        addX = -1
       
End Select
 
nX = X + addX
nY = Y + addY
 
MapData(nX, nY).CharIndex = CharIndex
.Pos.X = nX
.Pos.Y = nY
MapData(X, Y).CharIndex = 0
 
.MoveOffset.X = -1 * (32 * addX)
.MoveOffset.Y = -1 * (32 * addY)
 
.Moving = 1
.Heading = nheading
 
.scrollDirectionX = addX
.scrollDirectionY = addY
 
If UserEstado <> True Then Call DoPasosFx(CharIndex)
 
 
End With
 
End Sub
Sub MoveScreen(Heading As Byte)


Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer


Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select


tX = UserPos.X + X
tY = UserPos.Y + Y


If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.X - 8 To UserPos.X + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
Sub LoadGrhData()
On Error GoTo ErrorHandler
 
Dim grh As Integer
Dim Frame As Integer
Dim tempint As Integer
 
 
ReDim GrhData(1 To 32000) As GrhData
 
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1
 
Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
 
Get #1, , grh
 
Do Until grh <= 0
   
    Get #1, , GrhData(grh).NumFrames
    If GrhData(grh).NumFrames <= 0 Then GoTo ErrorHandler
   
    If GrhData(grh).NumFrames > 1 Then
   
       
        For Frame = 1 To GrhData(grh).NumFrames
       
            Get #1, , GrhData(grh).Frames(Frame)
            If GrhData(grh).Frames(Frame) <= 0 Or GrhData(grh).Frames(Frame) > 32000 Then
                GoTo ErrorHandler
            End If
       
        Next Frame
    Dim a As Integer
   
        Get #1, , a
       
        GrhData(grh).speed = a
       
        ÒoÒal grh
       
       
        If GrhData(grh).speed <= 0 Then GoTo ErrorHandler
       
       
        GrhData(grh).pixelHeight = GrhData(GrhData(grh).Frames(1)).pixelHeight
        If GrhData(grh).pixelHeight <= 0 Then GoTo ErrorHandler
       
        GrhData(grh).pixelWidth = GrhData(GrhData(grh).Frames(1)).pixelWidth
        If GrhData(grh).pixelWidth <= 0 Then GoTo ErrorHandler
       
        GrhData(grh).TileWidth = GrhData(GrhData(grh).Frames(1)).TileWidth
        If GrhData(grh).TileWidth <= 0 Then GoTo ErrorHandler
       
        GrhData(grh).TileHeight = GrhData(GrhData(grh).Frames(1)).TileHeight
        If GrhData(grh).TileHeight <= 0 Then GoTo ErrorHandler
   
    Else
   
       
        Get #1, , GrhData(grh).FileNum
        If GrhData(grh).FileNum <= 0 Then GoTo ErrorHandler
       
        Get #1, , GrhData(grh).sX
        If GrhData(grh).sX < 0 Then GoTo ErrorHandler
       
        Get #1, , GrhData(grh).sY
        If GrhData(grh).sY < 0 Then GoTo ErrorHandler
           
        Get #1, , GrhData(grh).pixelWidth
        If GrhData(grh).pixelWidth <= 0 Then GoTo ErrorHandler
       
        Get #1, , GrhData(grh).pixelHeight
        If GrhData(grh).pixelHeight <= 0 Then GoTo ErrorHandler
       
       
        GrhData(grh).TileWidth = GrhData(grh).pixelWidth / TilePixelHeight
        GrhData(grh).TileHeight = GrhData(grh).pixelHeight / TilePixelWidth
       
        GrhData(grh).Frames(1) = grh
           
    End If
   
   
    Get #1, , grh
   
 
Loop
 
 
Close #1
 
 
Exit Sub
 
ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & grh
 
End Sub
 
Sub ÒoÒal(grh As Integer)
 
GrhData(grh).speed = ((GrhData(grh).speed * 1000) / 18)
 
End Sub
Function LegalPos(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function

Function LegalPosMuerto(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPosMuerto = False
    Exit Function
End If

    
    If MapData(X, Y).Blocked = 1 Then
        LegalPosMuerto = False
        Exit Function
    End If
    
    
    If MapData(X, Y).CharIndex > 0 Then
    If CharList(MapData(X, Y).CharIndex).Muerto = True Then
        LegalPosMuerto = False
        Exit Function
    End If
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    End If
    
LegalPosMuerto = True

End Function




Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(X As Integer, Y As Integer) As Boolean




If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function
Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function
Sub PlayWaveAPI(file As String)
Dim rc As Integer

rc = sndPlaySound(file, SND_ASYNC)

End Sub
Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan MartÌn Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
On Error Resume Next 'Thusing
   
    Dim Y                   As Integer     'Keeps track of where on map we are
    Dim X                   As Integer     'Keeps track of where on map we are
    Dim screenminY          As Integer  'Start Y pos on current screen
    Dim screenmaxY          As Integer  'End Y pos on current screen
    Dim screenminX          As Integer  'Start X pos on current screen
    Dim screenmaxX          As Integer  'End X pos on current screen
    Dim minY                As Integer  'Start Y pos on current map
    Dim maxY                As Integer  'End Y pos on current map
    Dim minX                As Integer  'Start X pos on current map
    Dim maxX                As Integer  'End X pos on current map
    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer 'For centering grhs
    Dim PixelOffsetYTemp    As Integer 'For centering grhs
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer
    Dim TempChar As Char
    Dim Moved    As Byte
    Dim iPPx     As Integer
    Dim iPPy     As Integer
   
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
   
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
   
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
   
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
   
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
   
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
   
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
   
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
   
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
   
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
   
                Dim Blanco(3) As Long
                Blanco(0) = RGB(255, 255, 255)
                Blanco(1) = RGB(255, 255, 255)
                Blanco(2) = RGB(255, 255, 255)
                Blanco(3) = RGB(255, 255, 255)
   
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Layer 1 **********************************
        
#If HARDCODED = True Then
                    If MapData(X, Y).Graphic(1).Started = 1 Then
                        MapData(X, Y).Graphic(1).FrameCounter = MapData(X, Y).Graphic(1).FrameCounter + (timerElapsedTime * GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames / MapData(X, Y).Graphic(1).speed)
                        If MapData(X, Y).Graphic(1).FrameCounter > GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames Then
                            MapData(X, Y).Graphic(1).FrameCounter = (MapData(X, Y).Graphic(1).FrameCounter Mod GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames) + 1
                           
                            If MapData(X, Y).Graphic(1).Loops <> -1 Then
                                If MapData(X, Y).Graphic(1).Loops > 0 Then
                                    MapData(X, Y).Graphic(1).Loops = MapData(X, Y).Graphic(1).Loops - 1
                                Else
                                    MapData(X, Y).Graphic(1).Started = 0
                                End If
                            End If
                        End If
                    End If
 
                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(1).GrhIndex).Frames(MapData(X, Y).Graphic(1).FrameCounter)
 
                Device_Box_Textured_Render CurrentGrhIndex, _
                    (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, _
                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                    MapData(X, Y).light_value, _
                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                    False _
                    , 0
#Else
                Call Draw_Grh(MapData(X, Y).Graphic(1), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        0, 1, MapData(X, Y).light_value(), , , , , X, Y)
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value(), , , , , X, Y)
                End If
#End If
            '******************************************
            ScreenX = ScreenX + 1
        Next X
 
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
   
    'Draw floor layer 2
'    ScreenY = minYOffset
'    For Y = screenminY To screenmaxY
'        ScreenX = minXOffset
'        For X = screenminX To screenmaxX
                'Layer 2 **********************************
'                If MapData(X, Y).Graphic(2).grhindex <> 0 Then
'#If HARDCODED = True Then
'                    If MapData(X, Y).Graphic(2).Started = 1 Then
'                        MapData(X, Y).Graphic(2).FrameCounter = MapData(X, Y).Graphic(2).FrameCounter + (timerElapsedTime * GrhData(MapData(X,
 
'Y).Graphic(2).grhindex).NumFrames / MapData(X, Y).Graphic(2).Speed)
'                        If MapData(X, Y).Graphic(2).FrameCounter > GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames Then
'                            MapData(X, Y).Graphic(2).FrameCounter = (MapData(X, Y).Graphic(2).FrameCounter Mod GrhData(MapData(X,
 
'Y).Graphic(2).grhindex).NumFrames) + 1
'
'                            If MapData(X, Y).Graphic(2).Loops <> -1 Then
'                                If MapData(X, Y).Graphic(2).Loops > 0 Then
'                                    MapData(X, Y).Graphic(2).Loops = MapData(X, Y).Graphic(2).Loops - 1
'                                Else
'                                    MapData(X, Y).Graphic(2).Started = 0
'                                End If
'                            End If
'                        End If
'                    End If
'
'                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(2).grhindex).Frames(MapData(X, Y).Graphic(2).FrameCounter)
'
'                offx = 0
'                offy = 0
'                If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
'                    offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
'                If GrhData(MapData(X, Y).Graphic(2).grhindex).TileHeight <> 1 Then _
'                    offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
'
'                Device_Box_Textured_Render CurrentGrhIndex, _
'                    (ScreenX - 1) * 32 + PixelOffsetX + offx, (ScreenY - 1) * 32 + PixelOffsetY + offy, _
'                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
'                    MapData(X, Y).light_value, _
'                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
'                    False _
'                    , 0
'#Else
''                    Call Draw_Grh(MapData(X, Y).Graphic(2), _
'                            (ScreenX - 1) * 32 + PixelOffsetX, _
'                            (ScreenY - 1) * 32 + PixelOffsetY, _
'                            1, 1, , X, Y)
'#End If
''                End If
'
''            ScreenX = ScreenX + 1
''        Next X'
'
'        'Reset ScreenX to original value and increment ScreenY
'        'ScreenX = ScreenX - X + screenminX
'        'ScreenY = ScreenY + 1
'    'Next Y
 
   
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(X, Y)
                '******************************************
 
                'Object Layer **********************************
         '       If .ObjGrh.GrhIndex <> 0 Then
         '       If Abs(nX - X) < 1 And (Abs(nY - Y)) < 1 And MapData(X, Y).Blocked = 0 Then
         '           Call Draw_Grh(.ObjGrh, _
         '                   PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value(), , , , , X, Y)
         '       End If
 
 
                If .ObjGrh.GrhIndex <> 0 Then
                Call Draw_Grh(.ObjGrh, _
                PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value(), , , , , X, Y)
                            End If
 
               
                'Char layer ************************************
If MapData(X, Y).CharIndex > 0 Then
                TempChar = CharList(MapData(X, Y).CharIndex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                Moved = 0
   
 With TempChar
            'If needed, move left and right
            If TempChar.scrollDirectionX <> 0 Then
 
                .MoveOffset.X = .MoveOffset.X + ScrollPixelsPerFrame * Sgn(.scrollDirectionX) * timerElapsedTime * engineBaseSpeed
                 
                If .Body.Walk(.Heading).SpeedCounter > 0 Then _
                .Body.Walk(.Heading).Started = 1
                .arma.WeaponWalk(TempChar.Heading).Started = 1
                .escudo.ShieldWalk(TempChar.Heading).Started = 1
 
                Moved = 1
               
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffset.X >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffset.X <= 0) Then
                    .MoveOffset.X = 0
                    .scrollDirectionX = 0
                End If
            End If
 
            'If needed, move up and down
            If TempChar.scrollDirectionY <> 0 Then
               
 
                TempChar.MoveOffset.Y = TempChar.MoveOffset.Y + ScrollPixelsPerFrame * Sgn(.scrollDirectionY) * timerElapsedTime * engineBaseSpeed
 
               
               
                If .Body.Walk(.Heading).SpeedCounter > 0 Then _
                .Body.Walk(.Heading).Started = 1
                TempChar.arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.escudo.ShieldWalk(TempChar.Heading).Started = 1
               
                Moved = 1
               
                If (Sgn(TempChar.scrollDirectionY) = 1 And TempChar.MoveOffset.Y >= 0) Or _
                        (Sgn(TempChar.scrollDirectionY) = -1 And TempChar.MoveOffset.Y <= 0) Then
                    .MoveOffset.Y = 0
                    .scrollDirectionY = 0
                End If
           
               
            End If
 
            If .Heading = 0 Then .Heading = 3
 
            If Moved = 0 Then
                .Body.Walk(.Heading).Started = 0
                .Body.Walk(.Heading).FrameCounter = 1
               
                .arma.WeaponWalk(.Heading).Started = 0
                .arma.WeaponWalk(.Heading).FrameCounter = 1
               
                .escudo.ShieldWalk(.Heading).Started = 0
                .escudo.ShieldWalk(.Heading).FrameCounter = 1
               
                .Moving = 0
            End If
           
            If TempChar.haciendoataque = 0 And .MoveOffset.X = 0 And .MoveOffset.Y = 0 Then
                .arma.WeaponWalk(.Heading).Started = 0
                '.arma.WeaponWalk(.Heading).FrameCounter = 1
                .escudo.ShieldWalk(.Heading).Started = 0
               
                End If
               
            If TempChar.haciendoataque = 1 Then
                .arma.WeaponWalk(.Heading).Started = 1
                .escudo.ShieldWalk(.Heading).Started = 1
                .haciendoataque = 0
            End If
           
    End With
    PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
    PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
           
                iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp + 32
                iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp + 32
               
                If Len(TempChar.Nombre) = 0 Then
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        'Cabeza
                        If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                        End If
                Else
                    If TempChar.Navegando = 1 Then
                        'Cuerpo (Barca / Galeon / Galera)
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                   
                    ElseIf Not CharList(MapData(X, Y).CharIndex).invisible And TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                       
                        'Cuerpo
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value)
                       
                        'Cabeza
                        If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                        End If
                       
                        'Casco
                        If TempChar.casco.Head(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.casco.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                        End If
                       
                        'Arma
                        If TempChar.arma.WeaponWalk(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.arma.WeaponWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        End If
                       
                        'Escudo
                        If TempChar.escudo.ShieldWalk(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.escudo.ShieldWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        End If
                   
                    End If
                       
                    If Nombres Then
                       
                        If Not (TempChar.invisible Or TempChar.Navegando = 1) Then
                       
                            Dim lCenter As Long
                            If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                Dim sClan As String
                                lCenter = (frmMain.textwidth(Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                sClan = Mid$(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                Call Grh_Text_Render(True, Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), iPPx - lCenter, iPPy + 30, D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)))
                                lCenter = (frmMain.textwidth(sClan) / 2) - 16
                                Call Grh_Text_Render(True, sClan, iPPx - lCenter, iPPy + 45, D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)))
                            Else
                                lCenter = (frmMain.textwidth(TempChar.Nombre) / 2) - 16
                                Call Grh_Text_Render(True, TempChar.Nombre, iPPx - lCenter, iPPy + 30, D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)))
                            End If
                     
                        End If
                       
                    End If
                End If
   
                If Dialogos.CantidadDialogos > 0 Then Call Dialogos.Update_Dialog_Pos((iPPx + TempChar.Body.HeadOffset.X), (iPPy + TempChar.Body.HeadOffset.Y), MapData(X, Y).CharIndex)
               
                CharList(MapData(X, Y).CharIndex) = TempChar
 
                If CharList(MapData(X, Y).CharIndex).FX <> 0 Then Call Draw_Grh(FxData(TempChar.FX).FX, iPPx + FxData(TempChar.FX).OffsetX, iPPy + FxData(TempChar.FX).OffsetY, 1, 1, Blanco(), True, , , MapData(X, Y).CharIndex)
               
            End If
                '*************************************************
               
               
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), _
                            ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value(), , , , , X, Y)
'#End If
                End If
                '************************************************
 
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5
 
If Not bTecho Then
        'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    Call Draw_Techo(MapData(X, Y).Graphic(4), _
                ScreenX * 32 + PixelOffsetX, _
                ScreenY * 32 + PixelOffsetY, _
                1, 1)
                End If
                '**********************************
               
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
        Else
        ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    Call Draw_Techo(MapData(X, Y).Graphic(4), _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 1, True, X, Y)
                End If
                '**********************************
               
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
           
        'If LuzMouse Then
        '    Light_Move (Light_Find(20)), UserPos.X + frmMain.MouseX \ 32 - frmMain.Renderer.ScaleWidth \ 64, UserPos.Y + frmMain.MouseY / 32 -frmMain.Renderer.ScaleHeight \ 64
        'End If
        Light_Render_All
   
   
 
   
End Sub
Public Function RenderSounds()

    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> plLluviain Then
                    Call Audio.StopWave
                    Call Audio.PlayWave("lluviain.wav", True)
                    frmMain.IsPlaying = plLluviain
                End If
                
                
            Else
                If frmMain.IsPlaying <> plLluviaout Then
                    Call Audio.StopWave
                    Call Audio.PlayWave("lluviaout.wav", True)
                    frmMain.IsPlaying = plLluviaout
                End If
                
                
            End If
        End If
    End If

End Function


Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And CharList(UserCharIndex).Pos.Y <= Y
        
End If

End Function



Function PixelPos(X As Integer) As Integer




PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer) As Boolean
Dim i As Byte
 
IniPath = App.Path & "\Init\"
 
 
UserPos.X = MinXBorder
UserPos.Y = MinYBorder
 
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
 
 
MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
 
 
ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
 
Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs
Call CargarAnimArmas
Call CargarAnimEscudos
 
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
 
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
 
 
 
 
'Parra: Aca inician las variables globales del Directx8
                   
 
    '****** INIT DirectX ******
    ' Create the root D3D objects
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate()
    Set D3DX = New D3DX8
   
   
If Not InitD3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING, setDisplayFormhWnd) Then
        If Not InitD3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING, setDisplayFormhWnd) Then
            MsgBox "El D3DDevice no pudo iniciar..."
            End
        End If
    End If
   
    D3DDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
 
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
   
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
   
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
   
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    D3DDevice.SetVertexShader FVF
   
   
   
    'partÌculas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
 
        '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
   
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
 
 
    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
   
 
   
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate()
    Set D3DX = New D3DX8
   
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
   
   
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.renderer.ScaleWidth
        .BackBufferHeight = frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hWnd
    End With
    DispMode.Format = D3DFMT_X8R8G8B8
    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8
        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.Format = DispModeBK.Format
        End If
    Else
        bump_map_supported = False
        DispMode.Format = DispModeBK.Format
    End If
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                           
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
   
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
   
    D3DDevice.SetVertexShader FVF
   
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
   
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
 
 
    engineBaseSpeed = 0.029
   
 
    'partÌculas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    Call Engine_Font_Initialize
    'Fuentes
    'Font_Create "Tahoma", 8, True, 0
    'Font_Create "Morpheus", 42, True, 0
 
    'Set Memory Status
    GlobalMemoryStatus pUdtMemStatus
    mFreeMemoryBytes = pUdtMemStatus.dwAvailPhys
 
 
 
'PARTICULAS & LUCECITAS M¡GICAS
 
    Call Base_Luz(125, 125, 125)
 
 
Light_Remove_All
 
    'Light_Create 45, 48, RGB(255, 255, 255), 5
    'Light_Create 50, 70, &HFFFFFFFF, 5
   
    Light_Render_All
 
InitTileEngine = True
End Function
Public Sub ShowNextFrame()
 
Dim OffsetCounterX As Single
Dim OffsetCounterY As Single
Dim ulttick As Long, esttick As Long
Dim timers(1 To 5) As Long
Dim loopc As Long

ScrollPixelsPerFrame = 4.5

Do While prgRun
 If frmConnect.Visible Then
    DibujarConectar
End If
    Call RefreshAllChars
   
    If EngineRun Then
        If frmMain.WindowState <> 1 Then
 
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrame * AddtoUserPos.X * timerElapsedTime * engineBaseSpeed
                If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
     
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
               OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrame * AddtoUserPos.Y * timerElapsedTime * engineBaseSpeed
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
   
   If frmMain.Inventario.Visible Then DibujarInventarioB
   
    D3DDevice.BeginScene
     D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorXRGB(0, 0, 0), 1#, 0
           
           
            If UserCiego Then
                D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            Else
                RenderScreen UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY
            End If
           
            'FPS
            Dim color As Long
            If FramesPerSec >= 30 Then
            color = D3DColorARGB(255, 255, 255, 255)
            ElseIf FramesPerSec >= 15 Then
            color = D3DColorARGB(255, 255, 255, 0)
            ElseIf FramesPerSec >= 1 Then
            color = D3DColorARGB(255, 255, 0, 0)
            End If
            Call Grh_Text_Render(True, "FPS: " & FramesPerSec, 10, 10, color)
            '/FPS
                      
            'If ModoTrabajo Then Text_Render font_list(1), "MODO TRABAJO", 40, 10, 100, 20, D3DColorARGB(255, 255, 0, 0), DT_TOP Or DT_LEFT,True
            If Cartel Then DibujarCartel
            If Dialogos.CantidadDialogos <> 0 Then Dialogos.MostrarTexto
            RenderSounds
        D3DDevice.Present ByVal 0, ByVal 0, frmMain.renderer.hWnd, ByVal 0
    D3DDevice.EndScene
   
            lFrameLimiter = GetTickCount
            FramesPerSecCounter = FramesPerSecCounter + 1
            timerElapsedTime = GetElapsedTime()
            timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
            particletimer = timerElapsedTime * 0.05
           
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
    End If
   
    If Not Pausa And frmMain.Visible And Not frmForo.Visible Then
        CheckKeys
    End If
 
    If GetTickCount - lFrameTimer > 1000 Then
        FramesPerSec = FramesPerSecCounter
        If FPSFLAG Then frmMain.Caption = "FenixAO DirectX8" & " v" & App.Major & "." & App.Minor & "." & App.Revision
        frmMain.fpstext.Caption = FramesPerSec
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
   
    'Limitar FPS
    While (GetTickCount - lFrameTimer) \ 14 < FramesPerSecCounter
    Sleep 5
    Wend
   
    ' ### I N T E R V A L O S ###
    esttick = GetTickCount
    For loopc = 1 To UBound(timers)
        timers(loopc) = timers(loopc) + (esttick - ulttick)
       
        If timers(1) >= tUs Then
            timers(1) = 0
            NoPuedeUsar = False
        End If
    Next loopc
    ulttick = GetTickCount
   
    DoEvents
Loop
 
End Sub
 
Sub CrearGrh(GrhIndex As Integer, index As Integer)
ReDim Preserve grh(1 To index) As grh
grh(index).FrameCounter = 1
grh(index).GrhIndex = GrhIndex
'Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
grh(index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1)
Call CrearGrh(534, 2)
End Sub

Function ControlVelocidad(ByVal LastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - LastTime > 20)
End Function

Sub Draw_Techo(grh As grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte)
 
Dim iGrhIndex As Integer
Dim QuitarAnimacion As Boolean
 
 
If Animate Then
    If grh.Started = 1 Then
       
        grh.FrameCounter = grh.FrameCounter + ((timerElapsedTime * 0.1) * GrhData(grh.GrhIndex).NumFrames / grh.SpeedCounter)
            If grh.FrameCounter > GrhData(grh.GrhIndex).NumFrames Then
               
                grh.FrameCounter = (grh.FrameCounter Mod GrhData(grh.GrhIndex).NumFrames) + 1
                   
                If KillAnim <> 0 Then
                If CharList(KillAnim).FX > 0 Then
                    If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                          CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes <= 0 Then CharList(KillAnim).FX = 0: Exit Sub
                        End If
                    End If
                End If
                End If
    End If
End If
 
If grh.GrhIndex = 0 Then Exit Sub
 
 
iGrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)
 
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If
 
If map_x Or map_y = 0 Then map_x = 1: map_y = 1
 
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        MapData(map_x, map_y).light_value(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        False, 0
 
End Sub
 
Sub Draw_Grh(grh As grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByRef color() As Long, Optional Alpha As Boolean, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte)
'***************************
'/////By Thusing/////
'***************************
 
Dim iGrhIndex As Integer
Dim QuitarAnimacion As Boolean
 
 
If Animate Then
    If grh.Started = 1 Then
       
        grh.FrameCounter = grh.FrameCounter + ((timerElapsedTime * 0.1) * GrhData(grh.GrhIndex).NumFrames / grh.SpeedCounter)
            If grh.FrameCounter > GrhData(grh.GrhIndex).NumFrames Then
               
                grh.FrameCounter = (grh.FrameCounter Mod GrhData(grh.GrhIndex).NumFrames) + 1
                   
                If KillAnim <> 0 Then
                If CharList(KillAnim).FX > 0 Then
                    If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                          CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes <= 0 Then CharList(KillAnim).FX = 0: Exit Sub
                        End If
                    End If
                End If
                End If
    End If
End If
 
If grh.GrhIndex = 0 Then Exit Sub
 
 
iGrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)
 
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If
 
If map_x Or map_y = 0 Then map_x = 1: map_y = 1
 
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        color(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        Alpha, 0
 
 
End Sub
Public Sub DrawDam_GrhIndex(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
   
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
   
    DeviceDam_Box_Textured_Render grh_index, _
        X, Y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY
End Sub
Sub DrawGrhtoHdc(hdc As Long, GrhIndex As Integer)
 
    Dim hDCsrc As Long
 
    If GrhIndex <= 0 Then Exit Sub
       
        'If it's animated switch GrhIndex to first frame
        If GrhData(GrhIndex).NumFrames <> 1 Then
            GrhIndex = GrhData(GrhIndex).Frames(1)
        End If
           
        hDCsrc = CreateCompatibleDC(hdc)
       
        Call SelectObject(hDCsrc, LoadPicture(App.Path & "\Graficos\" & GrhData(GrhIndex).FileNum & ".bmp"))
 
        'Draw
        BitBlt hdc, 0, 0, _
        GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelWidth, _
        hDCsrc, _
        GrhData(GrhIndex).sX, GrhData(GrhIndex).sY, _
        vbSrcCopy
 
        DeleteDC hDCsrc
End Sub
 
Private Function InitD3DDevice(ByVal MODE As CONST_D3DCREATEFLAGS, ByRef setDisplayFormhWnd As Long) As Boolean
 
    'When there is an error, destroy the D3D device and get ready to make a new one
    On Error GoTo ErrOut
   
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim DispMode As D3DDISPLAYMODE
   
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3DWindow.Windowed = True
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = DispMode.Format
   
  '###################################
  '## CHECK THE DEVICE CAPABILITIES ##
  '###################################
   
    Dim DevCaps As D3DCAPS8
   
    D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DevCaps
   
    If Err.Number = D3DERR_INVALIDDEVICE Then
        'We couldn't get data from the hardware device - probably doesn't exist...
        D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, DevCaps
        Err.Clear
    End If
   
    'Set the D3DDevices
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, setDisplayFormhWnd, MODE, D3DWindow)
   
    frmMain.Visible = False
    DoEvents
   
    'Everything was successful
    InitD3DDevice = True
   
Exit Function
 
ErrOut:
    'MsgBox "Error Number Returned: " & Err.Number & vbNewLine & "Description: " & Err.Description
   
    'Return a failure
    InitD3DDevice = False
End Function
Public Sub DeInitTileEngine()
 
    Dim i As Long
    Dim j As Long
   
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).texture = Nothing
            Next j
           
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
 
    Set Dx = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    Set FontList = Nothing
   
    Erase CharList
    Erase grh
    Erase GrhData
    Erase MapData
End Sub
 
Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function
 
Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
Dim R As Integer, G As Integer, b As Integer, a As Integer
    R = 127 * Vec.X + 128
    G = 127 * Vec.Y + 128
    b = 127 * Vec.Z + 128
    a = 255 * fHeight
    VectorToRGBA = D3DColorARGB(a, R, G, b)
End Function
 
Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal range As Byte = 1, Optional ByVal id As Long, Optional ByVal Red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan MartÌn Sotuyo Dodero
'**************************************************************
    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, range, id, Red, Green, Blue
    End If
End Function
 
Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
       
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
   
            Light_Move = True
           
        End If
    End If
End Function
 
Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan MartÌn Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim addY As Byte
    Dim addX As Byte
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If
 
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
   
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
       
 
 
        Select Case Heading
            Case NORTH
                addY = -1
       
            Case EAST
                addX = 1
       
            Case SOUTH
                addY = 1
           
            Case WEST
                addX = -1
        End Select
       
        nX = map_x + addX
        nY = map_y + addY
       
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
       
            'Move it
            Light_Erase light_index
 
            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
   
            Light_Move_By_Head = True
           
        End If
    End If
End Function
 
Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                        ByVal range As Long, Optional ByVal id As Long, Optional ByVal Red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
   
    'Make active
    light_list(light_index).active = True
   
        'Le damos color
    light_list(light_index).RGBcolor.R = Red
    light_list(light_index).RGBcolor.G = Green
    light_list(light_index).RGBcolor.b = Blue
   
    'Alpha (Si borras esto RE KB!!)
    light_list(light_index).RGBcolor.a = 255
   
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub
 
Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function
 
Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
           
    For loop_counter = 1 To light_count
       
        If light_list(loop_counter).active Then
            LightRender loop_counter
        End If
   
    Next loop_counter
End Sub
Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
   
    Dim CurrentColor As D3DCOLORVALUE
   
    pRadio = cRadio * 32
   
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
   
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
   
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.R), Round(CurrentColor.G), Round(CurrentColor.b))
        'If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function
Private Sub LightRender(ByVal light_index As Integer)
 
    If light_index = 0 Then Exit Sub
    If light_list(light_index).active = False Then Exit Sub
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim color As Long
    Dim Ya As Integer
    Dim Xa As Integer
   
    Dim TileLight As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
   
    Dim XCoord As Integer
    Dim YCoord As Integer
   
    AmbientColor.R = ColorLuz.R
    AmbientColor.G = ColorLuz.G
    AmbientColor.b = ColorLuz.b
 
   
    LightColor = light_list(light_index).RGBcolor
       
    min_x = light_list(light_index).map_x - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(1) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)
 
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(3) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(0) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
   
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(2) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
               
            End If
        Next Xa
    Next Ya
End Sub
 
 
Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function
 
Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function
 
Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
   
    For index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(index) Then
            Light_Destroy index
        End If
    Next index
   
    Light_Remove_All = True
End Function
 
Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As tLight
   
    Light_Erase light_index
   
    light_list(light_index) = temp
   
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub
 
Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan MartÌn Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
   
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
   
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0
    End If
   
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = 0
            MapData(X, min_y).light_value(2) = 0
        End If
    Next X
   
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = 0
            MapData(X, max_y).light_value(3) = 0
        End If
    Next X
   
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = 0
            MapData(min_x, Y).light_value(3) = 0
        End If
    Next Y
   
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = 0
            MapData(max_x, Y).light_value(1) = 0
        End If
    Next Y
   
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
            End If
        Next Y
    Next X
End Sub
 Private Function CreateColorVal(a As Integer, R As Integer, G As Integer, b As Integer) As D3DCOLORVALUE
    CreateColorVal.a = a
    CreateColorVal.R = R
    CreateColorVal.G = G
    CreateColorVal.b = b
End Function
Public Function ARGB(ByVal R As Long, ByVal G As Long, ByVal b As Long, ByVal a As Long) As Long
       
    Dim c As Long
       
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or b
    End If
   
    ARGB = c
 
End Function
Public Sub DeviceDam_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Juan MartÌn Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
   
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
   
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
   
    'If Not char_current_blind Then
        If (light_value(0) = 0) Then light_value(0) = Base_Light
        If (light_value(1) = 0) Then light_value(1) = Base_Light
        If (light_value(2) = 0) Then light_value(2) = Base_Light
        If (light_value(3) = 0) Then light_value(3) = Base_Light
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
       
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
               
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
   
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
   
    'Set Textures
    D3DDeviceDam.SetTexture 0, d3dTextures.texture
   
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDeviceDam.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDeviceDam.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
   
    'Draw the triangles that make up our square Textures
    D3DDeviceDam.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
   
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDeviceDam.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDeviceDam.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
 
End Sub
Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
 
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
   
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
   
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
   
    If (light_value(0) = 0) Then light_value(0) = Base_Light
    If (light_value(1) = 0) Then light_value(1) = Base_Light
    If (light_value(2) = 0) Then light_value(2) = Base_Light
    If (light_value(3) = 0) Then light_value(3) = Base_Light
       
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
               
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
   
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
   
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
   
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
   
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
   
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
 
End Sub
Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function
Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Integer, Optional ByRef Textures_Height As Integer, Optional ByVal angle As Single)
 
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
   
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
       
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
       
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
   
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
   
   
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
   
   
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
   
   
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
   
   
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
    End If
 
End Sub
Private Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Integer, ByRef textheight As Integer) As Direct3DTexture8
If FileName = 0 Then Debug.Print "ERROR! GRH = 0": Exit Function
 
    Dim i As Long
    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).FileName = FileName Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).texture
                Exit Function
            End If
        Next i
    End With
 
    'Not in memory, load it!
    Set GetTexture = CrearGrafico(FileName, textwidth, textheight)
End Function
Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Integer, ByRef textheight As Integer) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim index As Integer
    index = Archivo Mod HASH_TABLE_SIZE
    With TexList(index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        With .SurfaceEntry(.surfaceCount)
            'Nombre
            .FileName = Archivo
           
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
   
            Set .texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\GRAFICOS\" & LTrim(Str(Archivo)) & ".bmp", _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
               
            .texture.GetLevelDesc 0, surface_desc
            .texture_width = texture_info.width
            .texture_height = texture_info.height
            .size = surface_desc.size
            texwidth = .texture_width
            textheight = .texture_height
            Set CrearGrafico = .texture
            mFreeMemoryBytes = mFreeMemoryBytes - surface_desc.size
        End With
    End With
   
    Do While mFreeMemoryBytes < 0
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
ErrHandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".bmp"
End Function
 
Private Function RemoveLRU() As Boolean
   
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim surface_desc As D3DSURFACE_DESC
   
    LRUtime = GetTickCount
   
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).UltimoAcceso Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).UltimoAcceso
                End If
            Next j
        End With
    Next i
   
    'Retrieve the surface desc
    Call TexList(LRUi).SurfaceEntry(LRUj).texture.GetLevelDesc(0, surface_desc)
   
    'Remove it
    Set TexList(LRUi).SurfaceEntry(LRUj).texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).FileName = 0
   
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
       
        With TexList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
           
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
   
    'Update the used bytes
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
End Function
Public Function GetElapsedTime() As Single
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency
 
    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
   
    'Get current time
    Call QueryPerformanceCounter(start_time)
   
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
   
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function


Public Sub GrhRenderToHdc(ByVal grh_index As Long, desthDC As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)

    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hDCsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long

    If grh_index <= 0 Then Exit Sub

    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If


        file_path = DirGraficos & GrhData(grh_index).FileNum & ".bmp"
        
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
            
        hDCsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hDCsrc, LoadPicture(file_path))
        
        If transparent = False Then
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hDCsrc, src_x, src_y, vbSrcCopy
        Else
            MaskDC = CreateCompatibleDC(desthDC)
            
            PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
            
            Grh_Create_Mask hDCsrc, MaskDC, src_x, src_y, src_width, src_height
            
            'Render tranparently
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hDCsrc, src_x, src_y, vbSrcPaint
            
            Call DeleteObject(SelectObject(MaskDC, PrevObj2))
            
            DeleteDC MaskDC
        End If
        
        Call DeleteObject(SelectObject(hDCsrc, PrevObj))
        DeleteDC hDCsrc

    Exit Sub
End Sub

Private Sub Grh_Create_Mask(ByRef hDCsrc As Long, ByRef MaskDC As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim TransColor As Long
    Dim ColorKey As String
    
    ColorKey = "0"
    TransColor = &H0

    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height + src_y
        For X = src_x To src_width + src_x
            If GetPixel(hDCsrc, X, Y) = TransColor Then
                SetPixel MaskDC, X, Y, vbWhite
                SetPixel hDCsrc, X, Y, vbBlack
            Else
                SetPixel MaskDC, X, Y, vbBlack
            End If
        Next X
    Next Y
End Sub
 
Private Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/08/2006
'Modified by Juan MartÌn Sotuyo Dodero
'Modified by Augusto JosÈ Rando
'Added centering
'**************************************************************
Dim tile_width As Integer
Dim tile_height As Integer
Dim grh_index As Long
Dim timer_ticks_per_frame As Single
Dim base_tile_size As Integer
If grh.GrhIndex = 0 Then Exit Sub
 
'Animation
If grh.Started Then
grh.FrameCounter = grh.FrameCounter + (timer_ticks_per_frame * grh.SpeedCounter)
If grh.FrameCounter > GrhData(grh.GrhIndex).NumFrames Then
'If Grh.noloop Then
' Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
'Else
grh.FrameCounter = 1
'End If
End If
End If
 
'particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
'If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
' particle_group_list(particle_group_index).frame_counter = 0
' no_move = False
'Else
' no_move = True
'End If
 
'Figure out what frame to draw (always 1 if not animated)
If grh.FrameCounter = 0 Then grh.FrameCounter = 1
grh_index = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)
If grh_index <= 0 Then Exit Sub
If GrhData(grh_index).FileNum = 0 Then Exit Sub
 
'Modified by Augusto JosÈ Rando
'Simplier function - according to basic ORE engine
If h_centered Then
If GrhData(grh.GrhIndex).TileWidth <> 1 Then
screen_x = screen_x - Int(GrhData(grh.GrhIndex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
End If
End If
 
If v_centered Then
If GrhData(grh.GrhIndex).TileHeight <> 1 Then
screen_y = screen_y - Int(GrhData(grh.GrhIndex).TileHeight * base_tile_size) + base_tile_size
End If
End If
 
'Draw it to device
Device_Box_Textured_Render_Advance grh_index, _
screen_x, screen_y, _
GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
rgb_list(), _
GrhData(grh_index).sX, GrhData(grh_index).sY, _
alpha_blend
 
End Sub

Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, color As Long, outlinecolor As Long)
 
    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
   
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
   
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
   
    With box_rect
        .bottom = Y + height - 1
        .Left = X + 1
        .Right = X + width - 1
        .Top = Y + 1
    End With
   
    With Outline
        .bottom = Y + height
        .Left = X
        .Right = X + width
        .Top = Y
    End With
   
   
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
   
   
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
 
   
End Sub

Sub DibujarInventarioB()
 
    Dim re As RECT
    re.Left = 0
    re.Top = 0
    re.bottom = 160
    re.Right = 160
   
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
   
    Dim i As Byte, X As Integer, Y As Integer
    Dim T As grh
    
    Dim s(3) As Long
    s(0) = -1
    s(1) = -1
    s(2) = -1
    s(3) = -1
 
Call Device_Box_Textured_Render_Advance(19938, 0, 0, 160, 160, s, 0, 0)

    For Y = 1 To 5
        For X = 1 To 5
        i = i + 1
       
        If UserInventory(i).GrhIndex Then

        InitGrh T, UserInventory(i).GrhIndex

            With UserInventory(i)
                'Draw_FilledBox X * 32 - 32, Y * 32 - 32, 32, 32, D3DColorARGB(255, 0, 0, 0), D3DColorXRGB(255, 0, 0)

            If ItemElegido = i Then _
            Device_Box_Textured_Render_Advance 19939, X * 32 - 32, Y * 32 - 32, 32, 32, s, 0, 0
            Dibujar_grh_Simple T, X * 32 - 32, Y * 32 - 32, D3DColorXRGB(255, 255, 255)
            'Text_Render_ext UserInventory(i).Amount, (Y * 32) + 20 - 32, X * 32 - 32, 32, 32, D3DColorARGB(255, 255, 255, 255)
            Grh_Text_Render True, UserInventory(i).Amount, X * 32 - 32, (Y * 32) + 20 - 32, D3DColorARGB(255, 255, 255, 255)
            If UserInventory(i).Equipped Then _
            'Text_Render_ext "+", Y * 32 - 32, (X * 32) + 25 - 32, 32, 32, D3DColorARGB(255, 255, 0, 0)
            Grh_Text_Render True, "+", (X * 32) + 25 - 32, Y * 32 - 32, D3DColorARGB(255, 255, 0, 0)
            End If
        End With
    End If
    Next X, Y
 
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, frmMain.Inventario.hWnd, ByVal 0
 
End Sub
 
Sub Dibujar_grh_Simple(grh As grh, ByVal X As Integer, ByVal Y As Integer, Optional color As Long)
Dim c(3) As Long
 
If grh.GrhIndex = 0 Then Exit Sub
 
c(0) = color
c(1) = color
c(2) = color
c(3) = color
 
If grh.FrameCounter = 0 Then grh.FrameCounter = 1
 
With GrhData(grh.GrhIndex)
 
    Device_Box_Textured_Render_Advance grh.GrhIndex, X, Y, .pixelWidth, .pixelHeight, c(), .sX, .sY
 
End With
 
End Sub

'TEXTOS CARGADOS DESDE GRH
Public Sub Engine_Font_Initialize()
 
Dim a As Long
 
Fuentes(1).Caracteres(48) = 19730 ' 0
Fuentes(1).Caracteres(49) = 19731 ' 1
Fuentes(1).Caracteres(50) = 19732 ' 2
Fuentes(1).Caracteres(51) = 19733 ' 3
Fuentes(1).Caracteres(52) = 19734 ' 4
Fuentes(1).Caracteres(53) = 19735 ' 5
Fuentes(1).Caracteres(54) = 19736 ' 6
Fuentes(1).Caracteres(55) = 19737 ' 7
Fuentes(1).Caracteres(56) = 19738 ' 8
Fuentes(1).Caracteres(57) = 19739 ' 9
 
For a = 0 To 25
Fuentes(1).Caracteres(a + 97) = 19779 + a
Next a
 
For a = 0 To 25
Fuentes(1).Caracteres(a + 65) = 19747 + a
Next a
 
Fuentes(1).Caracteres(32) = 19714 '
Fuentes(1).Caracteres(33) = 19715 ' !
Fuentes(1).Caracteres(34) = 19716 ' "
Fuentes(1).Caracteres(35) = 19717 ' #
Fuentes(1).Caracteres(36) = 19718 ' $
Fuentes(1).Caracteres(37) = 19719 ' %
Fuentes(1).Caracteres(38) = 19720 ' &
Fuentes(1).Caracteres(39) = 19721 ' '
Fuentes(1).Caracteres(40) = 19722 ' (
Fuentes(1).Caracteres(41) = 19723 ' )
Fuentes(1).Caracteres(42) = 19724 ' *
Fuentes(1).Caracteres(43) = 19725 ' +
Fuentes(1).Caracteres(44) = 19726 ' ,
Fuentes(1).Caracteres(45) = 19727 ' -
Fuentes(1).Caracteres(46) = 19728 ' .
Fuentes(1).Caracteres(47) = 19729 ' /
Fuentes(1).Caracteres(58) = 19740 ' :
Fuentes(1).Caracteres(59) = 19741 ' ;
Fuentes(1).Caracteres(60) = 19742 ' <
Fuentes(1).Caracteres(61) = 19743 ' =
Fuentes(1).Caracteres(62) = 19744 ' >
Fuentes(1).Caracteres(63) = 19745 ' ?
Fuentes(1).Caracteres(64) = 19746 ' @
Fuentes(1).Caracteres(91) = 19773 ' [
Fuentes(1).Caracteres(92) = 19774 ' \
Fuentes(1).Caracteres(93) = 19775 ' ]
Fuentes(1).Caracteres(94) = 19776 ' ^
Fuentes(1).Caracteres(95) = 19777 '
Fuentes(1).Caracteres(96) = 19778 ' `
Fuentes(1).Caracteres(123) = 19805 ' {
Fuentes(1).Caracteres(124) = 19806 ' |
Fuentes(1).Caracteres(125) = 19807 ' }
Fuentes(1).Caracteres(126) = 19808 ' ~
Fuentes(1).Caracteres(127) = 19809 ' 
Fuentes(1).Caracteres(63) = 19810 ' ?
Fuentes(1).Caracteres(129) = 19811 ' Å
Fuentes(1).Caracteres(63) = 19812 ' ?
Fuentes(1).Caracteres(63) = 19813 ' ?
Fuentes(1).Caracteres(63) = 19814 ' ?
Fuentes(1).Caracteres(63) = 19815 ' ?
Fuentes(1).Caracteres(63) = 19816 ' ?
Fuentes(1).Caracteres(63) = 19817 ' ?
Fuentes(1).Caracteres(63) = 19818 ' ?
Fuentes(1).Caracteres(63) = 19819 ' ?
Fuentes(1).Caracteres(63) = 19820 ' ?
Fuentes(1).Caracteres(63) = 19821 ' ?
Fuentes(1).Caracteres(63) = 19822 ' ?
Fuentes(1).Caracteres(141) = 19823 ' ç
Fuentes(1).Caracteres(63) = 19824 ' ?
Fuentes(1).Caracteres(143) = 19825 ' è
Fuentes(1).Caracteres(144) = 19826 ' ê
Fuentes(1).Caracteres(63) = 19827 ' ?
Fuentes(1).Caracteres(63) = 19828 ' ?
Fuentes(1).Caracteres(63) = 19829 ' ?
Fuentes(1).Caracteres(63) = 19830 ' ?
Fuentes(1).Caracteres(63) = 19831 ' ?
Fuentes(1).Caracteres(63) = 19832 ' ?
Fuentes(1).Caracteres(63) = 19833 ' ?
Fuentes(1).Caracteres(63) = 19834 ' ?
Fuentes(1).Caracteres(63) = 19835 ' ?
Fuentes(1).Caracteres(63) = 19836 ' ?
Fuentes(1).Caracteres(63) = 19837 ' ?
Fuentes(1).Caracteres(63) = 19838 ' ?
Fuentes(1).Caracteres(157) = 19839 ' ù
Fuentes(1).Caracteres(63) = 19840 ' ?
Fuentes(1).Caracteres(63) = 19841 ' ?
Fuentes(1).Caracteres(160) = 19842 '
Fuentes(1).Caracteres(161) = 19843 ' °
Fuentes(1).Caracteres(162) = 19844 ' ¢
Fuentes(1).Caracteres(163) = 19845 ' £
Fuentes(1).Caracteres(164) = 19846 ' §
Fuentes(1).Caracteres(165) = 19847 ' •
Fuentes(1).Caracteres(166) = 19848 ' ¶
Fuentes(1).Caracteres(167) = 19849 ' ß
Fuentes(1).Caracteres(168) = 19850 ' ®
Fuentes(1).Caracteres(169) = 19851 ' ©
Fuentes(1).Caracteres(170) = 19852 ' ™
Fuentes(1).Caracteres(171) = 19853 ' ´
Fuentes(1).Caracteres(172) = 19854 ' ¨
Fuentes(1).Caracteres(173) = 19855 '
Fuentes(1).Caracteres(174) = 19856 ' Æ
Fuentes(1).Caracteres(175) = 19857 ' Ø
Fuentes(1).Caracteres(176) = 19858 ' ∞
Fuentes(1).Caracteres(177) = 19859 ' ±
Fuentes(1).Caracteres(178) = 19860 ' ≤
Fuentes(1).Caracteres(179) = 19861 ' ≥
Fuentes(1).Caracteres(180) = 19862 ' ¥
Fuentes(1).Caracteres(181) = 19863 ' µ
Fuentes(1).Caracteres(182) = 19864 ' ∂
Fuentes(1).Caracteres(183) = 19865 ' ∑
Fuentes(1).Caracteres(184) = 19866 ' ∏
Fuentes(1).Caracteres(185) = 19867 ' π
Fuentes(1).Caracteres(186) = 19868 ' ∫
Fuentes(1).Caracteres(187) = 19869 ' ª
Fuentes(1).Caracteres(188) = 19870 ' º
Fuentes(1).Caracteres(189) = 19871 ' Ω
Fuentes(1).Caracteres(190) = 19872 ' æ
Fuentes(1).Caracteres(191) = 19873 ' ø
Fuentes(1).Caracteres(192) = 19874 ' ¿
Fuentes(1).Caracteres(193) = 19875 ' ¡
Fuentes(1).Caracteres(194) = 19876 ' ¬
Fuentes(1).Caracteres(195) = 19877 ' √
Fuentes(1).Caracteres(196) = 19878 ' ƒ
Fuentes(1).Caracteres(197) = 19879 ' ≈
Fuentes(1).Caracteres(198) = 19880 ' ∆
Fuentes(1).Caracteres(199) = 19881 ' «
Fuentes(1).Caracteres(200) = 19882 ' »
Fuentes(1).Caracteres(201) = 19883 ' …
Fuentes(1).Caracteres(202) = 19884 '  
Fuentes(1).Caracteres(203) = 19885 ' À
Fuentes(1).Caracteres(204) = 19886 ' Ã
Fuentes(1).Caracteres(205) = 19887 ' Õ
Fuentes(1).Caracteres(206) = 19888 ' Œ
Fuentes(1).Caracteres(207) = 19889 ' œ
Fuentes(1).Caracteres(208) = 19890 ' –
Fuentes(1).Caracteres(209) = 19891 ' —
Fuentes(1).Caracteres(210) = 19892 ' “
Fuentes(1).Caracteres(211) = 19893 ' ”
Fuentes(1).Caracteres(212) = 19894 ' ‘
Fuentes(1).Caracteres(213) = 19895 ' ’
Fuentes(1).Caracteres(214) = 19896 ' ÷
Fuentes(1).Caracteres(215) = 19897 ' ◊
Fuentes(1).Caracteres(216) = 19898 ' ÿ
Fuentes(1).Caracteres(217) = 19899 ' Ÿ
Fuentes(1).Caracteres(218) = 19900 ' ⁄
Fuentes(1).Caracteres(219) = 19901 ' €
Fuentes(1).Caracteres(220) = 19902 ' ‹
Fuentes(1).Caracteres(221) = 19903 ' ›
Fuentes(1).Caracteres(222) = 19904 ' ﬁ
Fuentes(1).Caracteres(223) = 19905 ' ﬂ
Fuentes(1).Caracteres(224) = 19906 ' ‡
Fuentes(1).Caracteres(225) = 19907 ' ·
Fuentes(1).Caracteres(226) = 19908 ' ‚
Fuentes(1).Caracteres(227) = 19909 ' „
Fuentes(1).Caracteres(228) = 19910 ' ‰
Fuentes(1).Caracteres(229) = 19911 ' Â
Fuentes(1).Caracteres(230) = 19912 ' Ê
Fuentes(1).Caracteres(231) = 19913 ' Á
Fuentes(1).Caracteres(232) = 19914 ' Ë
Fuentes(1).Caracteres(233) = 19915 ' È
Fuentes(1).Caracteres(234) = 19916 ' Í
Fuentes(1).Caracteres(235) = 19917 ' Î
Fuentes(1).Caracteres(236) = 19918 ' Ï
Fuentes(1).Caracteres(237) = 19919 ' Ì
Fuentes(1).Caracteres(238) = 19920 ' Ó
Fuentes(1).Caracteres(239) = 19921 ' Ô
Fuentes(1).Caracteres(240) = 19922 ' 
Fuentes(1).Caracteres(241) = 19923 ' Ò
Fuentes(1).Caracteres(242) = 19924 ' Ú
Fuentes(1).Caracteres(243) = 19925 ' Û
Fuentes(1).Caracteres(244) = 19926 ' Ù
Fuentes(1).Caracteres(245) = 19927 ' ı
Fuentes(1).Caracteres(246) = 19928 ' ˆ
Fuentes(1).Caracteres(247) = 19929 ' ˜
Fuentes(1).Caracteres(248) = 19930 ' ¯
Fuentes(1).Caracteres(249) = 19931 ' ˘
Fuentes(1).Caracteres(250) = 19932 ' ˙
Fuentes(1).Caracteres(251) = 19933 ' ˚
Fuentes(1).Caracteres(252) = 19934 ' ¸
Fuentes(1).Caracteres(253) = 19935 ' ˝
Fuentes(1).Caracteres(254) = 19936 ' ˛
Fuentes(1).Caracteres(255) = 19937 ' ˇ
End Sub
 
Public Sub Grh_Text_Render(ByVal Borde As Boolean, ByVal Texto As String, ByVal X As Integer, ByVal Y As Integer, ByRef color As Long, Optional ByVal Alpha As Boolean = False, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)
 
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer
Dim graf As grh
Dim text_color(3) As Long
text_color(0) = color
text_color(1) = color
text_color(2) = color
 
Dim Negro(3) As Long
Negro(0) = D3DColorXRGB(0, 0, 0)
Negro(1) = D3DColorXRGB(0, 0, 0)
Negro(2) = D3DColorXRGB(0, 0, 0)
Negro(3) = D3DColorXRGB(0, 0, 0)
 
text_color(3) = color
 
If (Len(Texto) = 0) Then Exit Sub
 
d = 0
If multi_line = False Then
For a = 1 To Len(Texto)
b = Asc(Mid$(Texto, a, 1))
graf.GrhIndex = Fuentes(font_index).Caracteres(b)
If b <> 32 Then
If graf.GrhIndex <> 0 Then
'mega sombra O-matica
graf.GrhIndex = Fuentes(font_index).Caracteres(b)
If Borde Then
Grh_Render graf, (X + d) - 1, Y, Negro(), False, False, False
Grh_Render graf, (X + d), Y - 1, Negro(), False, False, False
End If
 
Grh_Render graf, (X + d), Y, text_color, False, False, Alpha
'Draw_Grh graf, (x + d), y, 0, 0, text_color, Alpha, False, False
d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 2
End If
Else
d = d + 4
End If
Next a
Else
e = 0
f = 0
For a = 1 To Len(Texto)
b = Asc(Mid$(Texto, a, 1))
graf.GrhIndex = Fuentes(font_index).Caracteres(b)
If b = 32 Or b = 13 Then
If e >= 20 Then 'reemplazar por lo que os plazca
f = f + 1
e = 0
d = 0
Else
If b = 32 Then d = d + 4
End If
Else
If graf.GrhIndex > 12 Then
'mega sombra O-matica
graf.GrhIndex = Fuentes(font_index).Caracteres(b)
If Borde Then
Grh_Render graf, (X + d) - 1, Y + f * 13, Negro(), False, False, False
Grh_Render graf, (X + d), Y + f * 13 - 1, Negro(), False, False, False
End If
 
Grh_Render graf, (X + d), Y + f * 13, text_color, False, False, Alpha
'Draw_Grh graf, (x + d), y + f * 13, 0, 0, text_color, Alpha, False, False
d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 2
End If
End If
e = e + 1
Next a
End If
 
End Sub

Public Sub Engine_Init2()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
'On Error GoTo ErrHandler:
 
    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindowDam As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
   
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate()
    Set D3DX = New D3DX8
   
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
   
    With D3DWindowDam
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmConnect.ScaleWidth
        .BackBufferHeight = frmConnect.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmConnect.hWnd
    End With
   
    DispMode.Format = D3DFMT_X8R8G8B8
    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8
        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.Format = DispModeBK.Format
        End If
    Else
        bump_map_supported = False
        DispMode.Format = DispModeBK.Format
    End If
 
    Set D3DDeviceDam = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmConnect.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindowDam)
 
                                                           
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
   
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
   
    D3DDeviceDam.SetVertexShader FVF
   
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDeviceDam.SetRenderState D3DRS_LIGHTING, False
   
    D3DDeviceDam.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDeviceDam.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDeviceDam.SetRenderState D3DRS_ALPHABLENDENABLE, True
   
 
    engineBaseSpeed = 0.015
   
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

   
    UserPos.X = 50
    UserPos.Y = 50
   
    MinXBorder = XMinMapSize + (frmConnect.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmConnect.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmConnect.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmConnect.ScaleHeight / 64)
 
   
    'partÌculas
    D3DDeviceDam.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDeviceDam.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDeviceDam.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDeviceDam.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    'motion blur
    'Set D3DbackBuffer = D3DDeviceDam.GetRenderTarget
    'Set zTarget = D3DDeviceDam.GetDepthStencilSurface
    'Set stencil = D3DDeviceDam.CreateDepthStencilSurface(800, 600, D3DFMT_D16, D3DMULTISAMPLE_NONE)
    'Set Tex = D3DX.CreateTexture(D3DDeviceDam, dimeTex, dimeTex, 1, D3DUSAGE_RENDERTARGET, D3DFMT_X8R8G8B8, D3DPOOL_DEFAULT)
    'Set superTex = Tex.GetSurfaceLevel(0)
    'blur_factor = 10
    'bump mapping
   
    'Font_Create "Tahoma", 8, True, 0
    'Font_Create "Verdana", 8, False, 0
   
bRunning = True
Exit Sub
ErrHandler:
Debug.Print "Error Number Returned: " & Err.Number
bRunning = False
End Sub

Public Sub DibujarConectar()
 
D3DDeviceDam.BeginScene
D3DDeviceDam.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
 
Call DrawDam_GrhIndex(5000, 50, 50)
 
D3DDeviceDam.EndScene
D3DDeviceDam.Present ByVal 0, ByVal 0, frmConnect.hWnd, ByVal 0
End Sub
