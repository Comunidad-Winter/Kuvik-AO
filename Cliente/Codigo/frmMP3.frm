VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMP3 
   BorderStyle     =   0  'None
   Caption         =   "Reproductor MP3"
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMP3.frx":0000
   ScaleHeight     =   885
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   4080
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tu Musica en MP3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   840
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   2640
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   1560
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3240
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Función Api GetShortPathName para obtener _
los paths de los archivos en formato corto
Private Declare Function GetShortPathName _
    Lib "kernel32" _
    Alias "GetShortPathNameA" ( _
        ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, _
        ByVal lBuffer As Long) As Long

'Función Api mciExecute para reproducir los archivos de música
Private Declare Function mciExecute _
    Lib "winmm.dll" ( _
        ByVal lpstrCommand As String) As Long
Dim ret As Long, Path As String
'Le pasamos el comando Close a MciExecute para cerrar el dispositivo
Private Sub Form_Unload(Cancel As Integer)
    mciExecute "Close All"
End Sub

'Sub que obtiene el path corto del archivo a reproducir
Private Sub PathCorto(Archivo As String)
Dim temp As String * 250 'Buffer
    Path = String(255, 0)
    'Obtenemos el Path corto
    ret = GetShortPathName(Archivo, temp, 164)
    'Sacamos los nulos al path
    Path = Replace(temp, Chr(0), "")
End Sub

'Procedimiento que ejecuta el comando con el Api mciExecute
'************************************************************
Private Sub ejecutar(comando As String)
    If Path = "" Then MsgBox "Error", vbCritical: Exit Sub
    'Llamamos a mciExecute pasandole un string que tiene el comando y la ruta

    mciExecute comando & Path

End Sub

Private Sub Image1_Click()
    ejecutar ("Pause ")
End Sub

Private Sub Image2_Click()
frmMP3.Hide
End Sub

Private Sub Image3_Click()
    With CommonDialog1
        .Filter = "Archivos Mp3|*.mp3|Archivos Wav|*.wav|Archivos MIDI|*.mid"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            'Le pasamos a la sub que obtiene con _
            el Api GetShortPathName el nombre corto del archivo
            PathCorto .FileName
            Label1 = .FileName
            'cerramos todo
            mciExecute "Close All"
            'Para Habilitar y deshabilitar botones
        End If
    End With
End Sub

Private Sub Image4_Click()
    ejecutar ("Stop ")
End Sub

Private Sub Image5_Click()
    ejecutar ("Play ")
End Sub

Private Sub Image6_Click()
Unload Me
End Sub
