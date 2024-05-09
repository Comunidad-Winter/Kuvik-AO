VERSION 5.00
Begin VB.Form frmCuent 
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Cuentas"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   Picture         =   "frmCuenta.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   3240
      MousePointer    =   99  'Custom
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   32
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   5280
      MousePointer    =   99  'Custom
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   31
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   7320
      MousePointer    =   99  'Custom
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   1200
      MousePointer    =   99  'Custom
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   7320
      MousePointer    =   99  'Custom
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   5280
      MousePointer    =   99  'Custom
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   3240
      MousePointer    =   99  'Custom
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   1200
      MousePointer    =   99  'Custom
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   9000
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   9000
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Image imgWeb 
      Height          =   1455
      Left            =   8400
      Top             =   0
      Width           =   3615
   End
   Begin VB.Image imgConectar 
      Height          =   615
      Left            =   9000
      MouseIcon       =   "frmCuenta.frx":45604
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Image imgCrear 
      Height          =   735
      Left            =   9000
      MouseIcon       =   "frmCuenta.frx":4590E
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   29
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   28
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   27
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6960
      TabIndex        =   26
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4920
      TabIndex        =   25
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   24
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   23
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   22
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   21
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PJClick"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   19
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   18
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   17
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   16
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   15
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   14
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   13
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   11
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   9
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   8
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim i As Integer
Label3.Caption = UserName
End Sub
Private Sub Image1_Click()
frmMain.Socket1.Disconnect
frmConnect.MousePointer = 1
Unload Me
End Sub
Private Sub Label7_Click()
frmBorrar.Show , frmCuent
End Sub

Private Sub imgConectar_Click()
If PJClickeado = "Nada" Then
MsgBox "Seleccione un pj"
End If
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OLOGIO" & UserName)
Unload Me
End Sub

Private Sub imgCrear_Click()

Call Audio.PlayWave(SND_CLICK)

If Nombre(7).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

EstadoLogin = Dados
frmCrearPersonaje.Show vbModal
Me.MousePointer = 11
End Sub

Private Sub nombre_dblClick(index As Integer)
If PJClickeado = "Nada" Then Exit Sub

Call Audio.PlayWave(SND_CLICK)
Call SendData("OLOGIO" & PJClickeado)
Unload Me

End Sub
Private Sub nombre_Click(index As Integer)

PJClickeado = frmCuent.Nombre(index).Caption

End Sub
Private Sub PJ_Click(index As Integer)
PJClickeado = frmCuent.Nombre(index).Caption
End Sub
Private Sub PJ_dblClick(index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OLOGIO" & UserName)
Unload Me

End Sub

