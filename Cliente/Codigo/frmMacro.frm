VERSION 5.00
Begin VB.Form frmMacro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Macros"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtComando 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
   Begin VB.OptionButton optEquipar 
      Caption         =   "Equipar item elegido"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
   Begin VB.OptionButton optUsar 
      Caption         =   "Usar item elegido"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.OptionButton optHechi 
      Caption         =   "Usar hechizo elegido"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.OptionButton optComando 
      Caption         =   "Comando"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label numF 
      Height          =   135
      Left            =   1440
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblF 
      Alignment       =   2  'Center
      Caption         =   "Macro F1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
txtComando.Text = ""
Me.Hide
End Sub

Private Sub cmdOk_Click()

If optComando.value = True And txtComando.Text = "" Then
    msg "Ingrese un comando"
    Exit Sub
End If

Macros(Val(numF - 1)).comando = ""
Macros(Val(numF - 1)).Equipar = 0
Macros(Val(numF - 1)).Hechizo = 0
Macros(Val(numF - 1)).Usar = 0

If optComando.value Then
    Macros(Val(numF - 1)).comando = Trim(txtComando.Text)
ElseIf optHechi.value Then
    Macros(Val(numF - 1)).Hechizo = frmMain.hlst.listIndex + 1
ElseIf optUsar.value Then
    Macros(Val(numF - 1)).Usar = Inventario.SelectedItem
ElseIf optEquipar.value Then
    Macros(Val(numF - 1)).Equipar = Inventario.SelectedItem
Else
    msg "Seleccione alguna opción"
    Exit Sub
End If

Me.Hide
GuardarMacros
CargarMacros (False)
End Sub

