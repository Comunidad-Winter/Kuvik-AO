VERSION 5.00
Begin VB.Form frmCrearAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Cuenta"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   5970
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearCuenta.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Mail 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox RePass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Nombre 
      Height          =   285
      Left            =   2880
      MaxLength       =   25
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3240
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
If Pass <> RePass Then
    MsgBox "Lass passwords que tipeo no coinciden", , "MD5 Changed Info Tip"
    Exit Sub
End If

If Not CheckMailString(Mail) Then
    MsgBox "Direccion de mail invalida."
    Exit Sub
End If

If Nombre = "" Or Pass = "" Or RePass = "" Or Mail = "" Then
    MsgBox "Completa todo!"
    Exit Sub
End If

Pass = MD5String(Pass.Text)

Call SendData("NACCNT" & Nombre & "," & Pass & "," & Mail)

Unload Me
End Sub

Private Sub Mail_GotFocus()
MsgBox "Deberas ingresar tu MAIL correcto, de lo contrario, no tendras respuesta de los Game Master en tus Soportes"
End Sub
