VERSION 5.00
Begin VB.Form frmBorrar 
   BorderStyle     =   0  'None
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2325
      Width           =   2895
   End
   Begin VB.TextBox Nombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   1080
      TabIndex        =   0
      Top             =   1260
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmBorrar.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image cmdBorrar 
      Height          =   495
      Left            =   1080
      MouseIcon       =   "frmBorrar.frx":030A
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   3255
   End
End
Attribute VB_Name = "frmBorrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez


Private Sub cmdBorrar_Click()

EstadoLogin = BorrarPJ
frmMain.Socket1.HostName = IPdelServidor
frmMain.Socket1.RemotePort = PuertoDelServidor
Me.MousePointer = 11
frmMain.Socket1.Connect

End Sub
Private Sub command1_Click()

Me.Visible = False

End Sub
Private Sub Command2_Click()

Call Audio.PlayWave("BGM_2E.WAV", True)

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "BorrarPersonaje.gif")

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Image1_Click()

Me.Hide

End Sub
