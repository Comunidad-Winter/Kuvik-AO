VERSION 5.00
Begin VB.Form FrmIntro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   240
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   960
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":0614
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":091E
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   3135
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Ret As Long
Private cont2 As Integer

Private Sub Form_Load()

   cont2 = 255
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    Timer1.Interval = 1
    Timer2.Interval = 1
    Timer2.Enabled = False
    Timer1.Enabled = True
    
Me.Picture = LoadPicture(App.Path & "\Graficos\MenuRapido.jpg")

Dim corriendo As Integer
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim pepe As String

Dim exename As String
snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)
i = 0
While theloop <> 0
    exename = proc.szExeFile
    Text1.Text = proc.szExeFile
    If Text1.Text = "FenixAONoDinamico.exe" Or Text1.Text = "FenixAO.exe" Then
        corriendo = corriendo + 1
        Text1.Text = ""
    End If
    i = i + 1
    theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Timer2.Enabled = True
End Sub

Private Sub Image2_Click()

Call Main

End Sub

Private Sub Image3_Click()
ShellExecute Me.hWnd, "open", App.Path & "/aosetup.exe", "", "", 1
End Sub

Private Sub Image4_Click()
ShellExecute Me.hWnd, "open", "http://www.fenixao.com.ar/public_html/Html/manual/", "", "", 1

End Sub

Private Sub Image5_Click()
ShellExecute Me.hWnd, "open", "http://www.fenixao.com.ar", "", "", 1

End Sub

Private Sub Image6_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
    Static cont As Integer
    cont = cont + 5
    If cont > 255 Then
        cont = 0
        Timer1.Enabled = False
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, cont, LWA_ALPHA
    End If
End Sub
 
Private Sub Timer2_Timer()
    cont2 = cont2 - 5
    If cont2 < 0 Then
        Timer2.Enabled = False
        End
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, cont2, LWA_ALPHA
    End If
End Sub
