Attribute VB_Name = "ModIni"
'================================
'>>>>> WWW.FADICTOS.COM.AR <<<<<<
'================================
'Matute - matius_xd07@hotmail.com

Option Explicit

Public Declare Function INICarga Lib "LeeInis.dll" (ByVal Arch As String) As Long
Public Declare Function INIDescarga Lib "LeeInis.dll" (ByVal A As Long) As Long
Public Declare Function INIDarError Lib "LeeInis.dll" () As Long

Public Declare Function INIDarNumSecciones Lib "LeeInis.dll" (ByVal A As Long) As Long
Public Declare Function INIDarNombreSeccion Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long, ByVal Buff As String, ByVal Tam As Long) As Long
Public Declare Function INIBuscarSeccion Lib "LeeInis.dll" (ByVal A As Long, ByVal Buff As String) As Long

Public Declare Function INIDarClave Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long, ByVal Clave As String, ByVal Buff As String, ByVal Tam As Long) As Long
Public Declare Function INIDarClaveInt Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long, ByVal Clave As String) As Long
Public Declare Function INIDarNumClaves Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long) As Long
Public Declare Function INIDarNombreClave Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long, ByVal Clave As Long, ByVal Buff As String, ByVal Tam As Long) As Long

Public Declare Function INIConf Lib "LeeInis.dll" (ByVal A As Long, ByVal DefectoInt As Long, ByVal DefectoStr As String, ByVal CaseSensitive As Long) As Long


Public Function INIDarClaveStr(A As Long, Seccion As Long, Clave As String) As String
Dim Tmp As String
Dim P As Long, r As Long

Tmp = Space$(3000)
r = INIDarClave(A, Seccion, Clave, Tmp, 3000)
P = InStr(1, Tmp, Chr$(0))
If P Then
    Tmp = Left$(Tmp, P - 1)
    
    INIDarClaveStr = Tmp
End If

End Function

Public Function INIDarNombreSeccionStr(A As Long, Seccion As Long) As String
Dim Tmp As String
Dim P As Long, r As Long

Tmp = Space$(3000)
r = INIDarNombreSeccion(A, Seccion, Tmp, 3000)
P = InStr(1, Tmp, Chr$(0))
If P Then
    Tmp = Left$(Tmp, P - 1)
    INIDarNombreSeccionStr = Tmp
End If

End Function

Public Function INIDarNombreClaveStr(A As Long, Seccion As Long, Clave As Long) As String
Dim Tmp As String
Dim P As Long, r As Long

Tmp = Space$(3000)
r = INIDarNombreClave(A, Seccion, Clave, Tmp, 3000)
P = InStr(1, Tmp, Chr$(0))
If P Then
    Tmp = Left$(Tmp, P - 1)
    INIDarNombreClaveStr = Tmp
End If

End Function

