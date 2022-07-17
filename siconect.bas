Attribute VB_Name = "siconect"
Option Explicit

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255

'Función que devuelve TRUE si se está conectado a Internet
Public Function Conectado() As Boolean
    Dim ret As Long
    ret = InternetGetConnectedStateEx(ret, sConnType, 254, 0)
    If ret = 1 Then
        Conectado = True
    Else
        Conectado = False
    End If
End Function

