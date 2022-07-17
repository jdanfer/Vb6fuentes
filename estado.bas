Attribute VB_Name = "estado"
Option Explicit

'Array de frmStatus, uno por cada mail
Public ColFrmStatus() As frmstatus

Public Enum CONEXION
    CERRADO = 0
    CONECTED = 1
    AUTH1 = 2
    AUTH2 = 3
    AUTH3 = 4
    MailFrom = 5
    RCPTTO = 6
    DATAC = 7
    MESSAGGE = 8
    QUIT = 9
    MANDADO_OK = 10
    cERROR = 11
End Enum
Public SendStatus(0 To 255) As CONEXION

Public TextStatus(0 To 11) As String

'Variables usadas para el diálogo con el servidor SMTP
Public Respuesta As String
Public Code As Integer
Public DServer As String
Public DHelo As String
Public DMailFrom As String
Public DRcptTo As String
Public DSubject As String
Public DMensaje As String
Public DFrom As String

Public exCaption As String

'conectar al servidor
Sub Conectar(Sock As Winsock)
     Sock.Close
     Sock.Connect DServer, 25

End Sub

'cerrar conexion
Sub DesConectar(Sock As Winsock)
    Sock.Close

End Sub

Sub DesConectarTodos()
frm_envyrec.sck.Close

End Sub

'agregar status
Sub AddStatus(frm As frmstatus, Texto As String)
    frm.txtStatus = frm.txtStatus & vbCrLf & Texto
    frm.txtStatus.SelStart = Len(frm.txtStatus.Text)
    frm.txtStatus.Refresh
End Sub

'generador de codigos alfanumericos
Public Function GenerateCode(NumChar As Integer)
    Randomize Timer
    Dim Code As String
    Dim Chars As Integer
    Dim Alfa As Integer
    Code = ""
    For Chars = 1 To NumChar
        Alfa = Int(Rnd * 2 + 1)
        If Alfa = 2 Then
            Code = Chr(Int((Rnd * 25 + 1) + 97)) & Code
        Else
            Code = Int((Rnd * 9 + 1)) & Code
        End If
    Next
    GenerateCode = Code
End Function

Public Function Enviar(Sock As Winsock, From As String, MailFrom As String, MailTo As String, subject As String)
'Función que comienza el envío de un mail
'aunque el envío tiene lugar realmente en el evento DataArrival del control Sck en frmMain
        
    DHelo = GenerateCode(8)
    DMailFrom = MailFrom
    DFrom = From
    DSubject = subject
    DMensaje = ""
    DRcptTo = MailTo
    
    Conectar Sock
End Function

