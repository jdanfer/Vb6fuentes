VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_envyrec 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   2760
   End
   Begin VB.TextBox txtUU 
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   6255
   End
   Begin MSWinsockLib.Winsock sck 
      Left            =   4920
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox objeto 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox correopa 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox correode 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox nomde 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox servidor 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frm_envyrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variable que contendrá el nombre de usuario (Si se usa autenticación)
Public aUSUARIO As String
'Variable que contendrá la contraseña (Si se usa autenticación)
Public aCONTRASEÑA As String
Private Function Plain_base64(User As String, Password As String) As String
'Genera la cadena que hay que mandar para un AUTH PLAIN
'(en este caso no lo uso porque uso AUTH LOGIN)
'(ver http://www.technoids.org/saslmech.html)
Dim s As String, i As Long
Dim sUser As String, sPassw As String
Dim nArray() As Byte

sUser = User
sPassw = Password

ReDim nArray(0 To Len(sUser) + Len(sPassw) + 1)

nArray(0) = 0
For i = 1 To Len(sUser)
    nArray(i) = Asc(Mid(sUser, i, 1))
Next i
nArray(i) = 0
For i = 1 To Len(sPassw)
    nArray(i + Len(sUser) + 1) = Asc(Mid(sPassw, i, 1))
Next i

Base64Array_Encode nArray

s = ""
For i = 0 To UBound(nArray)
    s = s & Chr(nArray(i))
Next i
Plain_base64 = s
End Function

Private Function Str_to_base64(s As String) As String
'Convierte una cadena en formato base64 para el AUTH LOGIN
'(ver http://www.technoids.org/saslmech.html)
Dim nArray() As Byte, i As Integer, sTemp As String
ReDim nArray(0 To Len(s) + 1)

For i = 0 To Len(s) - 1
    nArray(i) = Asc(Mid(s, i + 1, 1))
Next i

Base64Array_Encode nArray

sTemp = ""
For i = 0 To UBound(nArray)
    sTemp = sTemp & Chr(nArray(i))
Next i
Str_to_base64 = sTemp
End Function


Private Sub Command1_Click()
Dim Archadjunto, Rutaynom As String
'''If CD.FileName = "" Then Exit Sub
Archadjunto = "controles.mdb"
Rutaynom = "c:\sapp\controles.mdb"
UUfiles(indexUUfiles) = UUEncodeFile(Rutaynom)

'Codifico el archivo en el formato válido para ser adjuntado a un mail
indexUUfiles = 1
txtUU.Text = txtUU.Text & Archadjunto & " (" & Fix(FileLen(Rutaynom) / 1024) + 1 & " Kb)   "

''' enviar

DHelo = GenerateCode(8)
DMailFrom = correode.Text
DFrom = nomde.Text
DSubject = objeto.Text
DMensaje = "Hola, es una prueba"
DRcptTo = correopa.Text
txtUU.Text = ""
DServer = servidor.Text
sck.Close

sck.Connect DServer, 25
'sck.State = 7
''''Load sck
'''Enviar sck, nomde.Text, correode.Text, correopa.Text, objeto.Text

Timer1.Interval = 400
Timer1.Enabled = True


End Sub

Private Sub Form_Load()
servidor.Text = "correo.movinet.com.uy"
nomde.Text = "Computos"
correode.Text = "sapp03@movinet.com.uy"
correopa.Text = "sapp03@adinet.com.uy"
objeto.Text = "Información"
aUSUARIO = "sapp03"
aCONTRASEÑA = "306883"


End Sub

Private Sub sck_DataArrival(ByVal bytesTotal As Long)
    sck.GetData Respuesta
    Dim s As String
    Code = Left(Respuesta, 3)
'    AddStatus ColFrmStatus(Index), "<- " & Respuesta
    If Code >= 200 And Code <= 399 Then
        'Envío comando "EHLO"
         sck.SendData "EHLO " & DHelo & vbCrLf
         'Envío comando "AUTH LOGIN"
         sck.SendData "AUTH LOGIN" & vbCrLf
         s = Str_to_base64(aUSUARIO)
        'Envío nombre de usuario codificado en base64
         sck.SendData s & "=" & vbCrLf
         s = Str_to_base64(aCONTRASEÑA)
        'Envío contraseña codificado en base64
         sck.SendData s & "=" & vbCrLf
        'Envío MAIL FROM
         sck.SendData "MAIL FROM:<" & DMailFrom & ">" & vbCrLf
        'Envío RCPT TO (Destino del mail)
         sck.SendData "RCPT TO:<" & DRcptTo & ">" & vbCrLf
        'Envío comando DATA
         sck.SendData "DATA" & vbCrLf
        'Envío de datos del mail
        'DE
         sck.SendData "FROM: " & DFrom & vbCrLf
        'ASUNTO
         sck.SendData "SUBJECT: " & DSubject & vbCrLf
                'Envío aviso de alta prioridad si es necesario
'                If chkHigh.Value = 1 Then sck(Index).SendData "X-Priority: 1" & vbCrLf & "X-MSMail-Priority: High" & vbCrLf
                'Envío mensaje propiamente dicho
         sck.SendData DMensaje & vbCrLf
                                
                'Envío archivos adjuntos si existen
'                Dim i As Byte, Buff As String
'                If indexUUfiles > 0 Then
'                    For i = 0 To indexUUfiles
'                        Buff = Buff & UUfiles(i)
'                    Next i
'                    sck(Index).SendData Buff
'                End If
                
                'Envío comando FIN DE MENSAJE
                sck.SendData vbCrLf & "." & vbCrLf
                'Envío comando SALIR
                sck.SendData "QUIT" & vbCrLf
                sck.Close
                MsgBox "Fin"
    Else
        MsgBox "Error"
        sck.Close
    End If
    
End Sub

Private Sub Timer1_Timer()
DoEvents

End Sub
