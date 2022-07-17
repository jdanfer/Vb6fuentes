VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   DrawStyle       =   5  'Transparent
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtbody 
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   6135
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_envs 
      Caption         =   "data_envs"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Timer Timer10 
      Interval        =   5000
      Left            =   6600
      Top             =   1320
   End
   Begin VB.Timer timgral 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   7200
      Top             =   360
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   6600
      Top             =   2760
   End
   Begin VB.CheckBox chkHigh 
      Caption         =   "Alta Prioridad"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   1300
      Width           =   1335
   End
   Begin VB.CheckBox chkErr 
      Caption         =   "Mostrar Errores"
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   4570
      Width           =   2655
   End
   Begin VB.Timer tmrResumen 
      Left            =   5520
      Top             =   720
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar Envío"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkStatus 
      Caption         =   "Mostrar estado por cada socket"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtCant 
      Height          =   285
      Left            =   960
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "1"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CheckBox chkAuth 
      Caption         =   "El servidor requiere autenticación"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   140
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   270
      Left            =   480
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtUU 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4680
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSMTP 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdjuntar 
      Caption         =   "Adjuntar Archivo..."
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1635
      Width           =   4935
   End
   Begin VB.TextBox txtMailTo 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1275
      Width           =   2295
   End
   Begin VB.TextBox txtMailFrom 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   915
      Width           =   2295
   End
   Begin VB.TextBox txtFrom 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   555
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   5280
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4480
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "SMTP Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Mail To:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Mail From:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variable que contendrá el nombre de usuario (Si se usa autenticación)
Public aUSUARIO As String
'Variable que contendrá la contraseña (Si se usa autenticación)
Public aCONTRASEÑA As String

Public YaMandeAlgunaVez As Boolean
Public Cant As Byte

Dim frmResumen As frmStatus


Private Sub chkStatus_Click()
If chkStatus.Value = 1 Then
    chkErr.Value = 1
    chkErr.Enabled = False
Else
    chkErr.Enabled = True
End If
End Sub



Private Sub cmdAdjuntar_Click()


'Codifico el archivo en el formato válido para ser adjuntado a un mail
UUfiles(indexUUfiles) = UUEncodeFile("C:\datos\envios.zip") 'toda la ruta y nombre
' igual a 1
txtUU.Visible = True
indexUUfiles = indexUUfiles + 1
txtUU.Text = txtUU.Text & "envios.zip" & " (" & Fix(FileLen("C:\datos\envios.zip") / 1024) + 1 & " Kb)   "
'file title solo el nombre de archivo

End Sub

Private Sub cmdCancel_Click()
Call DesConectarTodos
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub


Private Sub cmdEnviar_Click()
Dim i As Byte
If Dir("c:\datos\envios.zip") <> "" Then
   FileCopy "c:\datos\envios.zip", "c:\datos\envtot.zip"
'   Kill ("c:\datos\envtot.mdb")

    txtUU.Text = ""
    indexUUfiles = 0
    'Codifico el archivo en el formato válido para ser adjuntado a un mail
    UUfiles(indexUUfiles) = UUEncodeFile("C:\datos\envios.zip") 'toda la ruta y nombre
    ' igual a 1
    txtUU.Visible = True
    indexUUfiles = indexUUfiles + 1
    txtUU.Text = txtUU.Text & "envios.zip" & " (" & Fix(FileLen("C:\datos\envios.zip") / 1024) + 1 & " Kb)   "
    'file title solo el nombre de archivo
    If chkAuth.Value = 1 Then
        If Data1.Recordset("base") = 1 Then
           aUSUARIO = "sapp01"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 2 Then
           aUSUARIO = "sapp003"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 3 Then
           aUSUARIO = "sapp33"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 4 Then
           aUSUARIO = "sapp04"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 6 Then
           aUSUARIO = "sapp66"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 8 Then
           aUSUARIO = "sapp08"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 10 Then
           aUSUARIO = "sapp010"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 13 Then
           aUSUARIO = "sapp013"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 15 Then
           aUSUARIO = "sapp015"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 16 Then
           aUSUARIO = "sapp16"
           aCONTRASEÑA = "sapp1987"
        End If
        If Data1.Recordset("base") = 17 Then
           aUSUARIO = "sapp17"
           aCONTRASEÑA = "sapp1987"
        End If
        
    End If
    
    If Val(txtCant) < 1 Or txtSMTP = "" Or txtFrom = "" Or txtMailFrom = "" Or txtMailTo = "" Then
        MsgBox "Datos incompletos", vbCritical, "Error"
        Exit Sub
    End If
    
    If Val(txtCant) > 255 Then
        MsgBox "Se pueden enviar un máximo de 255 mails", vbCritical, "Error"
        txtCant.SetFocus
        Exit Sub
    End If
    
    If txtSubject = "" And txtbody = "" Then
        MsgBox "Debe escribir un Asunto o un Mensaje", vbCritical, "Error"
        txtSubject.SetFocus
        Exit Sub
    End If
    
    ''If MsgBox("¿Confirma envío?", vbYesNo Or vbQuestion, "") = vbNo Then Exit Sub
    
    DServer = txtSMTP
    
    cmdCancel.Visible = True
    
    If YaMandeAlgunaVez Then
        For i = 2 To sck.Count
            Unload sck(i - 1)
        Next i
        For i = LBound(ColFrmStatus) To UBound(ColFrmStatus)
            Unload ColFrmStatus(i)
        Next i
    End If
    
    Cant = Val(txtCant)
    ReDim ColFrmStatus(Cant - 1)
    
    For i = 0 To Cant - 1
        If i <> 0 Then
            Load sck(i)
        End If
        Set ColFrmStatus(i) = New frmStatus
        ColFrmStatus(i).Caption = "Status " & i + 1
        ColFrmStatus(i).txtStatus = ""
        If chkStatus.Value = 1 Then ColFrmStatus(i).Show
        YaMandeAlgunaVez = True
        Enviar sck(i), txtFrom, txtMailFrom, txtMailTo, txtSubject, txtbody.Text
    Next i
    
    frmResumen.txtStatus = ""
    frmResumen.Caption = "Resumen de envíos"
    frmResumen.Show
    tmrResumen.Interval = 400
    tmrResumen.Enabled = True
Else
    Contatime = 0
    timgral.Enabled = True
End If

End Sub

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

Private Sub Form_Activate()
txtFrom.SetFocus
exCaption = Me.Caption
End Sub

Private Sub Form_Load()
indexUUfiles = 0
YaMandeAlgunaVez = False
Set frmResumen = New frmStatus
TextStatus(0) = "Cerrado"
TextStatus(1) = "Conectado"
TextStatus(2) = "Autorizando"
TextStatus(3) = "Autorizando"
TextStatus(4) = "Autorizando"
TextStatus(5) = "Enviando 'From'"
TextStatus(6) = "Enviando 'To'"
TextStatus(7) = "Enviando Datos"
TextStatus(8) = "Enviando Mensaje"
TextStatus(9) = "Finalizando"
TextStatus(10) = "Finalizado OK !"
TextStatus(11) = "Finalizado Con Errores !"

Data2.DatabaseName = App.Path & "\conecta.mdb"
Data2.RecordSource = "conectado"
Data2.Refresh

Data1.DatabaseName = App.Path & "\parsec0.mdb"
Data1.RecordSource = "parsec0"
Data1.Refresh
If Data1.Recordset("base") = 1 Then
   txtSubject.Text = "Envio B1"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base 1"
   txtMailFrom.Text = "sapp01@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If

If Data1.Recordset("base") = 3 Then
   txtSubject.Text = "Envio B3"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base 3"
   txtMailFrom.Text = "sapp33@adinet.com.uy"
   txtMailTo.Text = "sapp01@adinet.com.uy"
End If
If Data1.Recordset("base") = 4 Then
   txtSubject.Text = "Envio B4"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base 4"
   txtMailFrom.Text = "sapp04@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 6 Then
   txtSubject.Text = "Envio B6"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base 6"
   txtMailFrom.Text = "sapp66@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 8 Then
   txtSubject.Text = "Envio B8"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base 8"
   txtMailFrom.Text = "sapp08@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 17 Then
   txtSubject.Text = "Envio B17"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base B17"
   txtMailFrom.Text = "sapp17@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 13 Then
   txtSubject.Text = "Envio B13"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base B13"
   txtMailFrom.Text = "sapp013@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 15 Then
   txtSubject.Text = "Envio B15"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base B15"
   txtMailFrom.Text = "sapp015@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 16 Then
   txtSubject.Text = "Envio B16"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base B16"
   txtMailFrom.Text = "sapp16@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 2 Then
   txtSubject.Text = "Envio B2"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base B2"
   txtMailFrom.Text = "sapp003@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 10 Then
   txtSubject.Text = "Envio B10"
   txtSMTP.Text = "adinet.com.uy"
   txtFrom.Text = "Base B10"
   txtMailFrom.Text = "sapp010@adinet.com.uy"
   txtMailTo.Text = "sapp33@adinet.com.uy"
End If

      
End Sub

Private Sub Form_Unload(Cancel As Integer)
'End
End Sub

Private Sub sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'Esta es la parte más importante, donde se produce el diálogo con el servidor SMTP
    
    sck(Index).GetData Respuesta
    Dim s As String
    Code = Left(Respuesta, 3)
    AddStatus ColFrmStatus(Index), "<- " & Respuesta
    If Code >= 200 And Code <= 399 Then
        Select Case SendStatus(Index)
            Case CONECTED
                'Envío comando "EHLO"
                sck(Index).SendData "EHLO " & DHelo & vbCrLf
                If chkAuth.Value = 1 Then
                    'Si estoy usando autenticación
                    SendStatus(Index) = AUTH1
                Else
                    'Si no uso autenticación
                    SendStatus(Index) = MailFrom
                End If
            Case AUTH1
                'Envío comando "AUTH LOGIN"
                sck(Index).SendData "AUTH LOGIN" & vbCrLf
                AddStatus ColFrmStatus(Index), ("-> AUTH LOGIN")
                SendStatus(Index) = AUTH2
            Case AUTH2
                s = Str_to_base64(aUSUARIO)
                'Envío nombre de usuario codificado en base64
                sck(Index).SendData s & "=" & vbCrLf
                AddStatus ColFrmStatus(Index), ("-> Usuario: " & s)
                SendStatus(Index) = AUTH3
            Case AUTH3
                s = Str_to_base64(aCONTRASEÑA)
                'Envío contraseña codificado en base64
                sck(Index).SendData s & "=" & vbCrLf
                AddStatus ColFrmStatus(Index), ("-> Contraseña: " & s)
                SendStatus(Index) = MailFrom
           Case MailFrom
                'Envío MAIL FROM
                sck(Index).SendData "MAIL FROM:<" & DMailFrom & ">" & vbCrLf
                AddStatus ColFrmStatus(Index), ("-> MAIL FROM:<" & DMailFrom & ">")
                SendStatus(Index) = RCPTTO
            Case RCPTTO
                'Envío RCPT TO (Destino del mail)
                sck(Index).SendData "RCPT TO:<" & DRcptTo & ">" & vbCrLf
                AddStatus ColFrmStatus(Index), ("-> RCPT TO:<" & DRcptTo & ">")
                SendStatus(Index) = DATAC
            Case DATAC
               'Envío comando DATA
                sck(Index).SendData "DATA" & vbCrLf
                SendStatus(Index) = MESSAGGE
            Case MESSAGGE
                'Envío de datos del mail
                'DE
                sck(Index).SendData "FROM: " & DFrom & vbCrLf
                AddStatus ColFrmStatus(Index), ("-> FROM: " & DFrom)
                'ASUNTO
                sck(Index).SendData "SUBJECT: " & DSubject & vbCrLf
                AddStatus ColFrmStatus(Index), ("-> SUBJECT: " & DSubject)
                'Envío aviso de alta prioridad si es necesario
                If chkHigh.Value = 1 Then sck(Index).SendData "X-Priority: 1" & vbCrLf & "X-MSMail-Priority: High" & vbCrLf
                'Envío mensaje propiamente dicho
                sck(Index).SendData DMensaje & vbCrLf
                                
                'Envío archivos adjuntos si existen
                Dim i As Byte, Buff As String
                If indexUUfiles > 0 Then
                    For i = 0 To indexUUfiles
                        Buff = Buff & UUfiles(i)
                    Next i
                    sck(Index).SendData Buff
                End If
                
                'Envío comando FIN DE MENSAJE
                sck(Index).SendData vbCrLf & "." & vbCrLf
                
               SendStatus(Index) = QUIT
           Case QUIT
               AddStatus ColFrmStatus(Index), "*** MAIL ENVIADO OK ***"
               ColFrmStatus(Index).Hide
                'Envío comando SALIR
                sck(Index).SendData "QUIT" & vbCrLf
               SendStatus(Index) = MANDADO_OK
                DesConectar sck(Index)
                If Contatime = 88 Then
                   Contatime = 0
                   timgral.Enabled = True
                                      
'                   End
                End If
       End Select
    Else
        SendStatus(Index) = cERROR
        If chkErr.Value = 1 Then
            ColFrmStatus(Index).Caption = ColFrmStatus(Index).Caption & " (Con errores)"
            ColFrmStatus(Index).Show
        End If
        DesConectar sck(Index)
'        MsgBox "Error"
    End If

End Sub

Private Sub sck_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddStatus ColFrmStatus(Index), "Error nº:" & Number & " " & Description
    SendStatus(Index) = cERROR
    DesConectar sck(Index)
End Sub

Private Sub Timer1_Timer()
Contatime = Contatime + 1
If Contatime = 3 Then
   If Data1.Recordset("base") <> 3 Then
      Contatime = 0
      Timer1.Enabled = False
      Contatime = 88
      cmdEnviar_Click
   Else
      txtSubject.Text = "Envio B3"
      txtSMTP.Text = "adinet.com.uy"
      txtFrom.Text = "Base 3"
      txtMailFrom.Text = "sapp33@adinet.com.uy"
      txtMailTo.Text = "sapp66@adinet.com.uy"
      Contatime = 0
'0
      cmdEnviar_Click
      Timer1.Enabled = False
'true
   End If
End If

End Sub

Private Sub Timer10_Timer()
Timer10.Enabled = False
If Conectado = True Then
   Data2.Recordset.Edit
   Data2.Recordset("siono") = "SI"
   Data2.Recordset.Update
   Timer1.Enabled = True
Else
   Data2.Recordset.Edit
   Data2.Recordset("siono") = "NO"
   Data2.Recordset.Update
   End
End If

'correo.Show
'Unload frmMain

End Sub


Private Sub timgral_Timer()
Contatime = Contatime + 1
If Contatime = 12 Then
   Contatime = 0
'   MsgBox "Comienza otra vez"
'   Timer10.Enabled = True
'   Unload Me
   Data1.Recordset.Edit
   Data1.Recordset("ult_env") = Data1.Recordset("ult_env") + 1
   Data1.Recordset.Update
   If Dir("c:\datos\envios.zip") <> "" Then
      Kill "c:\datos\envios.zip"
   End If
   timgral.Enabled = False
   End
   
'   frmMain.Hide
End If

End Sub

Private Sub tmrResumen_Timer()
DoEvents
DoRefresh False
End Sub

Public Sub DoRefresh(FinTodos As Boolean)
'Hace el refresh de las ventanas resúmenes (frmStatus)
Dim i As Byte, Posi As Byte
frmResumen.txtStatus = ""
For i = 0 To Cant - 1
    frmResumen.txtStatus = frmResumen.txtStatus & "Socket " & i + 1 & " (" & IIf(SendStatus(i) > 10, 10, SendStatus(i)) & "/10) - " & TextStatus(SendStatus(i)) & vbCrLf
    Posi = Posi + IIf(SendStatus(i) = MANDADO_OK, 1, 0)
Next i
If FinTodos Then
    frmResumen.txtStatus = frmResumen.txtStatus & vbCrLf & "Enviados Correctamente: " & Posi
    frmResumen.txtStatus = frmResumen.txtStatus & vbCrLf & "Con Errores: " & Cant - Posi
End If
End Sub

Private Sub txtBody_KeyDown(KeyCode As Integer, Shift As Integer)
'Esto es para que tocando la tecla TAB, en el cuadro de texto del cuerpo
'del mensaje, se produzca una tabulación y no un avance del foco
Dim i As Long
If Shift <> 0 Then Exit Sub
If KeyCode = 9 Then
    i = txtbody.SelStart
    txtbody.Text = Left(txtbody.Text, i) & Chr(9) & Mid(txtbody.Text, i + 1)
    txtbody.SelStart = i + 1
    KeyCode = 0
End If
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
'Sólo permite el ingreso de númeors
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub


