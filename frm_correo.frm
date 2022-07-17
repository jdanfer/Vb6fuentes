VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_correo 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Envio y Recepción"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   FillStyle       =   0  'Solid
   Icon            =   "frm_correo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Data data_pabase3 
      Caption         =   "data_pabase3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   3000
      Top             =   1080
   End
   Begin VB.Data data_temp 
      Caption         =   "data_temp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   2160
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   1920
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   2760
      Top             =   2040
   End
   Begin VB.Data data_env 
      Caption         =   "data_env"
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   1440
   End
   Begin VB.Data data_rec 
      Caption         =   "data_rec"
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   4200
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
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PARSEC0"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   12000
      Left            =   4080
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   2520
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3360
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   120
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1920
      Picture         =   "frm_correo.frx":030A
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Aguarde... Procesando actualizaciones..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   2055
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frm_correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim fecha As Date
Dim Archivo As String
Dim Remite As String
Dim Asunto As String
'Envio de e-mail desde VB:
'1.- Adjuntar al proyecto los controles MAPI
'(ya sabes: Proyecto/Componentes y señalar Microsoft MAPI controls)
'2.- En tu formulario, coloca los controles MAPISession y MAPIMessages
'3.- Para enviar el mail:
fecha = Data1.Recordset("ult_envio")
If Month(fecha) > 9 Then
   If Day(fecha) > 9 Then
      Archivo = Trim(Str(Year(Format(fecha, "dd/mm/yyyy")))) + Trim(Str(Month(Format(fecha, "dd/mm/yyyy")))) + Trim(Str(Day(Format(fecha, "dd/mm/yyyy")))) + ".zip"
   Else
      Archivo = Trim(Str(Year(Format(fecha, "dd/mm/yyyy")))) + Trim(Str(Month(Format(fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(fecha, "dd/mm/yyyy")))) + ".zip"
   End If
Else
   If Day(fecha) > 9 Then
      Archivo = Trim(Str(Year(Format(fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(fecha, "dd/mm/yyyy")))) + Trim(Str(Day(Format(fecha, "dd/mm/yyyy")))) + ".zip"
   Else
      Archivo = Trim(Str(Year(Format(fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(fecha, "dd/mm/yyyy")))) + ".zip"
   End If
End If
If Data1.Recordset("base") = 1 Then
   Remite = "sapp01@adinet.com.uy"
End If
If Data1.Recordset("base") = 3 Then
   Remite = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 4 Then
   Remite = "sapp04@adinet.com.uy"
End If
If Data1.Recordset("base") = 6 Then
   Remite = "sapp66@adinet.com.uy"
End If
If Data1.Recordset("base") = 8 Then
   Remite = "sapp08@adinet.com.uy"
End If
If Data1.Recordset("base") = 10 Then
   Remite = "sapp010@adinet.com.uy"
End If
If Data1.Recordset("base") = 13 Then
   Remite = "sapp013@adinet.com.uy"
End If
If Data1.Recordset("base") = 15 Then
   Remite = "sapp015@adinet.com.uy"
End If
If Data1.Recordset("base") = 16 Then
   Remite = "sapp16@adinet.com.uy"
End If
If Data1.Recordset("base") = 17 Then
   Remite = "sapp17@adinet.com.uy"
End If

MAPISession1.UserName = "sapp33@adinet.com.uy"
MAPISession1.NewSession = True
MAPISession1.DownLoadMail = False ' o false si no deseas recibir
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID

MAPIMessages1.MsgIndex = -1 ' nuevo mensaje
MAPIMessages1.RecipDisplayName = "sapp33@adinet.com.uy"

Asunto = "Info Base " + Trim(Str(Data1.Recordset("base")))
MAPIMessages1.MsgSubject = Asunto
MAPIMessages1.MsgNoteText = ""

' si deseas anexar algun archivo al mail:
MAPIMessages1.AttachmentIndex = MAPIMessages1.AttachmentCount
MAPIMessages1.AttachmentName = Archivo
MAPIMessages1.AttachmentPathName = "C:\Enviar\" + Trim(Archivo)
MAPIMessages1.AttachmentPosition = MAPIMessages1.AttachmentIndex
MAPIMessages1.AttachmentType = vbAttachTypeData

' (puedes anexar varios archivos, incrementando el numero 0,1,2,3....)
' Y por fin, enviarlo:
MAPIMessages1.Send

' Cuando ya no tengas que enviar ningun mail más:
MAPISession1.SignOff

End Sub

Private Sub Command2_Click()
Dim Bajaarch As String
Dim Fecrec As Date
Dim Remibaj As String
Dim nCanMsg As Integer
Dim cNomFic As String
Dim nX As Integer
Dim nY As Integer
Fecrec = Data1.Recordset("ult_recep") + 1
If Month(Fecrec) > 9 Then
   If Day(Fecrec) > 9 Then
      Bajaarch = Trim(Str(Year(Format(Fecrec, "dd/mm/yyyy")))) + Trim(Str(Month(Format(Fecrec, "dd/mm/yyyy")))) + Trim(Str(Day(Format(Fecrec, "dd/mm/yyyy")))) + ".zip"
   Else
      Bajaarch = Trim(Str(Year(Format(Fecrec, "dd/mm/yyyy")))) + Trim(Str(Month(Format(Fecrec, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(Fecrec, "dd/mm/yyyy")))) + ".zip"
   End If
Else
   If Day(Fecrec) > 9 Then
      Bajaarch = Trim(Str(Year(Format(Fecrec, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(Fecrec, "dd/mm/yyyy")))) + Trim(Str(Day(Format(Fecrec, "dd/mm/yyyy")))) + ".zip"
   Else
      Bajaarch = Trim(Str(Year(Format(Fecrec, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(Fecrec, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(Fecrec, "dd/mm/yyyy")))) + ".zip"
   End If
End If

If Data1.Recordset("base") = 1 Then
   Remibaj = "sapp01@adinet.com.uy"
End If
If Data1.Recordset("base") = 3 Then
   Remibaj = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 4 Then
   Remibaj = "sapp04@adinet.com.uy"
End If
If Data1.Recordset("base") = 6 Then
   Remibaj = "sapp66@adinet.com.uy"
End If
If Data1.Recordset("base") = 8 Then
   Remibaj = "sapp08@adinet.com.uy"
End If
If Data1.Recordset("base") = 10 Then
   Remibaj = "sapp010@adinet.com.uy"
End If
If Data1.Recordset("base") = 13 Then
   Remibaj = "sapp013@adinet.com.uy"
End If
If Data1.Recordset("base") = 15 Then
   Remibaj = "sapp015@adinet.com.uy"
End If
If Data1.Recordset("base") = 16 Then
   Remibaj = "sapp16@adinet.com.uy"
End If
If Data1.Recordset("base") = 17 Then
   Remibaj = "sapp17@adinet.com.uy"
End If

MAPISession1.UserName = Remibaj
MAPISession1.NewSession = True
MAPISession1.DownLoadMail = True
MAPISession1.SignOn

MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.FetchUnreadOnly = True ' Solo los no leidos
MAPIMessages1.FetchSorted = True ' ordenados segun llegada
MAPIMessages1.Fetch ' obtiene el conjunto de mensajes

nCanMsg = MAPIMessages1.MsgCount - 1
For nX = 0 To nCanMsg
MAPIMessages1.MsgIndex = nX
' Filtrado de los mensajes para seleccionar los deseados segun el asunto
If MAPIMessages1.MsgOrigAddress = "sapp33@adinet.com.uy" Then
' Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
' Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
For nY = 0 To MAPIMessages1.AttachmentCount - 1
MAPIMessages1.AttachmentIndex = nY
cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
FileCopy MAPIMessages1.AttachmentPathName, "C:\Recibir\Base_06\" + cNomFic
Next
' borrado del mensaje (si queremos hacerlo)
MAPIMessages1.Delete (mapMessageDelete)
End If
Next
' Cerrar las sesion
MAPISession1.SignOff

End Sub

Private Sub Command3_Click()
Dim fecha As Date
Dim Archivo As String
Dim Remite As String
Dim Asunto As String
Dim Bajaarch As String
Dim Fecrec As Date
Dim Remibaj As String
Dim nCanMsg As Integer
Dim cNomFic As String
Dim nX As Integer
Dim nY As Integer
'Envio de e-mail desde VB:
'1.- Adjuntar al proyecto los controles MAPI
'(ya sabes: Proyecto/Componentes y señalar Microsoft MAPI controls)
'2.- En tu formulario, coloca los controles MAPISession y MAPIMessages
'3.- Para enviar el mail:
fecha = Data1.Recordset("ult_envio")
If Month(fecha) > 9 Then
   If Day(fecha) > 9 Then
      Archivo = Trim(Str(Year(Format(fecha, "dd/mm/yyyy")))) + Trim(Str(Month(Format(fecha, "dd/mm/yyyy")))) + Trim(Str(Day(Format(fecha, "dd/mm/yyyy")))) + ".zip"
   Else
      Archivo = Trim(Str(Year(Format(fecha, "dd/mm/yyyy")))) + Trim(Str(Month(Format(fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(fecha, "dd/mm/yyyy")))) + ".zip"
   End If
Else
   If Day(fecha) > 9 Then
      Archivo = Trim(Str(Year(Format(fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(fecha, "dd/mm/yyyy")))) + Trim(Str(Day(Format(fecha, "dd/mm/yyyy")))) + ".zip"
   Else
      Archivo = Trim(Str(Year(Format(fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(fecha, "dd/mm/yyyy")))) + ".zip"
   End If
End If
If Data1.Recordset("base") = 1 Then
   Remite = "sapp01@adinet.com.uy"
End If
If Data1.Recordset("base") = 3 Then
   Remite = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 4 Then
   Remite = "sapp04@adinet.com.uy"
End If
If Data1.Recordset("base") = 6 Then
   Remite = "sapp66@adinet.com.uy"
End If
If Data1.Recordset("base") = 8 Then
   Remite = "sapp08@adinet.com.uy"
End If
If Data1.Recordset("base") = 10 Then
   Remite = "sapp010@adinet.com.uy"
End If
If Data1.Recordset("base") = 13 Then
   Remite = "sapp013@adinet.com.uy"
End If
If Data1.Recordset("base") = 15 Then
   Remite = "sapp015@adinet.com.uy"
End If
If Data1.Recordset("base") = 16 Then
   Remite = "sapp16@adinet.com.uy"
End If
If Data1.Recordset("base") = 17 Then
   Remite = "sapp17@adinet.com.uy"
End If
Fecrec = Data1.Recordset("ult_recep") + 1
If Month(Fecrec) > 9 Then
   If Day(Fecrec) > 9 Then
      Bajaarch = Trim(Str(Year(Format(Fecrec, "dd/mm/yyyy")))) + Trim(Str(Month(Format(Fecrec, "dd/mm/yyyy")))) + Trim(Str(Day(Format(Fecrec, "dd/mm/yyyy")))) + ".zip"
   Else
      Bajaarch = Trim(Str(Year(Format(Fecrec, "dd/mm/yyyy")))) + Trim(Str(Month(Format(Fecrec, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(Fecrec, "dd/mm/yyyy")))) + ".zip"
   End If
Else
   If Day(Fecrec) > 9 Then
      Bajaarch = Trim(Str(Year(Format(Fecrec, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(Fecrec, "dd/mm/yyyy")))) + Trim(Str(Day(Format(Fecrec, "dd/mm/yyyy")))) + ".zip"
   Else
      Bajaarch = Trim(Str(Year(Format(Fecrec, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(Fecrec, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(Fecrec, "dd/mm/yyyy")))) + ".zip"
   End If
End If
If Data1.Recordset("base") = 1 Then
   Remibaj = "sapp01@adinet.com.uy"
End If
If Data1.Recordset("base") = 3 Then
   Remibaj = "sapp33@adinet.com.uy"
End If
If Data1.Recordset("base") = 4 Then
   Remibaj = "sapp04@adinet.com.uy"
End If
If Data1.Recordset("base") = 6 Then
   Remibaj = "sapp66@adinet.com.uy"
End If
If Data1.Recordset("base") = 8 Then
   Remibaj = "sapp08@adinet.com.uy"
End If
If Data1.Recordset("base") = 10 Then
   Remibaj = "sapp010@adinet.com.uy"
End If
If Data1.Recordset("base") = 13 Then
   Remibaj = "sapp013@adinet.com.uy"
End If
If Data1.Recordset("base") = 15 Then
   Remibaj = "sapp015@adinet.com.uy"
End If
If Data1.Recordset("base") = 16 Then
   Remibaj = "sapp16@adinet.com.uy"
End If
If Data1.Recordset("base") = 17 Then
   Remibaj = "sapp17@adinet.com.uy"
End If
MAPISession1.UserName = "sapp33@adinet.com.uy"
MAPISession1.NewSession = True
MAPISession1.DownLoadMail = True ' o false si no deseas recibir
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID

MAPIMessages1.MsgIndex = -1 ' nuevo mensaje
MAPIMessages1.RecipDisplayName = "sapp33@adinet.com.uy"
Asunto = "Info Base " + Trim(Str(Data1.Recordset("base")))
MAPIMessages1.MsgSubject = Asunto
MAPIMessages1.MsgNoteText = ""

' si deseas anexar algun archivo al mail:
MAPIMessages1.AttachmentIndex = MAPIMessages1.AttachmentCount
MAPIMessages1.AttachmentName = Archivo
MAPIMessages1.AttachmentPathName = "C:\Enviar\" + Trim(Archivo)
MAPIMessages1.AttachmentPosition = MAPIMessages1.AttachmentIndex
MAPIMessages1.AttachmentType = vbAttachTypeData

MAPIMessages1.Send

MAPIMessages1.FetchUnreadOnly = True ' Solo los no leidos
MAPIMessages1.FetchSorted = True ' ordenados segun llegada
MAPIMessages1.Fetch ' obtiene el conjunto de mensajes

nCanMsg = MAPIMessages1.MsgCount - 1
nCanMsg = 1
For nX = 0 To nCanMsg
MAPIMessages1.MsgIndex = nX
' Filtrado de los mensajes para seleccionar los deseados segun el asunto
If MAPIMessages1.MsgOrigAddress = "sapp33@adinet.com.uy" Then
' Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
' Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
For nY = 0 To nCanMsg
MAPIMessages1.AttachmentIndex = nY
cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
FileCopy MAPIMessages1.AttachmentPathName, "C:\Recibir\Base_06\" + cNomFic
Next
' borrado del mensaje (si queremos hacerlo)
MAPIMessages1.Delete (mapMessageDelete)
End If
Next
' Cerrar las sesion
MAPISession1.SignOff
'Timer1.Enabled = False
' Cuando ya no tengas que enviar ningun mail más:
MAPISession1.SignOff

End Sub

Private Sub Form_Activate()
'If Not Conectado Then
'   Timer3.Enabled = False
'   MsgBox "No hay conexión a internet", vbCritical, "Mensaje"
'   Unload Me
'End If

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\parsec0.mdb"
Data1.RecordSource = "parsec0"
Data1.Refresh
data_parsec.DatabaseName = App.Path & "\parsec0.mdb"
data_parsec.RecordSource = "parsec0"
data_parsec.Refresh
Data2.DatabaseName = App.Path & "\conecta.mdb"
Data2.RecordSource = "conectado"
Data2.Refresh
data_pabase3.DatabaseName = App.Path & "\sapp.mdb"
data_pabase3.RecordSource = "parsec0"
data_pabase3.Refresh

End Sub

Private Sub Timer1_Timer()
Shell App.Path & "\conecta.exe", vbMinimizedNoFocus

Timer1.Enabled = False
Timer8.Enabled = True

'Rei = ExitWindowsEx(2, 0&) 'Reinicia el Sistema

End Sub

Private Sub Timer2_Timer()
Dim Bajaarch As String
Dim Remibaj As String
Dim nCanMsg As Integer
Dim cNomFic As String
Dim nX As Integer
Dim nY As Integer
Dim Verarch As String
Dim Contra As String
Bajaarch = "envios.zip"
    If Data1.Recordset("base") = 1 Then
       Remibaj = "sapp01@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 2 Then
       Remibaj = "sapp003@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 3 Then
       Remibaj = "sapp33@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 4 Then
       Remibaj = "sapp04@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 6 Then
       Remibaj = "sapp66@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 8 Then
       Remibaj = "sapp08@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 10 Then
       Remibaj = "sapp010@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 13 Then
       Remibaj = "sapp013@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 15 Then
       Remibaj = "sapp015@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 16 Then
       Remibaj = "sapp16@adinet.com.uy"
       Contra = "sapp1987"
    End If
    If Data1.Recordset("base") = 17 Then
       Remibaj = "sapp17@adinet.com.uy"
       Contra = "sapp1987"
    End If
    MAPISession1.UserName = Remibaj
    MAPISession1.NewSession = True
    MAPISession1.DownLoadMail = True
'    MAPISession1.SignOff
    MAPISession1.SignOn
    
    MAPIMessages1.SessionID = MAPISession1.SessionID
    MAPIMessages1.FetchUnreadOnly = True ' Solo los no leidos
    MAPIMessages1.FetchSorted = True  ' ordenados segun llegada
    
    MAPIMessages1.Fetch ' obtiene el conjunto de mensajes
    
'    MAPIMessages1.FetchMsgType
    
    nCanMsg = MAPIMessages1.MsgCount - 1
    If nCanMsg >= 0 Then
'    For nX = 0 To nCanMsg
        nX = 0
        MAPIMessages1.MsgIndex = 0
    '   Filtrado de los mensajes para seleccionar los deseados segun el asunto
'        If MAPIMessages1.MsgOrigAddress = "sapp03@movinet.com.uy" Then
        If MAPIMessages1.MsgSubject = "Envio B3" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\" + cNomFic
              Next
    'borrado del mensaje (si queremos hacerlo)
'           MAPIMessages1.Delete (mapMessageDelete)
    
          MAPIMessages1.FetchUnreadOnly = False
          MAPIMessages1.Fetch ' obtiene el conjunto de mensajes
    
          nX = 0
          MAPIMessages1.MsgIndex = 0
          MAPIMessages1.Delete (mapMessageDelete)
    
          MAPISession1.SignOff
          Timer2.Enabled = False
          Timer5.Enabled = True
        Else
    
           MAPIMessages1.FetchUnreadOnly = False
           MAPIMessages1.Fetch ' obtiene el conjunto de mensajes
    
           nX = 0
           MAPIMessages1.MsgIndex = 0
           MAPIMessages1.Delete (mapMessageDelete)
    
           MAPISession1.SignOff
           Timer2.Enabled = True
           Timer5.Enabled = False
'           frm_correo.MousePointer = 0
           
'           Unload Me
      End If
'      MAPISession1.SignOff
    
    Else
       MAPISession1.SignOff
       Timer2.Enabled = False
'       Timer5.Enabled = True
       Timer5.Enabled = False
       frm_correo.MousePointer = 0
       Unload Me
    End If
'    Next
    
'End If

      
End Sub

Private Function ExtraerNombreArchivo(cArchivo As String) As String
' extrae el nombre de un archivo de una cadena con path completo
Dim nX As Integer
ExtraerNombreArchivo = ""
For nX = Len(cArchivo) To 1 Step -1
If Not Mid(cArchivo, nX, 1) = "\" Then
ExtraerNombreArchivo = Mid(cArchivo, nX, 1) + ExtraerNombreArchivo
Else
Exit For 'salir del bucle, ya esta.
End If
Next
End Function

Private Sub Timer3_Timer()
Dim Fecenv As Date
Dim Xdias, Xdiaste As Long
Dim XCanD As Integer
Dim XFecrecib, XFecenvte As Date
If Data1.Recordset("base") <> 3 Then
    frm_correo.MousePointer = 11
    FileCopy "c:\datos\vacios\env_clia.dbf", "C:\datos\env_clia.dbf"
    
    FileCopy "c:\datos\vacios\env_clib.dbf", "C:\datos\env_clib.dbf"
    
    FileCopy "c:\datos\vacios\env_clim.dbf", "C:\datos\env_clim.dbf"
    
    FileCopy "c:\datos\vacios\env_lin.dbf", "C:\datos\env_lin.dbf"
    
    FileCopy "c:\datos\vacios\env_caja.dbf", "C:\datos\env_caja.dbf"
    
    FileCopy "c:\datos\vacios\env_abm.dbf", "C:\datos\env_abm.dbf"
    
    FileCopy "c:\datos\vacios\env_tes.dbf", "C:\datos\env_tes.dbf"
    
    FileCopy "c:\datos\vacios\env_lla.dbf", "C:\datos\env_lla.dbf"
    
    FileCopy "c:\datos\vacios\env_conv.dbf", "C:\datos\env_conv.dbf"
    FileCopy "c:\datos\vacios\env_conv.mdx", "C:\datos\env_conv.mdx"
    
    FileCopy "c:\datos\vacios\env_estu.dbf", "C:\datos\env_estu.dbf"
    FileCopy "c:\datos\vacios\env_estu.mdx", "C:\datos\env_estu.mdx"
    
    FileCopy "c:\datos\vacios\env_arq.dbf", "C:\datos\env_arq.dbf"
    FileCopy "c:\datos\vacios\env_arq.mdx", "C:\datos\env_arq.mdx"
        
    FileCopy "c:\datos\vacios\env_cob.dbf", "C:\datos\env_cob.dbf"
    
    FileCopy "c:\datos\vacios\env_codc.dbf", "C:\datos\env_codc.dbf"
    
    FileCopy "c:\datos\vacios\env_rubt.dbf", "C:\datos\env_rubt.dbf"
    
    If data_parsec.Recordset("base") <> 3 Then
       Fecenv = data_parsec.Recordset("ult_env") + 1
       Xdias = Date - data_parsec.Recordset("ult_env")
       XCanD = 1
    Else
       Fecenv = data_parsec.Recordset("ult_env") + 1
       Xdias = Date - data_parsec.Recordset("ult_env")
       XFecenvte = data_parsec.Recordset("ult_env3") + 1
       Xdiaste = Date - data_parsec.Recordset("ult_env3")
       XFecrecib = data_parsec.Recordset("ult_env") + 1
    End If
    
    If Xdias > 1 And XCanD = 1 Then
        data_env.DatabaseName = "c:\Datos"
        data_env.RecordSource = "env_clia"
        data_env.Refresh
        XLafecha = Fecenv
        If data_env.Recordset.RecordCount > 0 Then
           data_env.Recordset.MoveFirst
           Do While Not data_env.Recordset.EOF
              data_env.Recordset.Delete
              data_env.Recordset.MoveNext
           Loop
        End If
        data_cli.DatabaseName = App.Path & "\sapp.mdb"
        data_cli.RecordSource = "clientes"
        data_cli.Refresh
        data_cli.RecordSource = "Select * from clientes where fecha_sys =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
        data_cli.Refresh
        If data_cli.Recordset.RecordCount > 0 Then
           data_cli.Recordset.MoveFirst
           Do While Not data_cli.Recordset.EOF
                data_env.Recordset.AddNew
                data_env.Recordset("estado") = data_cli.Recordset("estado")
                data_env.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                data_env.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                data_env.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                data_env.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                If IsNull(data_cli.Recordset("cl_cedula")) = False Then
                   data_env.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                Else
                   data_env.Recordset("cl_cedula") = 0
                End If
                If IsNull(data_cli.Recordset("cl_codced")) = False Then
                   data_env.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                Else
                   data_env.Recordset("cl_codced") = 0
                End If
                If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                   data_env.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                End If
                If IsNull(data_cli.Recordset("cl_edad")) = False Then
                   data_env.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                Else
                   data_env.Recordset("cl_edad") = 0
                End If
                If IsNull(data_cli.Recordset("cl_uniedad")) = False Then
                   data_env.Recordset("cl_uniedad") = data_cli.Recordset("cl_uniedad")
                Else
                   data_env.Recordset("cl_uniedad") = "A"
                End If
                If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
                   data_env.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                Else
                   data_env.Recordset("cl_ultmesp") = 0
                End If
                If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
                   data_env.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                Else
                   data_env.Recordset("cl_ultanop") = 0
                End If
                If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
                   data_env.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa")
                Else
                   data_env.Recordset("cl_atrasoa") = 0
                End If
                If IsNull(data_cli.Recordset("saldo_cc")) = False Then
                   data_env.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                Else
                   data_env.Recordset("saldo_cc") = 0
                End If
                If IsNull(data_cli.Recordset("cl_direcci")) = False Then
                   data_env.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                Else
                   data_env.Recordset("cl_direcci") = ""
                End If
                If IsNull(data_cli.Recordset("cl_entre")) = False Then
                   data_env.Recordset("cl_entre") = data_cli.Recordset("cl_entre")
                Else
                   data_env.Recordset("cl_entre") = ""
                End If
                If IsNull(data_cli.Recordset("cl_grupo")) = False Then
                   data_env.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                   data_env.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                Else
                   data_env.Recordset("cl_grupo") = 0
                   data_env.Recordset("cl_zona") = ""
                End If
                data_env.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                   data_env.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                Else
                   data_env.Recordset("cl_telefon") = ""
                End If
                If IsNull(data_cli.Recordset("cl_dircobr")) = False Then
                   data_env.Recordset("cl_dircobr") = data_cli.Recordset("cl_dircobr")
                Else
                   data_env.Recordset("cl_dircobr") = ""
                End If
                If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
                   data_env.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                Else
                   data_env.Recordset("cl_socmnom") = ""
                End If
                If IsNull(data_cli.Recordset("cl_socmnro")) = False Then
                   data_env.Recordset("cl_socmnro") = data_cli.Recordset("cl_socmnro")
                Else
                   data_env.Recordset("cl_socmnro") = ""
                End If
                If IsNull(data_cli.Recordset("cl_nrosocm")) = False Then
                   data_env.Recordset("cl_nrosocm") = data_cli.Recordset("cl_nrosocm")
                Else
                   data_env.Recordset("cl_nrosocm") = ""
                End If
                If IsNull(data_cli.Recordset("cl_fecing")) = False Then
                   data_env.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                End If
                If IsNull(data_cli.Recordset("fecha_baja")) = False Then
                   data_env.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                End If
                If IsNull(data_cli.Recordset("cl_nrovend")) = False Then
                   data_env.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                   data_env.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                Else
                   data_env.Recordset("cl_nrovend") = 799
                   data_env.Recordset("cl_nomvend") = "*TODOS"
                End If
                If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
                   data_env.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                   data_env.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                Else
                   data_env.Recordset("cl_nrocobr") = 0
                   data_env.Recordset("cl_nomcobr") = "*TODOS"
                End If
                If IsNull(data_cli.Recordset("cl_forpago")) = False Then
                   data_env.Recordset("cl_forpago") = data_cli.Recordset("cl_forpago")
                   data_env.Recordset("cl_descpag") = data_cli.Recordset("cl_descpag")
                Else
                   data_env.Recordset("cl_forpago") = 1
                   data_env.Recordset("cl_descpag") = "Abono Mensual"
                End If
                If IsNull(data_cli.Recordset("cl_diacobr")) = False Then
                   data_env.Recordset("cl_diacobr") = data_cli.Recordset("cl_diacobr")
                Else
                   data_env.Recordset("cl_diacobr") = ""
                End If
                If IsNull(data_cli.Recordset("tit_tarj")) = False Then
                   data_env.Recordset("tit_tarj") = data_cli.Recordset("tit_tarj")
                Else
                   data_env.Recordset("tit_tarj") = ""
                End If
                If IsNull(data_cli.Recordset("cl_nrotarj")) = False Then
                   data_env.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
                Else
                   data_env.Recordset("cl_nrotarj") = 0
                End If
                If IsNull(data_cli.Recordset("ci_tarj")) = False Then
                   data_env.Recordset("ci_tarj") = data_cli.Recordset("ci_tarj")
                Else
                   data_env.Recordset("ci_tarj") = 0
                End If
                If IsNull(data_cli.Recordset("codcitarj")) = False Then
                   data_env.Recordset("codcitarj") = data_cli.Recordset("codcitarj")
                Else
                   data_env.Recordset("codcitarj") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tjemi_c")) = False Then
                   data_env.Recordset("cl_tjemi_c") = data_cli.Recordset("cl_tjemi_c")
                Else
                   data_env.Recordset("cl_tjemi_c") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tjemi_n")) = False Then
                   data_env.Recordset("cl_tjemi_n") = data_cli.Recordset("cl_tjemi_n")
                Else
                   data_env.Recordset("cl_tjemi_n") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tj_venc")) = False Then
                   data_env.Recordset("cl_tj_venc") = data_cli.Recordset("cl_tj_venc")
                End If
                If IsNull(data_cli.Recordset("fecha_sys")) = False Then
                   data_env.Recordset("fecha_sys") = data_cli.Recordset("fecha_sys")
                End If
                If IsNull(data_cli.Recordset("fecha_modi")) = False Then
                   data_env.Recordset("fecha_modi") = data_cli.Recordset("fecha_modi")
                End If
                data_env.Recordset.Update
                data_cli.Recordset.MoveNext
           Loop
        End If
        data_env.RecordSource = "env_clib"
        data_env.Refresh
        If data_env.Recordset.RecordCount > 0 Then
           data_env.Recordset.MoveFirst
           Do While Not data_env.Recordset.EOF
              data_env.Recordset.Delete
              data_env.Recordset.MoveNext
           Loop
        End If
        
        data_cli.RecordSource = "Select * from clientes where fecha_baja =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
        data_cli.Refresh
        If data_cli.Recordset.RecordCount > 0 Then
           data_cli.Recordset.MoveFirst
           Do While Not data_cli.Recordset.EOF
                data_env.Recordset.AddNew
                data_env.Recordset("estado") = data_cli.Recordset("estado")
                data_env.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                data_env.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                data_env.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                data_env.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                If IsNull(data_cli.Recordset("cl_cedula")) = False Then
                   data_env.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                Else
                   data_env.Recordset("cl_cedula") = 0
                End If
                If IsNull(data_cli.Recordset("cl_codced")) = False Then
                   data_env.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                Else
                   data_env.Recordset("cl_codced") = 0
                End If
                If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                   data_env.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                End If
                If IsNull(data_cli.Recordset("cl_edad")) = False Then
                   data_env.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                Else
                   data_env.Recordset("cl_edad") = 0
                End If
                If IsNull(data_cli.Recordset("cl_uniedad")) = False Then
                   data_env.Recordset("cl_uniedad") = data_cli.Recordset("cl_uniedad")
                Else
                   data_env.Recordset("cl_uniedad") = "A"
                End If
                If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
                   data_env.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                Else
                   data_env.Recordset("cl_ultmesp") = 0
                End If
                If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
                   data_env.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                Else
                   data_env.Recordset("cl_ultanop") = 0
                End If
                If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
                   data_env.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa")
                Else
                   data_env.Recordset("cl_atrasoa") = 0
                End If
                If IsNull(data_cli.Recordset("saldo_cc")) = False Then
                   data_env.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                Else
                   data_env.Recordset("saldo_cc") = 0
                End If
                If IsNull(data_cli.Recordset("cl_direcci")) = False Then
                   data_env.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                Else
                   data_env.Recordset("cl_direcci") = ""
                End If
                If IsNull(data_cli.Recordset("cl_entre")) = False Then
                   data_env.Recordset("cl_entre") = data_cli.Recordset("cl_entre")
                Else
                   data_env.Recordset("cl_entre") = ""
                End If
                If IsNull(data_cli.Recordset("cl_grupo")) = False Then
                   data_env.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                   data_env.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                Else
                   data_env.Recordset("cl_grupo") = 0
                   data_env.Recordset("cl_zona") = ""
                End If
                data_env.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                   data_env.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                Else
                   data_env.Recordset("cl_telefon") = ""
                End If
                If IsNull(data_cli.Recordset("cl_dircobr")) = False Then
                   data_env.Recordset("cl_dircobr") = data_cli.Recordset("cl_dircobr")
                Else
                   data_env.Recordset("cl_dircobr") = ""
                End If
                If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
                   data_env.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                Else
                   data_env.Recordset("cl_socmnom") = ""
                End If
                If IsNull(data_cli.Recordset("cl_socmnro")) = False Then
                   data_env.Recordset("cl_socmnro") = data_cli.Recordset("cl_socmnro")
                Else
                   data_env.Recordset("cl_socmnro") = ""
                End If
                If IsNull(data_cli.Recordset("cl_nrosocm")) = False Then
                   data_env.Recordset("cl_nrosocm") = data_cli.Recordset("cl_nrosocm")
                Else
                   data_env.Recordset("cl_nrosocm") = ""
                End If
                If IsNull(data_cli.Recordset("cl_fecing")) = False Then
                   data_env.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                End If
                If IsNull(data_cli.Recordset("fecha_baja")) = False Then
                   data_env.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                End If
                If IsNull(data_cli.Recordset("cl_nrovend")) = False Then
                   data_env.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                   data_env.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                Else
                   data_env.Recordset("cl_nrovend") = 799
                   data_env.Recordset("cl_nomvend") = "*TODOS"
                End If
                If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
                   data_env.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                   data_env.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                Else
                   data_env.Recordset("cl_nrocobr") = 0
                   data_env.Recordset("cl_nomcobr") = "*TODOS"
                End If
                If IsNull(data_cli.Recordset("cl_forpago")) = False Then
                   data_env.Recordset("cl_forpago") = data_cli.Recordset("cl_forpago")
                   data_env.Recordset("cl_descpag") = data_cli.Recordset("cl_descpag")
                Else
                   data_env.Recordset("cl_forpago") = 1
                   data_env.Recordset("cl_descpag") = "Abono Mensual"
                End If
                If IsNull(data_cli.Recordset("cl_diacobr")) = False Then
                   data_env.Recordset("cl_diacobr") = data_cli.Recordset("cl_diacobr")
                Else
                   data_env.Recordset("cl_diacobr") = ""
                End If
                If IsNull(data_cli.Recordset("tit_tarj")) = False Then
                   data_env.Recordset("tit_tarj") = data_cli.Recordset("tit_tarj")
                Else
                   data_env.Recordset("tit_tarj") = ""
                End If
                If IsNull(data_cli.Recordset("cl_nrotarj")) = False Then
                   data_env.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
                Else
                   data_env.Recordset("cl_nrotarj") = 0
                End If
                If IsNull(data_cli.Recordset("ci_tarj")) = False Then
                   data_env.Recordset("ci_tarj") = data_cli.Recordset("ci_tarj")
                Else
                   data_env.Recordset("ci_tarj") = 0
                End If
                If IsNull(data_cli.Recordset("codcitarj")) = False Then
                   data_env.Recordset("codcitarj") = data_cli.Recordset("codcitarj")
                Else
                   data_env.Recordset("codcitarj") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tjemi_c")) = False Then
                   data_env.Recordset("cl_tjemi_c") = data_cli.Recordset("cl_tjemi_c")
                Else
                   data_env.Recordset("cl_tjemi_c") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tjemi_n")) = False Then
                   data_env.Recordset("cl_tjemi_n") = data_cli.Recordset("cl_tjemi_n")
                Else
                   data_env.Recordset("cl_tjemi_n") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tj_venc")) = False Then
                   data_env.Recordset("cl_tj_venc") = data_cli.Recordset("cl_tj_venc")
                End If
                If IsNull(data_cli.Recordset("fecha_sys")) = False Then
                   data_env.Recordset("fecha_sys") = data_cli.Recordset("fecha_sys")
                End If
                If IsNull(data_cli.Recordset("fecha_modi")) = False Then
                   data_env.Recordset("fecha_modi") = data_cli.Recordset("fecha_modi")
                End If
                data_env.Recordset.Update
                data_cli.Recordset.MoveNext
           Loop
        End If
    ' Modificaciones
        data_env.RecordSource = "env_clim"
        data_env.Refresh
        If data_env.Recordset.RecordCount > 0 Then
           data_env.Recordset.MoveFirst
           Do While Not data_env.Recordset.EOF
              data_env.Recordset.Delete
              data_env.Recordset.MoveNext
           Loop
        End If
        
        data_cli.RecordSource = "Select * from clientes where fecha_modi =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
        data_cli.Refresh
        If data_cli.Recordset.RecordCount > 0 Then
           data_cli.Recordset.MoveFirst
           Do While Not data_cli.Recordset.EOF
                data_env.Recordset.AddNew
                data_env.Recordset("estado") = data_cli.Recordset("estado")
                data_env.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                data_env.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                data_env.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                data_env.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                If IsNull(data_cli.Recordset("cl_cedula")) = False Then
                   data_env.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                Else
                   data_env.Recordset("cl_cedula") = 0
                End If
                If IsNull(data_cli.Recordset("cl_codced")) = False Then
                   data_env.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                Else
                   data_env.Recordset("cl_codced") = 0
                End If
                If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                   data_env.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                End If
                If IsNull(data_cli.Recordset("cl_edad")) = False Then
                   data_env.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                Else
                   data_env.Recordset("cl_edad") = 0
                End If
                If IsNull(data_cli.Recordset("cl_uniedad")) = False Then
                   data_env.Recordset("cl_uniedad") = data_cli.Recordset("cl_uniedad")
                Else
                   data_env.Recordset("cl_uniedad") = "A"
                End If
                If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
                   data_env.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                Else
                   data_env.Recordset("cl_ultmesp") = 0
                End If
                If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
                   data_env.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                Else
                   data_env.Recordset("cl_ultanop") = 0
                End If
                If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
                   data_env.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa")
                Else
                   data_env.Recordset("cl_atrasoa") = 0
                End If
                If IsNull(data_cli.Recordset("saldo_cc")) = False Then
                   data_env.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                Else
                   data_env.Recordset("saldo_cc") = 0
                End If
                If IsNull(data_cli.Recordset("cl_direcci")) = False Then
                   data_env.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                Else
                   data_env.Recordset("cl_direcci") = ""
                End If
                If IsNull(data_cli.Recordset("cl_entre")) = False Then
                   data_env.Recordset("cl_entre") = data_cli.Recordset("cl_entre")
                Else
                   data_env.Recordset("cl_entre") = ""
                End If
                If IsNull(data_cli.Recordset("cl_grupo")) = False Then
                   data_env.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                   data_env.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                Else
                   data_env.Recordset("cl_grupo") = 0
                   data_env.Recordset("cl_zona") = ""
                End If
                data_env.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                   data_env.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                Else
                   data_env.Recordset("cl_telefon") = ""
                End If
                If IsNull(data_cli.Recordset("cl_dircobr")) = False Then
                   data_env.Recordset("cl_dircobr") = data_cli.Recordset("cl_dircobr")
                Else
                   data_env.Recordset("cl_dircobr") = ""
                End If
                If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
                   data_env.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                Else
                   data_env.Recordset("cl_socmnom") = ""
                End If
                If IsNull(data_cli.Recordset("cl_socmnro")) = False Then
                   data_env.Recordset("cl_socmnro") = data_cli.Recordset("cl_socmnro")
                Else
                   data_env.Recordset("cl_socmnro") = ""
                End If
                If IsNull(data_cli.Recordset("cl_nrosocm")) = False Then
                   data_env.Recordset("cl_nrosocm") = data_cli.Recordset("cl_nrosocm")
                Else
                   data_env.Recordset("cl_nrosocm") = ""
                End If
                If IsNull(data_cli.Recordset("cl_fecing")) = False Then
                   data_env.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                End If
                If IsNull(data_cli.Recordset("fecha_baja")) = False Then
                   data_env.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                End If
                If IsNull(data_cli.Recordset("cl_nrovend")) = False Then
                   data_env.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                   data_env.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                Else
                   data_env.Recordset("cl_nrovend") = 799
                   data_env.Recordset("cl_nomvend") = "*TODOS"
                End If
                If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
                   data_env.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                   data_env.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                Else
                   data_env.Recordset("cl_nrocobr") = 0
                   data_env.Recordset("cl_nomcobr") = "*TODOS"
                End If
                If IsNull(data_cli.Recordset("cl_forpago")) = False Then
                   data_env.Recordset("cl_forpago") = data_cli.Recordset("cl_forpago")
                   data_env.Recordset("cl_descpag") = data_cli.Recordset("cl_descpag")
                Else
                   data_env.Recordset("cl_forpago") = 1
                   data_env.Recordset("cl_descpag") = "Abono Mensual"
                End If
                If IsNull(data_cli.Recordset("cl_diacobr")) = False Then
                   data_env.Recordset("cl_diacobr") = data_cli.Recordset("cl_diacobr")
                Else
                   data_env.Recordset("cl_diacobr") = ""
                End If
                If IsNull(data_cli.Recordset("tit_tarj")) = False Then
                   data_env.Recordset("tit_tarj") = data_cli.Recordset("tit_tarj")
                Else
                   data_env.Recordset("tit_tarj") = ""
                End If
                If IsNull(data_cli.Recordset("cl_nrotarj")) = False Then
                   data_env.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
                Else
                   data_env.Recordset("cl_nrotarj") = 0
                End If
                If IsNull(data_cli.Recordset("ci_tarj")) = False Then
                   data_env.Recordset("ci_tarj") = data_cli.Recordset("ci_tarj")
                Else
                   data_env.Recordset("ci_tarj") = 0
                End If
                If IsNull(data_cli.Recordset("codcitarj")) = False Then
                   data_env.Recordset("codcitarj") = data_cli.Recordset("codcitarj")
                Else
                   data_env.Recordset("codcitarj") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tjemi_c")) = False Then
                   data_env.Recordset("cl_tjemi_c") = data_cli.Recordset("cl_tjemi_c")
                Else
                   data_env.Recordset("cl_tjemi_c") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tjemi_n")) = False Then
                   data_env.Recordset("cl_tjemi_n") = data_cli.Recordset("cl_tjemi_n")
                Else
                   data_env.Recordset("cl_tjemi_n") = 0
                End If
                If IsNull(data_cli.Recordset("cl_tj_venc")) = False Then
                   data_env.Recordset("cl_tj_venc") = data_cli.Recordset("cl_tj_venc")
                End If
                If IsNull(data_cli.Recordset("fecha_sys")) = False Then
                   data_env.Recordset("fecha_sys") = data_cli.Recordset("fecha_sys")
                End If
                If IsNull(data_cli.Recordset("fecha_modi")) = False Then
                   data_env.Recordset("fecha_modi") = data_cli.Recordset("fecha_modi")
                End If
                data_env.Recordset.Update
                data_cli.Recordset.MoveNext
           Loop
        End If
        ' envío de archivo de historial
        data_env.RecordSource = "env_abm"
        data_env.Refresh
        If data_env.Recordset.RecordCount > 0 Then
           data_env.Recordset.MoveFirst
           Do While Not data_env.Recordset.EOF
              data_env.Recordset.Delete
              data_env.Recordset.MoveNext
           Loop
        End If
        
        data_cli.RecordSource = "Select * from abmsocio where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
        data_cli.Refresh
        If data_cli.Recordset.RecordCount > 0 Then
           data_cli.Recordset.MoveFirst
           Do While Not data_cli.Recordset.EOF
                data_env.Recordset.AddNew
                data_env.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                data_env.Recordset("cl_motivo") = data_cli.Recordset("cl_motivo")
                data_env.Recordset("desc") = data_cli.Recordset("desc")
                data_env.Recordset("fecha") = data_cli.Recordset("fecha")
                data_env.Recordset("hora") = data_cli.Recordset("hora")
                data_env.Recordset("usuario") = data_cli.Recordset("usuario")
                data_env.Recordset("convenio") = data_cli.Recordset("convenio")
                data_env.Recordset("base") = data_cli.Recordset("base")
                data_env.Recordset.Update
                data_cli.Recordset.MoveNext
           Loop
        End If
        data_env.RecordSource = "env_caja"
        data_env.Refresh
        If data_env.Recordset.RecordCount > 0 Then
           data_env.Recordset.MoveFirst
           Do While Not data_env.Recordset.EOF
              data_env.Recordset.Delete
              data_env.Recordset.MoveNext
           Loop
        End If
        If data_parsec.Recordset("base") = 2 Then
           data_cli.RecordSource = "Select * from caja where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "# And base =" & 3
           data_cli.Refresh
        Else
           If data_parsec.Recordset("base") = 3 Then
              data_cli.RecordSource = "Select * from caja where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
              data_cli.Refresh
           Else
              data_cli.RecordSource = "Select * from caja where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "# And base =" & data_parsec.Recordset("base")
              data_cli.Refresh
           End If
        End If
        If data_cli.Recordset.RecordCount > 0 Then
           data_cli.Recordset.MoveFirst
           Do While Not data_cli.Recordset.EOF
               data_env.Recordset.AddNew
                data_env.Recordset("fecha") = data_cli.Recordset("fecha")
                data_env.Recordset("moneda") = data_cli.Recordset("moneda")
                data_env.Recordset("numero") = data_cli.Recordset("numero")
                data_env.Recordset("nombre") = data_cli.Recordset("nombre")
                data_env.Recordset("movimiento") = data_cli.Recordset("movimiento")
                data_env.Recordset("imp_fact") = data_cli.Recordset("imp_fact")
                data_env.Recordset("observ") = data_cli.Recordset("observ")
                data_env.Recordset("saldo") = data_cli.Recordset("saldo")
                data_env.Recordset("usuario") = data_cli.Recordset("usuario")
                data_env.Recordset("hora") = data_cli.Recordset("hora")
                data_env.Recordset("saldo_user") = data_cli.Recordset("saldo_user")
                data_env.Recordset("base") = data_cli.Recordset("base")
                data_env.Recordset("cod_serv") = data_cli.Recordset("cod_serv")
                data_env.Recordset("nom_serv") = data_cli.Recordset("nom_serv")
                data_env.Recordset("cod_socio") = data_cli.Recordset("cod_socio")
                data_env.Recordset("nom_socio") = data_cli.Recordset("nom_socio")
                data_env.Recordset("documento") = data_cli.Recordset("documento")
                data_env.Recordset("caja_mesp") = data_cli.Recordset("caja_mesp")
                data_env.Recordset("caja_anop") = data_cli.Recordset("caja_anop")
                data_env.Recordset("imp_iva") = data_cli.Recordset("imp_iva")
                data_env.Recordset("opiva") = data_cli.Recordset("opiva")
                data_env.Recordset.Update
                data_cli.Recordset.MoveNext
           Loop
        End If
        data_env.RecordSource = "env_lin"
        data_env.Refresh
        If data_env.Recordset.RecordCount > 0 Then
           data_env.Recordset.MoveFirst
           Do While Not data_env.Recordset.EOF
              data_env.Recordset.Delete
              data_env.Recordset.MoveNext
           Loop
        End If
        If data_parsec.Recordset("base") = 2 Then
           data_cli.RecordSource = "Select * from linmmdd where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "# And base =" & 3
           data_cli.Refresh
        Else
           If data_parsec.Recordset("base") = 3 Then
              data_cli.RecordSource = "Select * from linmmdd where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
              data_cli.Refresh
           Else
              data_cli.RecordSource = "Select * from linmmdd where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "# And base =" & data_parsec.Recordset("base")
              data_cli.Refresh
           End If
        End If
        If data_cli.Recordset.RecordCount > 0 Then
           data_cli.Recordset.MoveFirst
           Do While Not data_cli.Recordset.EOF
                data_env.Recordset.AddNew
                data_env.Recordset("tipo_mov") = data_cli.Recordset("tipo_mov")
                data_env.Recordset("factura") = data_cli.Recordset("factura")
                data_env.Recordset("tipo") = data_cli.Recordset("tipo")
                data_env.Recordset("realizada") = data_cli.Recordset("realizada")
                data_env.Recordset("fecha") = data_cli.Recordset("fecha")
                data_env.Recordset("cod_cli") = data_cli.Recordset("cod_cli")
                data_env.Recordset("nom_cli") = data_cli.Recordset("nom_cli")
                data_env.Recordset("cod_prod") = data_cli.Recordset("cod_prod")
                data_env.Recordset("nom_prod") = data_cli.Recordset("nom_prod")
                data_env.Recordset("cantidad") = data_cli.Recordset("cantidad")
                data_env.Recordset("moneda") = data_cli.Recordset("moneda")
                data_env.Recordset("operador") = data_cli.Recordset("operador")
                data_env.Recordset("hora") = data_cli.Recordset("hora")
                data_env.Recordset("nro_flia") = data_cli.Recordset("nro_flia")
                data_env.Recordset("nom_flia") = data_cli.Recordset("nom_flia")
                data_env.Recordset("linea") = data_cli.Recordset("linea")
                data_env.Recordset("convenio") = data_cli.Recordset("convenio")
                data_env.Recordset("rub_cont") = data_cli.Recordset("rub_cont")
                data_env.Recordset("usa_timbre") = data_cli.Recordset("usa_timbre")
                data_env.Recordset("imp_timbre") = data_cli.Recordset("imp_timbre")
                data_env.Recordset("tot_lin") = data_cli.Recordset("tot_lin")
                data_env.Recordset("rub_nomb") = data_cli.Recordset("rub_nomb")
                data_env.Recordset("nro_med_a") = data_cli.Recordset("nro_med_a")
                data_env.Recordset("nom_med_a") = data_cli.Recordset("nom_med_a")
                data_env.Recordset("precio_est") = data_cli.Recordset("precio_est")
                data_env.Recordset("mes_paga") = data_cli.Recordset("mes_paga")
                data_env.Recordset("ano_paga") = data_cli.Recordset("ano_paga")
                data_env.Recordset("base") = data_cli.Recordset("base")
                data_env.Recordset("imp_iva") = data_cli.Recordset("imp_iva")
                data_env.Recordset("ruc") = data_cli.Recordset("ruc")
                data_env.Recordset.Update
                data_cli.Recordset.MoveNext
           Loop
        End If
        If data_parsec.Recordset("base") = 1 Then
           data_env.RecordSource = "env_lla"
           data_env.Refresh
           If data_env.Recordset.RecordCount > 0 Then
              data_env.Recordset.MoveFirst
              Do While Not data_env.Recordset.EOF
                 data_env.Recordset.Delete
                 data_env.Recordset.MoveNext
              Loop
           End If
           data_cli.RecordSource = "select * from llamado where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
           data_cli.Refresh
           If data_cli.Recordset.RecordCount > 0 Then
              data_cli.Recordset.MoveFirst
              Do While Not data_cli.Recordset.EOF
                 data_env.Recordset.AddNew
                 data_env.Recordset("nro") = data_cli.Recordset("nro")
                 data_env.Recordset("fecha") = data_cli.Recordset("fecha")
                 data_env.Recordset("hora") = data_cli.Recordset("hora")
                 data_env.Recordset("usuario") = data_cli.Recordset("usuario")
                 data_env.Recordset("matric") = data_cli.Recordset("matric")
                 data_env.Recordset("nombre") = data_cli.Recordset("nombre")
                 data_env.Recordset("edad") = data_cli.Recordset("edad")
                 data_env.Recordset("unied") = data_cli.Recordset("unied")
                 data_env.Recordset("categ") = data_cli.Recordset("categ")
                 data_env.Recordset("nomcat") = data_cli.Recordset("nomcat")
                 data_env.Recordset("ci") = data_cli.Recordset("ci")
                 data_env.Recordset("direcc") = data_cli.Recordset("direcc")
                 data_env.Recordset("telef") = data_cli.Recordset("telef")
                 data_env.Recordset("codzon") = data_cli.Recordset("codzon")
                 data_env.Recordset("base") = data_cli.Recordset("base")
                 data_env.Recordset("referen") = data_cli.Recordset("referen")
                 data_env.Recordset("motcon") = data_cli.Recordset("motcon")
                 data_env.Recordset("obsmot") = data_cli.Recordset("obsmot")
                 data_env.Recordset("codmot") = data_cli.Recordset("codmot")
                 data_env.Recordset("descol") = data_cli.Recordset("descol")
                 data_env.Recordset("movilpas") = data_cli.Recordset("movilpas")
                 data_env.Recordset("fec_rea") = data_cli.Recordset("fec_rea")
                 data_env.Recordset("pend") = data_cli.Recordset("pend")
                 data_env.Recordset("hor_rea") = data_cli.Recordset("hor_rea")
                 data_env.Recordset("fecpas") = data_cli.Recordset("fecpas")
                 data_env.Recordset("horpas") = data_cli.Recordset("horpas")
                 data_env.Recordset("fecsali") = data_cli.Recordset("fecsali")
                 data_env.Recordset("horsali") = data_cli.Recordset("horsali")
                 data_env.Recordset("fec_llega") = data_cli.Recordset("fec_llega")
                 data_env.Recordset("hor_llega") = data_cli.Recordset("hor_llega")
                 data_env.Recordset("fec_rea") = data_cli.Recordset("fec_rea")
                 data_env.Recordset("hor_rea") = data_cli.Recordset("hor_rea")
                 data_env.Recordset("diag") = data_cli.Recordset("diag")
                 data_env.Recordset("colormot") = data_cli.Recordset("colormot")
                 data_env.Recordset("codmed") = data_cli.Recordset("codmed")
                 data_env.Recordset("nommed") = data_cli.Recordset("nommed")
                 data_env.Recordset("trasla") = data_cli.Recordset("trasla")
                 data_env.Recordset("lugar") = data_cli.Recordset("lugar")
                 data_env.Recordset("hsald") = data_cli.Recordset("hsald")
                 data_env.Recordset("hllega") = data_cli.Recordset("hllega")
                 data_env.Recordset("hzona") = data_cli.Recordset("hzona")
                 data_env.Recordset("movil_rea") = data_cli.Recordset("movil_rea")
                 data_env.Recordset("totdem") = data_cli.Recordset("totdem")
                 data_env.Recordset("totend") = data_cli.Recordset("totend")
                 data_env.Recordset.Update
                 data_cli.Recordset.MoveNext
              Loop
           End If
        End If
    ' Tesorería
        data_env.DatabaseName = "c:\datos"
        data_env.RecordSource = "env_tes"
        data_env.Refresh
        If data_env.Recordset.RecordCount > 0 Then
           data_env.Recordset.MoveFirst
           Do While Not data_env.Recordset.EOF
              data_env.Recordset.Delete
              data_env.Recordset.MoveNext
           Loop
        End If
        If data_parsec.Recordset("base") = 15 Or _
           data_parsec.Recordset("base") = 6 Or _
           data_parsec.Recordset("base") = 3 Or _
           data_parsec.Recordset("base") = 1 Then
           data_cli.DatabaseName = App.Path & "\tesorero.mdb"
           data_cli.RecordSource = "Select * from tesorero where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
           data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_cli.Recordset.MoveFirst
               Do While Not data_cli.Recordset.EOF
                   data_env.Recordset.AddNew
                    data_env.Recordset("nromov") = data_cli.Recordset("nromov")
                    data_env.Recordset("fecha") = data_cli.Recordset("fecha")
                    data_env.Recordset("hora") = data_cli.Recordset("hora")
                    data_env.Recordset("usuario") = data_cli.Recordset("usuario")
                    data_env.Recordset("cod_rub") = data_cli.Recordset("cod_rub")
                    data_env.Recordset("nom_rub") = data_cli.Recordset("nom_rub")
                    data_env.Recordset("moneda") = data_cli.Recordset("moneda")
                    data_env.Recordset("monto") = data_cli.Recordset("monto")
                    data_env.Recordset("obs") = data_cli.Recordset("obs")
                    data_env.Recordset("cod_debe") = data_cli.Recordset("cod_debe")
                    data_env.Recordset("saldos") = data_cli.Recordset("saldos")
                    data_env.Recordset("concep") = data_cli.Recordset("concep")
                    data_env.Recordset("cod_haber") = data_cli.Recordset("cod_haber")
                    data_env.Recordset("saldou") = data_cli.Recordset("saldou")
                    data_env.Recordset("tipoc") = data_cli.Recordset("tipoc")
                    data_env.Recordset("libro") = data_cli.Recordset("libro")
                    data_env.Recordset("iva") = data_cli.Recordset("iva")
                    data_env.Recordset("base") = data_cli.Recordset("base")
                    data_env.Recordset("descon") = data_cli.Recordset("descon")
                    data_env.Recordset("bandera") = data_cli.Recordset("bandera")
                    data_env.Recordset("impiva") = data_cli.Recordset("impiva")
                    data_env.Recordset("tcam") = data_cli.Recordset("tcam")
                    data_env.Recordset.Update
                    data_cli.Recordset.MoveNext
               Loop
            End If
        End If
           
        data_env.DatabaseName = ""
        data_env.RecordSource = ""
        data_env.Refresh
        data_cli.DatabaseName = ""
        data_cli.RecordSource = ""
        data_cli.Refresh
        Timer6.Enabled = True
        Timer3.Enabled = False
           
    Else
    
    ' Arqueo
    ' Deudas
        data_env.DatabaseName = ""
        data_env.RecordSource = ""
        data_env.Refresh
        data_cli.DatabaseName = ""
        data_cli.RecordSource = ""
        data_cli.Refresh
        If Dir("C:\Datos\envios.zip") <> "" Then
           Kill "c:\datos\envios.zip"
        End If
        Timer3.Enabled = False
        Timer2.Enabled = True
    End If
Else
    Unload Me
End If

End Sub

Private Sub Timer4_Timer()

   data_rec.DatabaseName = "c:\datos\recibe"
   data_rec.RecordSource = "env_clia"
   data_rec.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveFirst
      data_cli.DatabaseName = App.Path & "\sapp.mdb"
      data_cli.RecordSource = "clientes"
      data_cli.Refresh
      Do While Not data_rec.Recordset.EOF
          data_cli.Recordset.FindFirst "cl_codigo =" & data_rec.Recordset("cl_codigo")
          If Not data_cli.Recordset.NoMatch Then
             data_rec.Recordset.MoveNext
          Else
             data_cli.Recordset.AddNew
              data_cli.Recordset("estado") = data_rec.Recordset("estado")
              data_cli.Recordset("cl_codigo") = data_rec.Recordset("cl_codigo")
              data_cli.Recordset("cl_codconv") = data_rec.Recordset("cl_codconv")
              data_cli.Recordset("cl_nomconv") = data_rec.Recordset("cl_nomconv")
              data_cli.Recordset("cl_apellid") = data_rec.Recordset("cl_apellid")
              If IsNull(data_rec.Recordset("cl_cedula")) = False Then
                 data_cli.Recordset("cl_cedula") = data_rec.Recordset("cl_cedula")
              Else
                 data_cli.Recordset("cl_cedula") = 0
              End If
              If IsNull(data_rec.Recordset("cl_codced")) = False Then
                 data_cli.Recordset("cl_codced") = data_rec.Recordset("cl_codced")
              Else
                 data_cli.Recordset("cl_codced") = 0
              End If
              If IsNull(data_rec.Recordset("cl_fnac")) = False Then
                 data_cli.Recordset("cl_fnac") = data_rec.Recordset("cl_fnac")
              End If
              If IsNull(data_rec.Recordset("cl_edad")) = False Then
                 data_cli.Recordset("cl_edad") = data_rec.Recordset("cl_edad")
              Else
                 data_cli.Recordset("cl_edad") = 0
              End If
              If IsNull(data_rec.Recordset("cl_uniedad")) = False Then
                 data_cli.Recordset("cl_uniedad") = data_rec.Recordset("cl_uniedad")
              Else
                 data_cli.Recordset("cl_uniedad") = "A"
              End If
              If IsNull(data_rec.Recordset("cl_ultmesp")) = False Then
                 data_cli.Recordset("cl_ultmesp") = data_rec.Recordset("cl_ultmesp")
              Else
                 data_cli.Recordset("cl_ultmesp") = 0
              End If
              If IsNull(data_rec.Recordset("cl_ultanop")) = False Then
                 data_cli.Recordset("cl_ultanop") = data_rec.Recordset("cl_ultanop")
              Else
                 data_cli.Recordset("cl_ultanop") = 0
              End If
              If IsNull(data_rec.Recordset("cl_atrasoa")) = False Then
                 data_cli.Recordset("cl_atrasoa") = data_rec.Recordset("cl_atrasoa")
              Else
                 data_cli.Recordset("cl_atrasoa") = 0
              End If
              If IsNull(data_rec.Recordset("saldo_cc")) = False Then
                 data_cli.Recordset("saldo_cc") = data_rec.Recordset("saldo_cc")
              Else
                 data_cli.Recordset("saldo_cc") = 0
              End If
              If IsNull(data_rec.Recordset("cl_direcci")) = False Then
                 data_cli.Recordset("cl_direcci") = data_rec.Recordset("cl_direcci")
              Else
                 data_cli.Recordset("cl_direcci") = ""
              End If
              If IsNull(data_rec.Recordset("cl_entre")) = False Then
                 data_cli.Recordset("cl_entre") = data_rec.Recordset("cl_entre")
              Else
                 data_cli.Recordset("cl_entre") = ""
              End If
              If IsNull(data_rec.Recordset("cl_grupo")) = False Then
                 data_cli.Recordset("cl_grupo") = data_rec.Recordset("cl_grupo")
                 data_cli.Recordset("cl_zona") = data_rec.Recordset("cl_zona")
              Else
                 data_cli.Recordset("cl_grupo") = ""
                 data_cli.Recordset("cl_zona") = ""
              End If
              data_cli.Recordset("cl_sexo") = data_rec.Recordset("cl_sexo")
              If IsNull(data_rec.Recordset("cl_telefon")) = False Then
                 data_cli.Recordset("cl_telefon") = data_rec.Recordset("cl_telefon")
              Else
                 data_cli.Recordset("cl_telefon") = ""
              End If
              If IsNull(data_rec.Recordset("cl_dircobr")) = False Then
                 data_cli.Recordset("cl_dircobr") = data_rec.Recordset("cl_dircobr")
              Else
                 data_cli.Recordset("cl_dircobr") = ""
              End If
              If IsNull(data_rec.Recordset("cl_socmnom")) = False Then
                 data_cli.Recordset("cl_socmnom") = data_rec.Recordset("cl_socmnom")
              Else
                 data_cli.Recordset("cl_socmnom") = ""
              End If
              If IsNull(data_rec.Recordset("cl_socmnro")) = False Then
                 data_cli.Recordset("cl_socmnro") = data_rec.Recordset("cl_socmnro")
              Else
                 data_cli.Recordset("cl_socmnro") = ""
              End If
              If IsNull(data_rec.Recordset("cl_fecing")) = False Then
                 data_cli.Recordset("cl_fecing") = data_rec.Recordset("cl_fecing")
              End If
              If IsNull(data_rec.Recordset("fecha_baja")) = False Then
                 data_cli.Recordset("fecha_baja") = data_rec.Recordset("fecha_baja")
              End If
              If IsNull(data_rec.Recordset("cl_nrovend")) = False Then
                 data_cli.Recordset("cl_nrovend") = data_rec.Recordset("cl_nrovend")
                 data_cli.Recordset("cl_nomvend") = data_rec.Recordset("cl_nomvend")
              Else
                 data_cli.Recordset("cl_nrovend") = 799
                 data_cli.Recordset("cl_nomvend") = "*TODOS"
              End If
              If IsNull(data_rec.Recordset("cl_nrocobr")) = False Then
                 data_cli.Recordset("cl_nrocobr") = data_rec.Recordset("cl_nrocobr")
                 data_cli.Recordset("cl_nomcobr") = data_rec.Recordset("cl_nomcobr")
              Else
                 data_cli.Recordset("cl_nrocobr") = 0
                 data_cli.Recordset("cl_nomcobr") = "*TODOS"
              End If
              If IsNull(data_rec.Recordset("cl_forpago")) = False Then
                 data_cli.Recordset("cl_forpago") = data_rec.Recordset("cl_forpago")
                 data_cli.Recordset("cl_descpag") = data_rec.Recordset("cl_descpag")
              Else
                 data_cli.Recordset("cl_forpago") = 1
                 data_cli.Recordset("cl_descpag") = "Abono Mensual"
              End If
              If IsNull(data_rec.Recordset("cl_diacobr")) = False Then
                 data_cli.Recordset("cl_diacobr") = data_rec.Recordset("cl_diacobr")
              Else
                 data_cli.Recordset("cl_diacobr") = ""
              End If
              If IsNull(data_rec.Recordset("tit_tarj")) = False Then
                 data_cli.Recordset("tit_tarj") = data_rec.Recordset("tit_tarj")
              Else
                 data_cli.Recordset("tit_tarj") = ""
              End If
              If IsNull(data_rec.Recordset("cl_nrotarj")) = False Then
                 data_cli.Recordset("cl_nrotarj") = data_rec.Recordset("cl_nrotarj")
              Else
                 data_cli.Recordset("cl_nrotarj") = 0
              End If
              If IsNull(data_rec.Recordset("ci_tarj")) = False Then
                 data_cli.Recordset("ci_tarj") = data_rec.Recordset("ci_tarj")
              Else
                 data_cli.Recordset("ci_tarj") = 0
              End If
              If IsNull(data_rec.Recordset("codcitarj")) = False Then
                 data_cli.Recordset("codcitarj") = data_rec.Recordset("codcitarj")
              Else
                 data_cli.Recordset("codcitarj") = 0
              End If
              If IsNull(data_rec.Recordset("cl_tjemi_c")) = False Then
                 data_cli.Recordset("cl_tjemi_c") = data_rec.Recordset("cl_tjemi_c")
              Else
                 data_cli.Recordset("cl_tjemi_c") = 0
              End If
              If IsNull(data_rec.Recordset("cl_tjemi_n")) = False Then
                 data_cli.Recordset("cl_tjemi_n") = data_rec.Recordset("cl_tjemi_n")
              Else
                 data_cli.Recordset("cl_tjemi_n") = 0
              End If
              If IsNull(data_rec.Recordset("cl_tj_venc")) = False Then
                 data_cli.Recordset("cl_tj_venc") = data_rec.Recordset("cl_tj_venc")
              End If
              If IsNull(data_rec.Recordset("fecha_sys")) = False Then
                 data_cli.Recordset("fecha_sys") = data_rec.Recordset("fecha_sys")
              End If
              If IsNull(data_rec.Recordset("fecha_modi")) = False Then
                 data_cli.Recordset("fecha_modi") = data_rec.Recordset("fecha_modi")
              End If
              data_cli.Recordset.Update
              data_rec.Recordset.MoveNext
          End If
      Loop
      data_rec.Recordset.MoveFirst
      Do While Not data_rec.Recordset.EOF
          data_rec.Recordset.Delete
          data_rec.Recordset.MoveNext
      Loop
   End If
   data_rec.RecordSource = "env_clib"
   data_rec.Refresh
   Dim XBanbajmod As Integer
   XBanbajmod = 0
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveFirst
      data_cli.DatabaseName = App.Path & "\sapp.mdb"
      data_cli.RecordSource = "clientes"
      data_cli.Refresh
      Do While Not data_rec.Recordset.EOF
          data_cli.Recordset.FindFirst "cl_codigo =" & data_rec.Recordset("cl_codigo")
          If Not data_cli.Recordset.NoMatch Then
             If IsNull(data_cli.Recordset("fecha_baja")) = False Then
                If data_cli.Recordset("fecha_baja") >= data_rec.Recordset("fecha_baja") Then
                   XBanbajmod = 2
                Else
                   XBanbajmod = 0
                End If
             Else
                XBanbajmod = 0
             End If
             If XBanbajmod <> 2 Then
                data_cli.Recordset.Edit
                data_cli.Recordset("estado") = data_rec.Recordset("estado")
                data_cli.Recordset("cl_codigo") = data_rec.Recordset("cl_codigo")
                data_cli.Recordset("cl_codconv") = data_rec.Recordset("cl_codconv")
                data_cli.Recordset("cl_nomconv") = data_rec.Recordset("cl_nomconv")
                data_cli.Recordset("cl_apellid") = data_rec.Recordset("cl_apellid")
                If IsNull(data_rec.Recordset("cl_cedula")) = False Then
                   data_cli.Recordset("cl_cedula") = data_rec.Recordset("cl_cedula")
                Else
                   data_cli.Recordset("cl_cedula") = 0
                End If
                If IsNull(data_rec.Recordset("cl_codced")) = False Then
                   data_cli.Recordset("cl_codced") = data_rec.Recordset("cl_codced")
                Else
                   data_cli.Recordset("cl_codced") = 0
                End If
                If IsNull(data_rec.Recordset("cl_fnac")) = False Then
                   data_cli.Recordset("cl_fnac") = data_rec.Recordset("cl_fnac")
                End If
                If IsNull(data_rec.Recordset("cl_edad")) = False Then
                   data_cli.Recordset("cl_edad") = data_rec.Recordset("cl_edad")
                Else
                   data_cli.Recordset("cl_edad") = 0
                End If
                If IsNull(data_rec.Recordset("cl_uniedad")) = False Then
                   data_cli.Recordset("cl_uniedad") = data_rec.Recordset("cl_uniedad")
                Else
                   data_cli.Recordset("cl_uniedad") = "A"
                End If
                If IsNull(data_rec.Recordset("cl_ultmesp")) = False Then
                   data_cli.Recordset("cl_ultmesp") = data_rec.Recordset("cl_ultmesp")
                Else
                   data_cli.Recordset("cl_ultmesp") = 0
                End If
                If IsNull(data_rec.Recordset("cl_ultanop")) = False Then
                   data_cli.Recordset("cl_ultanop") = data_rec.Recordset("cl_ultanop")
                Else
                   data_cli.Recordset("cl_ultanop") = 0
                End If
                If IsNull(data_rec.Recordset("cl_atrasoa")) = False Then
                   data_cli.Recordset("cl_atrasoa") = data_rec.Recordset("cl_atrasoa")
                Else
                   data_cli.Recordset("cl_atrasoa") = 0
                End If
                If IsNull(data_rec.Recordset("saldo_cc")) = False Then
                   data_cli.Recordset("saldo_cc") = data_rec.Recordset("saldo_cc")
                Else
                   data_cli.Recordset("saldo_cc") = 0
                End If
                If IsNull(data_rec.Recordset("cl_direcci")) = False Then
                   data_cli.Recordset("cl_direcci") = data_rec.Recordset("cl_direcci")
                Else
                   data_cli.Recordset("cl_direcci") = ""
                End If
                If IsNull(data_rec.Recordset("cl_entre")) = False Then
                   data_cli.Recordset("cl_entre") = data_rec.Recordset("cl_entre")
                Else
                   data_cli.Recordset("cl_entre") = ""
                End If
                If IsNull(data_rec.Recordset("cl_grupo")) = False Then
                   data_cli.Recordset("cl_grupo") = data_rec.Recordset("cl_grupo")
                   data_cli.Recordset("cl_zona") = data_rec.Recordset("cl_zona")
                Else
                   data_cli.Recordset("cl_grupo") = ""
                   data_cli.Recordset("cl_zona") = ""
                End If
                data_cli.Recordset("cl_sexo") = data_rec.Recordset("cl_sexo")
                If IsNull(data_rec.Recordset("cl_telefon")) = False Then
                   data_cli.Recordset("cl_telefon") = data_rec.Recordset("cl_telefon")
                Else
                   data_cli.Recordset("cl_telefon") = ""
                End If
                If IsNull(data_rec.Recordset("cl_dircobr")) = False Then
                   data_cli.Recordset("cl_dircobr") = data_rec.Recordset("cl_dircobr")
                Else
                   data_cli.Recordset("cl_dircobr") = ""
                End If
                If IsNull(data_rec.Recordset("cl_socmnom")) = False Then
                   data_cli.Recordset("cl_socmnom") = data_rec.Recordset("cl_socmnom")
                Else
                   data_cli.Recordset("cl_socmnom") = ""
                End If
                If IsNull(data_rec.Recordset("cl_socmnro")) = False Then
                   data_cli.Recordset("cl_socmnro") = data_rec.Recordset("cl_socmnro")
                Else
                   data_cli.Recordset("cl_socmnro") = ""
                End If
                If IsNull(data_rec.Recordset("cl_fecing")) = False Then
                   data_cli.Recordset("cl_fecing") = data_rec.Recordset("cl_fecing")
                End If
                If IsNull(data_rec.Recordset("fecha_baja")) = False Then
                   data_cli.Recordset("fecha_baja") = data_rec.Recordset("fecha_baja")
                Else
                   data_cli.Recordset("fecha_baja") = data_rec.Recordset("fecha_baja")
                End If
                If IsNull(data_rec.Recordset("cl_nrovend")) = False Then
                   data_cli.Recordset("cl_nrovend") = data_rec.Recordset("cl_nrovend")
                   data_cli.Recordset("cl_nomvend") = data_rec.Recordset("cl_nomvend")
                Else
                   data_cli.Recordset("cl_nrovend") = 799
                   data_cli.Recordset("cl_nomvend") = "*TODOS"
                End If
                If IsNull(data_rec.Recordset("cl_nrocobr")) = False Then
                   data_cli.Recordset("cl_nrocobr") = data_rec.Recordset("cl_nrocobr")
                   data_cli.Recordset("cl_nomcobr") = data_rec.Recordset("cl_nomcobr")
                Else
                   data_cli.Recordset("cl_nrocobr") = 0
                   data_cli.Recordset("cl_nomcobr") = "*TODOS"
                End If
                If IsNull(data_rec.Recordset("cl_forpago")) = False Then
                   data_cli.Recordset("cl_forpago") = data_rec.Recordset("cl_forpago")
                   data_cli.Recordset("cl_descpag") = data_rec.Recordset("cl_descpag")
                Else
                   data_cli.Recordset("cl_forpago") = 1
                   data_cli.Recordset("cl_descpag") = "Abono Mensual"
                End If
                If IsNull(data_rec.Recordset("cl_diacobr")) = False Then
                   data_cli.Recordset("cl_diacobr") = data_rec.Recordset("cl_diacobr")
                Else
                   data_cli.Recordset("cl_diacobr") = ""
                End If
                If IsNull(data_rec.Recordset("tit_tarj")) = False Then
                   data_cli.Recordset("tit_tarj") = data_rec.Recordset("tit_tarj")
                Else
                   data_cli.Recordset("tit_tarj") = ""
                End If
                If IsNull(data_rec.Recordset("cl_nrotarj")) = False Then
                   data_cli.Recordset("cl_nrotarj") = data_rec.Recordset("cl_nrotarj")
                Else
                   data_cli.Recordset("cl_nrotarj") = 0
                End If
                If IsNull(data_rec.Recordset("ci_tarj")) = False Then
                   data_cli.Recordset("ci_tarj") = data_rec.Recordset("ci_tarj")
                Else
                   data_cli.Recordset("ci_tarj") = 0
                End If
                If IsNull(data_rec.Recordset("codcitarj")) = False Then
                   data_cli.Recordset("codcitarj") = data_rec.Recordset("codcitarj")
                Else
                   data_cli.Recordset("codcitarj") = 0
                End If
                If IsNull(data_rec.Recordset("cl_tjemi_c")) = False Then
                   data_cli.Recordset("cl_tjemi_c") = data_rec.Recordset("cl_tjemi_c")
                Else
                   data_cli.Recordset("cl_tjemi_c") = 0
                End If
                If IsNull(data_rec.Recordset("cl_tjemi_n")) = False Then
                   data_cli.Recordset("cl_tjemi_n") = data_rec.Recordset("cl_tjemi_n")
                Else
                   data_cli.Recordset("cl_tjemi_n") = 0
                End If
                If IsNull(data_rec.Recordset("cl_tj_venc")) = False Then
                   data_cli.Recordset("cl_tj_venc") = data_rec.Recordset("cl_tj_venc")
                End If
                If IsNull(data_rec.Recordset("fecha_sys")) = False Then
                   data_cli.Recordset("fecha_sys") = data_rec.Recordset("fecha_sys")
                End If
                If IsNull(data_rec.Recordset("fecha_modi")) = False Then
                   data_cli.Recordset("fecha_modi") = data_rec.Recordset("fecha_modi")
                End If
                data_cli.Recordset.Update
             End If
             data_rec.Recordset.MoveNext
          Else
              data_rec.Recordset.MoveNext
          End If
      Loop
      data_rec.Recordset.MoveFirst
      Do While Not data_rec.Recordset.EOF
          data_rec.Recordset.Delete
          data_rec.Recordset.MoveNext
      Loop
   End If
   XBanbajmod = 0
   data_rec.RecordSource = "env_clim"
   data_rec.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveFirst
      data_cli.DatabaseName = App.Path & "\sapp.mdb"
      data_cli.RecordSource = "clientes"
      data_cli.Refresh
      Do While Not data_rec.Recordset.EOF
          data_cli.Recordset.FindFirst "cl_codigo =" & data_rec.Recordset("cl_codigo")
          If Not data_cli.Recordset.NoMatch Then
             If IsNull(data_cli.Recordset("fecha_modi")) = False Then
                If data_cli.Recordset("fecha_modi") >= data_rec.Recordset("fecha_modi") Then
                   XBanbajmod = 2
                Else
                   XBanbajmod = 0
                End If
             Else
                XBanbajmod = 0
             End If
             If XBanbajmod <> 2 Then
                data_cli.Recordset.Edit
                data_cli.Recordset("estado") = data_rec.Recordset("estado")
                data_cli.Recordset("cl_codigo") = data_rec.Recordset("cl_codigo")
                data_cli.Recordset("cl_codconv") = data_rec.Recordset("cl_codconv")
                data_cli.Recordset("cl_nomconv") = data_rec.Recordset("cl_nomconv")
                data_cli.Recordset("cl_apellid") = data_rec.Recordset("cl_apellid")
                If IsNull(data_rec.Recordset("cl_cedula")) = False Then
                   data_cli.Recordset("cl_cedula") = data_rec.Recordset("cl_cedula")
                Else
                   data_cli.Recordset("cl_cedula") = 0
                End If
                If IsNull(data_rec.Recordset("cl_codced")) = False Then
                   data_cli.Recordset("cl_codced") = data_rec.Recordset("cl_codced")
                Else
                   data_cli.Recordset("cl_codced") = 0
                End If
                If IsNull(data_rec.Recordset("cl_fnac")) = False Then
                   data_cli.Recordset("cl_fnac") = data_rec.Recordset("cl_fnac")
                End If
                If IsNull(data_rec.Recordset("cl_edad")) = False Then
                   data_cli.Recordset("cl_edad") = data_rec.Recordset("cl_edad")
                Else
                   data_cli.Recordset("cl_edad") = 0
                End If
                If IsNull(data_rec.Recordset("cl_uniedad")) = False Then
                   data_cli.Recordset("cl_uniedad") = data_rec.Recordset("cl_uniedad")
                Else
                   data_cli.Recordset("cl_uniedad") = "A"
                End If
                If IsNull(data_rec.Recordset("cl_ultmesp")) = False Then
                   data_cli.Recordset("cl_ultmesp") = data_rec.Recordset("cl_ultmesp")
                Else
                   data_cli.Recordset("cl_ultmesp") = 0
                End If
                If IsNull(data_rec.Recordset("cl_ultanop")) = False Then
                   data_cli.Recordset("cl_ultanop") = data_rec.Recordset("cl_ultanop")
                Else
                   data_cli.Recordset("cl_ultanop") = 0
                End If
                If IsNull(data_rec.Recordset("cl_atrasoa")) = False Then
                   data_cli.Recordset("cl_atrasoa") = data_rec.Recordset("cl_atrasoa")
                Else
                   data_cli.Recordset("cl_atrasoa") = 0
                End If
                If IsNull(data_rec.Recordset("saldo_cc")) = False Then
                   data_cli.Recordset("saldo_cc") = data_rec.Recordset("saldo_cc")
                Else
                   data_cli.Recordset("saldo_cc") = 0
                End If
                If IsNull(data_rec.Recordset("cl_direcci")) = False Then
                   data_cli.Recordset("cl_direcci") = data_rec.Recordset("cl_direcci")
                Else
                   data_cli.Recordset("cl_direcci") = ""
                End If
                If IsNull(data_rec.Recordset("cl_entre")) = False Then
                   data_cli.Recordset("cl_entre") = data_rec.Recordset("cl_entre")
                Else
                   data_cli.Recordset("cl_entre") = ""
                End If
                If IsNull(data_rec.Recordset("cl_grupo")) = False Then
                   data_cli.Recordset("cl_grupo") = data_rec.Recordset("cl_grupo")
                   data_cli.Recordset("cl_zona") = data_rec.Recordset("cl_zona")
                Else
                   data_cli.Recordset("cl_grupo") = ""
                   data_cli.Recordset("cl_zona") = ""
                End If
                data_cli.Recordset("cl_sexo") = data_rec.Recordset("cl_sexo")
                If IsNull(data_rec.Recordset("cl_telefon")) = False Then
                   data_cli.Recordset("cl_telefon") = data_rec.Recordset("cl_telefon")
                Else
                   data_cli.Recordset("cl_telefon") = ""
                End If
                If IsNull(data_rec.Recordset("cl_dircobr")) = False Then
                   data_cli.Recordset("cl_dircobr") = data_rec.Recordset("cl_dircobr")
                Else
                   data_cli.Recordset("cl_dircobr") = ""
                End If
                If IsNull(data_rec.Recordset("cl_socmnom")) = False Then
                   data_cli.Recordset("cl_socmnom") = data_rec.Recordset("cl_socmnom")
                Else
                   data_cli.Recordset("cl_socmnom") = ""
                End If
                If IsNull(data_rec.Recordset("cl_socmnro")) = False Then
                   data_cli.Recordset("cl_socmnro") = data_rec.Recordset("cl_socmnro")
                Else
                   data_cli.Recordset("cl_socmnro") = ""
                End If
                If IsNull(data_rec.Recordset("cl_fecing")) = False Then
                   data_cli.Recordset("cl_fecing") = data_rec.Recordset("cl_fecing")
                End If
                If IsNull(data_rec.Recordset("fecha_baja")) = False Then
                   data_cli.Recordset("fecha_baja") = data_rec.Recordset("fecha_baja")
                Else
                   data_cli.Recordset("fecha_baja") = data_rec.Recordset("fecha_baja")
                End If
                If IsNull(data_rec.Recordset("cl_nrovend")) = False Then
                   data_cli.Recordset("cl_nrovend") = data_rec.Recordset("cl_nrovend")
                   data_cli.Recordset("cl_nomvend") = data_rec.Recordset("cl_nomvend")
                Else
                   data_cli.Recordset("cl_nrovend") = 799
                   data_cli.Recordset("cl_nomvend") = "*TODOS"
                End If
                If IsNull(data_rec.Recordset("cl_nrocobr")) = False Then
                   data_cli.Recordset("cl_nrocobr") = data_rec.Recordset("cl_nrocobr")
                   data_cli.Recordset("cl_nomcobr") = data_rec.Recordset("cl_nomcobr")
                Else
                   data_cli.Recordset("cl_nrocobr") = 0
                   data_cli.Recordset("cl_nomcobr") = "*TODOS"
                End If
                If IsNull(data_rec.Recordset("cl_forpago")) = False Then
                   data_cli.Recordset("cl_forpago") = data_rec.Recordset("cl_forpago")
                   data_cli.Recordset("cl_descpag") = data_rec.Recordset("cl_descpag")
                Else
                   data_cli.Recordset("cl_forpago") = 1
                   data_cli.Recordset("cl_descpag") = "Abono Mensual"
                End If
                If IsNull(data_rec.Recordset("cl_diacobr")) = False Then
                   data_cli.Recordset("cl_diacobr") = data_rec.Recordset("cl_diacobr")
                Else
                   data_cli.Recordset("cl_diacobr") = ""
                End If
                If IsNull(data_rec.Recordset("tit_tarj")) = False Then
                   data_cli.Recordset("tit_tarj") = data_rec.Recordset("tit_tarj")
                Else
                   data_cli.Recordset("tit_tarj") = ""
                End If
                If IsNull(data_rec.Recordset("cl_nrotarj")) = False Then
                   data_cli.Recordset("cl_nrotarj") = data_rec.Recordset("cl_nrotarj")
                Else
                   data_cli.Recordset("cl_nrotarj") = 0
                End If
                If IsNull(data_rec.Recordset("ci_tarj")) = False Then
                   data_cli.Recordset("ci_tarj") = data_rec.Recordset("ci_tarj")
                Else
                   data_cli.Recordset("ci_tarj") = 0
                End If
                If IsNull(data_rec.Recordset("codcitarj")) = False Then
                   data_cli.Recordset("codcitarj") = data_rec.Recordset("codcitarj")
                Else
                   data_cli.Recordset("codcitarj") = 0
                End If
                If IsNull(data_rec.Recordset("cl_tjemi_c")) = False Then
                   data_cli.Recordset("cl_tjemi_c") = data_rec.Recordset("cl_tjemi_c")
                Else
                   data_cli.Recordset("cl_tjemi_c") = 0
                End If
                If IsNull(data_rec.Recordset("cl_tjemi_n")) = False Then
                   data_cli.Recordset("cl_tjemi_n") = data_rec.Recordset("cl_tjemi_n")
                Else
                   data_cli.Recordset("cl_tjemi_n") = 0
                End If
                If IsNull(data_rec.Recordset("cl_tj_venc")) = False Then
                   data_cli.Recordset("cl_tj_venc") = data_rec.Recordset("cl_tj_venc")
                End If
                If IsNull(data_rec.Recordset("fecha_sys")) = False Then
                   data_cli.Recordset("fecha_sys") = data_rec.Recordset("fecha_sys")
                End If
                If IsNull(data_rec.Recordset("fecha_modi")) = False Then
                   data_cli.Recordset("fecha_modi") = data_rec.Recordset("fecha_modi")
                End If
                data_cli.Recordset.Update
             End If
             data_rec.Recordset.MoveNext
          Else
              data_rec.Recordset.MoveNext
          End If
      Loop
      data_rec.Recordset.MoveFirst
      Do While Not data_rec.Recordset.EOF
          data_rec.Recordset.Delete
          data_rec.Recordset.MoveNext
      Loop
   End If
   XBanbajmod = 0
   data_rec.RecordSource = "env_caja"
   data_rec.Refresh
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "Select * from caja"
   data_cli.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_cli.RecordSource = "select * from caja where fecha =#" & Format(data_rec.Recordset("fecha"), "yyyy/mm/dd") & "# and base <>" & data_pabase3.Recordset("base")
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         data_cli.Recordset.MoveFirst
         If data_parsec.Recordset("Base") = 6 Then
            Do While Not data_cli.Recordset.EOF
               If data_cli.Recordset("base") = 6 Then
                  data_cli.Recordset.MoveNext
               Else
                  data_cli.Recordset.Delete
                  data_cli.Recordset.MoveNext
               End If
            Loop
            data_rec.Recordset.MoveFirst
            Do While Not data_rec.Recordset.EOF
               If data_rec.Recordset("base") = 6 Then
                  data_rec.Recordset.Delete
               End If
               data_rec.Recordset.MoveNext
            Loop
         Else
            Do While Not data_cli.Recordset.EOF
               data_cli.Recordset.Delete
               data_cli.Recordset.MoveNext
            Loop
         End If
      End If
      data_rec.Recordset.MoveFirst
      data_cli.RecordSource = "Select * from caja"
      data_cli.Refresh
      Do While Not data_rec.Recordset.EOF
         If data_rec.Recordset("base") = data_pabase3.Recordset("base") Then
            data_rec.Recordset.MoveNext
         Else
            data_cli.Recordset.AddNew
            data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
            data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
            data_cli.Recordset("numero") = data_rec.Recordset("numero")
            data_cli.Recordset("nombre") = data_rec.Recordset("nombre")
            data_cli.Recordset("movimiento") = data_rec.Recordset("movimiento")
            data_cli.Recordset("imp_fact") = data_rec.Recordset("imp_fact")
            data_cli.Recordset("observ") = data_rec.Recordset("observ")
            data_cli.Recordset("saldo") = data_rec.Recordset("saldo")
            data_cli.Recordset("usuario") = data_rec.Recordset("usuario")
            data_cli.Recordset("hora") = data_rec.Recordset("hora")
            data_cli.Recordset("saldo_user") = data_rec.Recordset("saldo_user")
            data_cli.Recordset("base") = data_rec.Recordset("base")
            data_cli.Recordset("documento") = data_rec.Recordset("documento")
            data_cli.Recordset("cod_serv") = data_rec.Recordset("cod_serv")
            data_cli.Recordset("nom_serv") = data_rec.Recordset("nom_serv")
            data_cli.Recordset("cod_socio") = data_rec.Recordset("cod_socio")
            data_cli.Recordset("nom_socio") = data_rec.Recordset("nom_socio")
            data_cli.Recordset("caja_mesp") = data_rec.Recordset("caja_mesp")
            data_cli.Recordset("caja_anop") = data_rec.Recordset("caja_anop")
            data_cli.Recordset("imp_iva") = data_rec.Recordset("imp_iva")
            data_cli.Recordset("opiva") = data_rec.Recordset("opiva")
            data_cli.Recordset.Update
            data_rec.Recordset.MoveNext
         End If
      Loop
      data_rec.Recordset.MoveFirst
      Do While Not data_rec.Recordset.EOF
         data_rec.Recordset.Delete
         data_rec.Recordset.MoveNext
      Loop
   End If

''' Tesorería

   If data_parsec.Recordset("base") = 6 Then
      data_rec.RecordSource = "env_tes"
      data_rec.Refresh
      data_cli.DatabaseName = App.Path & "\tesorero.mdb"
      data_cli.RecordSource = "Select * from tesorero"
      data_cli.Refresh
      If data_rec.Recordset.RecordCount > 0 Then
         data_rec.Recordset.MoveFirst
         Do While Not data_rec.Recordset.EOF
            data_cli.Recordset.FindFirst "fecha =#" & Format(data_rec.Recordset("fecha"), "yyyy/mm/dd") & "# And usuario ='" & data_rec.Recordset("usuario") & "' and nromov =" & data_rec.Recordset("nromov")
            If Not data_cli.Recordset.NoMatch Then
            Else
               data_cli.Recordset.AddNew
               data_cli.Recordset("nromov") = data_rec.Recordset("nromov")
               data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
               data_cli.Recordset("hora") = data_rec.Recordset("hora")
               data_cli.Recordset("usuario") = data_rec.Recordset("usuario")
               data_cli.Recordset("cod_rub") = data_rec.Recordset("cod_rub")
               data_cli.Recordset("nom_rub") = data_rec.Recordset("nom_rub")
               data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
               data_cli.Recordset("monto") = data_rec.Recordset("monto")
               data_cli.Recordset("obs") = data_rec.Recordset("obs")
               data_cli.Recordset("cod_debe") = data_rec.Recordset("cod_debe")
               data_cli.Recordset("saldos") = data_rec.Recordset("saldos")
               data_cli.Recordset("concep") = data_rec.Recordset("concep")
               data_cli.Recordset("cod_haber") = data_rec.Recordset("cod_haber")
               data_cli.Recordset("saldou") = data_rec.Recordset("saldou")
               data_cli.Recordset("tipoc") = data_rec.Recordset("tipoc")
               data_cli.Recordset("libro") = data_rec.Recordset("libro")
               data_cli.Recordset("iva") = data_rec.Recordset("iva")
               data_cli.Recordset("base") = data_rec.Recordset("base")
               data_cli.Recordset("descon") = data_rec.Recordset("descon")
               data_cli.Recordset("bandera") = data_rec.Recordset("bandera")
               data_cli.Recordset("impiva") = data_rec.Recordset("impiva")
               data_cli.Recordset("tcam") = data_rec.Recordset("tcam")
               data_cli.Recordset.Update
            End If
            data_rec.Recordset.MoveNext
         Loop
         data_rec.Recordset.MoveFirst
         Do While Not data_rec.Recordset.EOF
            data_rec.Recordset.Delete
            data_rec.Recordset.MoveNext
         Loop
      End If
   End If
   data_rec.RecordSource = "env_lin"
   data_rec.Refresh
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "Select * from linmmdd"
   data_cli.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_cli.RecordSource = "select * from linmmdd where fecha =#" & Format(data_rec.Recordset("fecha"), "yyyy/mm/dd") & "# And base <>" & data_pabase3.Recordset("base")
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         data_cli.Recordset.MoveFirst
         If data_parsec.Recordset("Base") = 6 Then
            Do While Not data_cli.Recordset.EOF
               If data_cli.Recordset("base") = 6 Then
                  data_cli.Recordset.MoveNext
               Else
                  data_cli.Recordset.Delete
                  data_cli.Recordset.MoveNext
               End If
            Loop
            data_rec.Recordset.MoveFirst
            Do While Not data_rec.Recordset.EOF
               If data_rec.Recordset("base") = 6 Then
                  data_rec.Recordset.Delete
               End If
               data_rec.Recordset.MoveNext
            Loop
         Else
            Do While Not data_cli.Recordset.EOF
               data_cli.Recordset.Delete
               data_cli.Recordset.MoveNext
            Loop
         End If
      End If
      data_rec.Recordset.MoveFirst
      data_cli.RecordSource = "Select * from linmmdd"
      data_cli.Refresh
      Do While Not data_rec.Recordset.EOF
         If data_rec.Recordset("base") = data_pabase3.Recordset("base") Then
            data_rec.Recordset.MoveNext
         Else
            data_cli.Recordset.AddNew
            data_cli.Recordset("tipo_mov") = data_rec.Recordset("tipo_mov")
            data_cli.Recordset("factura") = data_rec.Recordset("factura")
            data_cli.Recordset("tipo") = data_rec.Recordset("tipo")
            data_cli.Recordset("realizada") = data_rec.Recordset("realizada")
            data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
            data_cli.Recordset("cod_cli") = data_rec.Recordset("cod_cli")
            data_cli.Recordset("nom_cli") = data_rec.Recordset("nom_cli")
            data_cli.Recordset("cod_prod") = data_rec.Recordset("cod_prod")
            data_cli.Recordset("nom_prod") = data_rec.Recordset("nom_prod")
            data_cli.Recordset("cantidad") = data_rec.Recordset("cantidad")
            data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
            data_cli.Recordset("operador") = data_rec.Recordset("operador")
            data_cli.Recordset("hora") = data_rec.Recordset("hora")
            data_cli.Recordset("nro_flia") = data_rec.Recordset("nro_flia")
            data_cli.Recordset("nom_flia") = data_rec.Recordset("nom_flia")
            data_cli.Recordset("linea") = data_rec.Recordset("linea")
            data_cli.Recordset("convenio") = data_rec.Recordset("convenio")
            data_cli.Recordset("rub_cont") = data_rec.Recordset("rub_cont")
            data_cli.Recordset("usa_timbre") = data_rec.Recordset("usa_timbre")
            data_cli.Recordset("imp_timbre") = data_rec.Recordset("imp_timbre")
            data_cli.Recordset("tot_lin") = data_rec.Recordset("tot_lin")
            data_cli.Recordset("rub_nomb") = data_rec.Recordset("rub_nomb")
            data_cli.Recordset("nro_med_a") = data_rec.Recordset("nro_med_a")
            data_cli.Recordset("nom_med_a") = data_rec.Recordset("nom_med_a")
            data_cli.Recordset("precio_est") = data_rec.Recordset("precio_est")
            data_cli.Recordset("mes_paga") = data_rec.Recordset("mes_paga")
            data_cli.Recordset("ano_paga") = data_rec.Recordset("ano_paga")
            data_cli.Recordset("base") = data_rec.Recordset("base")
            data_cli.Recordset("imp_iva") = data_rec.Recordset("imp_iva")
            data_cli.Recordset("ruc") = data_rec.Recordset("ruc")
            data_cli.Recordset.Update
            data_rec.Recordset.MoveNext
         End If
      Loop
      data_rec.Recordset.MoveFirst
      Do While Not data_rec.Recordset.EOF
         data_rec.Recordset.Delete
         data_rec.Recordset.MoveNext
      Loop
   End If
   Timer4.Enabled = False
   Timer7.Enabled = True
'         Kill "c:\datos\recibe\" & Trim(Xbase) & "\envios.zip"

End Sub

Private Sub Timer5_Timer()
If Dir("c:\datos\recibe\envios.zip") <> "" Then
   Shell (App.Path & "\pkunzip -o c:\datos\recibe\envios.zip c:\datos\recibe"), vbNormalFocus
   Timer5.Enabled = False
   Timer4.Enabled = True
Else
   Timer5.Enabled = False
   Timer2.Enabled = True
End If
'data_rec.DatabaseName = ""
'data_rec.RecordSource = ""
'data_rec.Refresh
'frm_correo.MousePointer = 0
'Unload Me

End Sub

Private Sub Timer6_Timer()
If data_parsec.Recordset("base") = 15 Then
   data_env.DatabaseName = "c:\Datos"
   data_env.RecordSource = "env_conv"
   data_env.Refresh
   If data_env.Recordset.RecordCount > 0 Then
      data_env.Recordset.MoveFirst
      Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
      Loop
   End If
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "convenio"
   data_cli.Refresh
'        data_cli.RecordSource = "Select * from clientes where fecha_sys =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
'        data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_env.Recordset.AddNew
         data_env.Recordset("cnv_codigo") = data_cli.Recordset("cnv_codigo")
         data_env.Recordset("cnv_desc") = data_cli.Recordset("cnv_desc")
         data_env.Recordset("cnv_colrec") = data_cli.Recordset("cnv_colrec")
         data_env.Recordset("cnv_codmon") = data_cli.Recordset("cnv_codmon")
         data_env.Recordset("cnv_desde") = data_cli.Recordset("cnv_desde")
         data_env.Recordset("cnv_hasta") = data_cli.Recordset("cnv_hasta")
         data_env.Recordset("cnv_estado") = data_cli.Recordset("cnv_estado")
         data_env.Recordset("cnv_precio") = data_cli.Recordset("cnv_precio")
         data_env.Recordset("cnv_emite") = data_cli.Recordset("cnv_emite")
         data_env.Recordset("cnv_alta") = data_cli.Recordset("cnv_alta")
         data_env.Recordset("cnv_modi") = data_cli.Recordset("cnv_modi")
         data_env.Recordset("cnv_baja") = data_cli.Recordset("cnv_baja")
         data_env.Recordset("cnv_grupo") = data_cli.Recordset("cnv_grupo")
         data_env.Recordset("cnv_cuenta") = data_cli.Recordset("cnv_cuenta")
         data_env.Recordset("cnv_cant_r") = data_cli.Recordset("cnv_cant_r")
         data_env.Recordset.Update
         data_cli.Recordset.MoveNext
      Loop
   End If
   data_env.RecordSource = "env_cob"
   data_env.Refresh
   If data_env.Recordset.RecordCount > 0 Then
      data_env.Recordset.MoveFirst
      Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
      Loop
   End If
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "cobrador"
   data_cli.Refresh
'        data_cli.RecordSource = "Select * from clientes where fecha_sys =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
'        data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_env.Recordset.AddNew
         data_env.Recordset("cb_numero") = data_cli.Recordset("cb_numero")
         data_env.Recordset("cb_nombre") = data_cli.Recordset("cb_nombre")
         data_env.Recordset.Update
         data_cli.Recordset.MoveNext
      Loop
   End If

   data_env.RecordSource = "env_estu"
   data_env.Refresh
   If data_env.Recordset.RecordCount > 0 Then
      data_env.Recordset.MoveFirst
      Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
      Loop
   End If
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "estudios"
   data_cli.Refresh
'        data_cli.RecordSource = "Select * from clientes where fecha_sys =#" & Format(Fecenv, "yyyy/mm/dd") & "#"
'        data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_env.Recordset.AddNew
         data_env.Recordset("codest") = data_cli.Recordset("codest")
         data_env.Recordset("flia") = data_cli.Recordset("flia")
         data_env.Recordset("nomflia") = data_cli.Recordset("nomflia")
         data_env.Recordset("descrip") = data_cli.Recordset("descrip")
         data_env.Recordset("cons") = data_cli.Recordset("cons")
         data_env.Recordset("uc") = data_cli.Recordset("uc")
         data_env.Recordset("part") = data_cli.Recordset("part")
         data_env.Recordset.Update
         data_cli.Recordset.MoveNext
      Loop
   End If
End If
If data_parsec.Recordset("base") = 15 Or _
   data_parsec.Recordset("base") = 13 Or _
   data_parsec.Recordset("base") = 6 Or _
   data_parsec.Recordset("base") = 1 Then
   data_env.DatabaseName = "c:\Datos"
   data_env.RecordSource = "env_arq"
   data_env.Refresh
   If data_env.Recordset.RecordCount > 0 Then
      data_env.Recordset.MoveFirst
      Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
      Loop
   End If
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "arqueo"
   data_cli.Refresh
   data_cli.RecordSource = "Select * from arqueo where fecha =#" & Format(XLafecha, "yyyy/mm/dd") & "#"
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_env.Recordset.AddNew
         data_env.Recordset("matricula") = data_cli.Recordset("matricula")
         data_env.Recordset("nombre") = data_cli.Recordset("nombre")
         data_env.Recordset("mes") = data_cli.Recordset("mes")
         data_env.Recordset("ano") = data_cli.Recordset("ano")
         data_env.Recordset("color") = data_cli.Recordset("color")
         data_env.Recordset("cat") = data_cli.Recordset("cat")
         data_env.Recordset("nomcat") = data_cli.Recordset("nomcat")
         data_env.Recordset("arqueo") = data_cli.Recordset("arqueo")
         data_env.Recordset("importe") = data_cli.Recordset("importe")
         data_env.Recordset("fecha") = data_cli.Recordset("fecha")
         data_env.Recordset("nrorec") = data_cli.Recordset("nrorec")
         data_env.Recordset("usuar") = data_cli.Recordset("usuar")
         data_env.Recordset("moneda") = data_cli.Recordset("moneda")
         data_env.Recordset("cob") = data_cli.Recordset("cob")
         data_env.Recordset("nomcob") = data_cli.Recordset("nomcob")
         data_env.Recordset("codzon") = data_cli.Recordset("codzon")
         data_env.Recordset("codpro") = data_cli.Recordset("codpro")
         data_env.Recordset("codsup") = data_cli.Recordset("codsup")
         data_env.Recordset("tiquet") = data_cli.Recordset("tiquet")
         data_env.Recordset("total") = data_cli.Recordset("total")
         data_env.Recordset("iva") = data_cli.Recordset("iva")
         data_env.Recordset("deudas") = data_cli.Recordset("deudas")
         data_env.Recordset("servi") = data_cli.Recordset("servi")
         data_env.Recordset.Update
         data_cli.Recordset.MoveNext
      Loop
   End If
End If
If data_parsec.Recordset("base") = 6 Then
   data_env.DatabaseName = "c:\Datos"
   data_env.RecordSource = "env_codc"
   data_env.Refresh
   If data_env.Recordset.RecordCount > 0 Then
      data_env.Recordset.MoveFirst
      Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
      Loop
   End If
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "cod_caja"
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_env.Recordset.AddNew
         data_env.Recordset("numero") = data_cli.Recordset("numero")
         data_env.Recordset("nombre") = data_cli.Recordset("nombre")
         data_env.Recordset("moneda") = data_cli.Recordset("moneda")
         data_env.Recordset("movimiento") = data_cli.Recordset("movimiento")
         data_env.Recordset.Update
         data_cli.Recordset.MoveNext
      Loop
   End If
   
   data_env.RecordSource = "env_rubt"
   data_env.Refresh
   If data_env.Recordset.RecordCount > 0 Then
      data_env.Recordset.MoveFirst
      Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
      Loop
   End If
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "rubteso"
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_env.Recordset.AddNew
         data_env.Recordset("codigo") = data_cli.Recordset("codigo")
         data_env.Recordset("nombre") = data_cli.Recordset("nombre")
         data_env.Recordset("es") = data_cli.Recordset("es")
         data_env.Recordset("moneda") = data_cli.Recordset("moneda")
         data_env.Recordset("debe") = data_cli.Recordset("debe")
         data_env.Recordset("concep") = data_cli.Recordset("concep")
         data_env.Recordset("haber") = data_cli.Recordset("haber")
         data_env.Recordset("cta") = data_cli.Recordset("cta")
         data_env.Recordset("libro") = data_cli.Recordset("libro")
         data_env.Recordset("moneda") = data_cli.Recordset("moneda")
         data_env.Recordset("iva") = data_cli.Recordset("iva")
         data_env.Recordset.Update
         data_cli.Recordset.MoveNext
      Loop
   End If
End If
data_env.DatabaseName = ""
data_env.RecordSource = ""
data_env.Refresh
data_cli.DatabaseName = ""
data_cli.RecordSource = ""
data_cli.Refresh
If Dir("C:\Datos\envios.zip") <> "" Then
   Kill "c:\datos\envios.zip"
End If
Shell (App.Path & "\pkzip -a c:\datos\envios.zip c:\datos\env_*.*"), vbNormalFocus
Timer6.Enabled = False
Timer1.Enabled = True

End Sub

Private Sub Timer7_Timer()
If Data1.Recordset("base") <> 15 Then
'   data_rec.DatabaseName = "c:\datos\recibe"
'   data_rec.RecordSource = "env_conv"
'   data_rec.Refresh
'   If data_rec.Recordset.RecordCount > 0 Then
'      data_rec.Recordset.MoveFirst
'      data_cli.DatabaseName = App.Path & "\sapp.mdb"
'      data_cli.RecordSource = "convenio"
'      data_cli.Refresh
'      data_cli.Recordset.MoveFirst
'      Do While Not data_cli.Recordset.EOF
'         data_cli.Recordset.Delete
'         data_cli.Recordset.MoveNext
'      Loop
'      Do While Not data_rec.Recordset.EOF
'         data_cli.Recordset.AddNew
'         data_cli.Recordset("cnv_codigo") = data_rec.Recordset("cnv_codigo")
'         data_cli.Recordset("cnv_desc") = data_rec.Recordset("cnv_desc")
'         data_cli.Recordset("cnv_colrec") = data_rec.Recordset("cnv_colrec")
'         data_cli.Recordset("cnv_codmon") = data_rec.Recordset("cnv_codmon")
'         data_cli.Recordset("cnv_desde") = data_rec.Recordset("cnv_desde")
'         data_cli.Recordset("cnv_hasta") = data_rec.Recordset("cnv_hasta")
'         data_cli.Recordset("cnv_estado") = data_rec.Recordset("cnv_estado")
'         data_cli.Recordset("cnv_precio") = data_rec.Recordset("cnv_precio")
'         data_cli.Recordset("cnv_emite") = data_rec.Recordset("cnv_emite")
'         data_cli.Recordset("cnv_alta") = data_rec.Recordset("cnv_alta")
'         data_cli.Recordset("cnv_modi") = data_rec.Recordset("cnv_modi")
'         data_cli.Recordset("cnv_baja") = data_rec.Recordset("cnv_baja")
'         data_cli.Recordset("cnv_grupo") = data_rec.Recordset("cnv_grupo")
'         data_cli.Recordset("cnv_cuenta") = data_rec.Recordset("cnv_cuenta")
'         data_cli.Recordset("cnv_cant_r") = data_rec.Recordset("cnv_cant_r")
'         data_cli.Recordset.Update
'         data_rec.Recordset.MoveNext
'      Loop
'   End If
   
'   data_rec.RecordSource = "env_cob"
'   data_rec.Refresh
'   If data_rec.Recordset.RecordCount > 0 Then
'      data_rec.Recordset.MoveFirst
'      data_cli.DatabaseName = App.Path & "\sapp.mdb"
'      data_cli.RecordSource = "cobrador"
'      data_cli.Refresh
'      data_cli.Recordset.MoveFirst
'      Do While Not data_cli.Recordset.EOF
'         data_cli.Recordset.Delete
'         data_cli.Recordset.MoveNext
'      Loop
'      Do While Not data_rec.Recordset.EOF
'         data_cli.Recordset.AddNew
'         data_cli.Recordset("cb_numero") = data_rec.Recordset("cb_numero")
'         data_cli.Recordset("cb_nombre") = data_rec.Recordset("cb_nombre")
'         data_cli.Recordset.Update
'         data_rec.Recordset.MoveNext
'      Loop
'   End If

'   data_rec.RecordSource = "env_estu"
'   data_rec.Refresh
'   If data_rec.Recordset.RecordCount > 0 Then
'      data_rec.Recordset.MoveFirst
'      data_cli.DatabaseName = App.Path & "\sapp.mdb"
'      data_cli.RecordSource = "estudios"
'      data_cli.Refresh
'      data_cli.Recordset.MoveFirst
'      Do While Not data_cli.Recordset.EOF
'         data_cli.Recordset.Delete
'         data_cli.Recordset.MoveNext
'      Loop
'      Do While Not data_rec.Recordset.EOF
'         data_cli.Recordset.AddNew
'         data_cli.Recordset("codest") = data_rec.Recordset("codest")
'         data_cli.Recordset("flia") = data_rec.Recordset("flia")
'         data_cli.Recordset("nomflia") = data_rec.Recordset("nomflia")
'         data_cli.Recordset("descrip") = data_rec.Recordset("descrip")
'         data_cli.Recordset("cons") = data_rec.Recordset("cons")
'         data_cli.Recordset("uc") = data_rec.Recordset("uc")
'         data_cli.Recordset("part") = data_rec.Recordset("part")
'         data_cli.Recordset.Update
'         data_rec.Recordset.MoveNext
'      Loop
'   End If
End If
Dim Xarqvacio As Integer
data_rec.DatabaseName = "c:\datos\recibe"

If Data1.Recordset("base") = 15 Or _
   Data1.Recordset("base") = 13 Or _
   Data1.Recordset("base") = 6 Or _
   Data1.Recordset("base") = 1 Then
   data_rec.RecordSource = "env_arq"
   data_rec.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveLast
      data_rec.Recordset.MoveFirst
      data_cli.DatabaseName = App.Path & "\sapp.mdb"
      data_cli.RecordSource = "arqueo"
      data_cli.Refresh
      If data_rec.Recordset.RecordCount > 6000 Then
         data_cli.Recordset.MoveFirst
         Xarqvacio = 1
         Do While Not data_cli.Recordset.EOF
            data_cli.Recordset.Delete
            data_cli.Recordset.MoveNext
         Loop
      Else
         Xarqvacio = 0
      End If
      Do While Not data_rec.Recordset.EOF
         If Xarqvacio = 1 Then
            data_cli.Recordset.AddNew
            data_cli.Recordset("matricula") = data_rec.Recordset("matricula")
            data_cli.Recordset("nombre") = Mid(data_rec.Recordset("nombre"), 1, 30)
            data_cli.Recordset("mes") = data_rec.Recordset("mes")
            data_cli.Recordset("ano") = data_rec.Recordset("ano")
            data_cli.Recordset("color") = data_rec.Recordset("color")
            data_cli.Recordset("cat") = Mid(data_rec.Recordset("cat"), 1, 6)
            data_cli.Recordset("nomcat") = Mid(data_rec.Recordset("nomcat"), 1, 25)
            data_cli.Recordset("arqueo") = data_rec.Recordset("arqueo")
            data_cli.Recordset("importe") = data_rec.Recordset("importe")
            data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
            data_cli.Recordset("nrorec") = data_rec.Recordset("nrorec")
            data_cli.Recordset("usuar") = Mid(data_rec.Recordset("usuar"), 1, 10)
            data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
            data_cli.Recordset("cob") = data_rec.Recordset("cob")
            data_cli.Recordset("nomcob") = Mid(data_rec.Recordset("nomcob"), 1, 25)
            data_cli.Recordset("codzon") = data_rec.Recordset("codzon")
            data_cli.Recordset("codpro") = data_rec.Recordset("codpro")
            data_cli.Recordset("codsup") = data_rec.Recordset("codsup")
            data_cli.Recordset("tiquet") = data_rec.Recordset("tiquet")
            data_cli.Recordset("total") = data_rec.Recordset("total")
            data_cli.Recordset("iva") = data_rec.Recordset("iva")
            data_cli.Recordset("deudas") = data_rec.Recordset("deudas")
            data_cli.Recordset("servi") = data_rec.Recordset("servi")
            data_cli.Recordset.Update
            data_rec.Recordset.MoveNext
         Else
            data_cli.Recordset.FindFirst "nrorec =" & data_rec.Recordset("nrorec") & " and matricula =" & data_rec.Recordset("matricula")
            If Not data_cli.Recordset.NoMatch Then
               data_cli.Recordset.Edit
               data_cli.Recordset("matricula") = data_rec.Recordset("matricula")
               data_cli.Recordset("nombre") = Mid(data_rec.Recordset("nombre"), 1, 30)
               data_cli.Recordset("mes") = data_rec.Recordset("mes")
               data_cli.Recordset("ano") = data_rec.Recordset("ano")
               data_cli.Recordset("color") = data_rec.Recordset("color")
               data_cli.Recordset("cat") = Mid(data_rec.Recordset("cat"), 1, 6)
               data_cli.Recordset("nomcat") = Mid(data_rec.Recordset("nomcat"), 1, 25)
               data_cli.Recordset("arqueo") = data_rec.Recordset("arqueo")
               data_cli.Recordset("importe") = data_rec.Recordset("importe")
               data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
               data_cli.Recordset("nrorec") = data_rec.Recordset("nrorec")
               data_cli.Recordset("usuar") = Mid(data_rec.Recordset("usuar"), 1, 10)
               data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
               data_cli.Recordset("cob") = data_rec.Recordset("cob")
               data_cli.Recordset("nomcob") = Mid(data_rec.Recordset("nomcob"), 1, 25)
               data_cli.Recordset("codzon") = data_rec.Recordset("codzon")
               data_cli.Recordset("codpro") = data_rec.Recordset("codpro")
               data_cli.Recordset("codsup") = data_rec.Recordset("codsup")
               data_cli.Recordset("tiquet") = data_rec.Recordset("tiquet")
               data_cli.Recordset("total") = data_rec.Recordset("total")
               data_cli.Recordset("iva") = data_rec.Recordset("iva")
               data_cli.Recordset("deudas") = data_rec.Recordset("deudas")
               data_cli.Recordset("servi") = data_rec.Recordset("servi")
               data_cli.Recordset.Update
               data_rec.Recordset.MoveNext
            Else
               data_cli.Recordset.AddNew
               data_cli.Recordset("matricula") = data_rec.Recordset("matricula")
               data_cli.Recordset("nombre") = Mid(data_rec.Recordset("nombre"), 1, 30)
               data_cli.Recordset("mes") = data_rec.Recordset("mes")
               data_cli.Recordset("ano") = data_rec.Recordset("ano")
               data_cli.Recordset("color") = data_rec.Recordset("color")
               data_cli.Recordset("cat") = Mid(data_rec.Recordset("cat"), 1, 6)
               data_cli.Recordset("nomcat") = Mid(data_rec.Recordset("nomcat"), 1, 25)
               data_cli.Recordset("arqueo") = data_rec.Recordset("arqueo")
               data_cli.Recordset("importe") = data_rec.Recordset("importe")
               data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
               data_cli.Recordset("nrorec") = data_rec.Recordset("nrorec")
               data_cli.Recordset("usuar") = Mid(data_rec.Recordset("usuar"), 1, 10)
               data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
               data_cli.Recordset("cob") = data_rec.Recordset("cob")
               data_cli.Recordset("nomcob") = Mid(data_rec.Recordset("nomcob"), 1, 25)
               data_cli.Recordset("codzon") = data_rec.Recordset("codzon")
               data_cli.Recordset("codpro") = data_rec.Recordset("codpro")
               data_cli.Recordset("codsup") = data_rec.Recordset("codsup")
               data_cli.Recordset("tiquet") = data_rec.Recordset("tiquet")
               data_cli.Recordset("total") = data_rec.Recordset("total")
               data_cli.Recordset("iva") = data_rec.Recordset("iva")
               data_cli.Recordset("deudas") = data_rec.Recordset("deudas")
               data_cli.Recordset("servi") = data_rec.Recordset("servi")
               data_cli.Recordset.Update
               data_rec.Recordset.MoveNext
            End If
         End If
      Loop
      Xarqvacio = 0
   End If
End If
If Data1.Recordset("base") <> 6 Then
'   data_rec.RecordSource = "env_codc"
'   data_rec.Refresh
'   If data_rec.Recordset.RecordCount > 0 Then
'      data_rec.Recordset.MoveFirst
'      data_cli.DatabaseName = App.Path & "\sapp.mdb"
'      data_cli.RecordSource = "cod_caja"
'      data_cli.Refresh
'      data_cli.Recordset.MoveFirst
'      Do While Not data_cli.Recordset.EOF
'         data_cli.Recordset.Delete
'         data_cli.Recordset.MoveNext
'      Loop
'      Do While Not data_rec.Recordset.EOF
'         data_cli.Recordset.AddNew
'         data_cli.Recordset("numero") = data_rec.Recordset("numero")
'         data_cli.Recordset("nombre") = data_rec.Recordset("nombre")
'         data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
'         data_cli.Recordset("movimiento") = data_rec.Recordset("movimiento")
'         data_cli.Recordset.Update
'         data_rec.Recordset.MoveNext
'      Loop
'   End If
   
   data_rec.RecordSource = "env_rubt"
   data_rec.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveFirst
      data_cli.DatabaseName = App.Path & "\sapp.mdb"
      data_cli.RecordSource = "rubteso"
      data_cli.Refresh
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_cli.Recordset.Delete
         data_cli.Recordset.MoveNext
      Loop
      Do While Not data_rec.Recordset.EOF
         data_cli.Recordset.AddNew
         data_cli.Recordset("codigo") = data_rec.Recordset("codigo")
         data_cli.Recordset("nombre") = data_rec.Recordset("nombre")
         data_cli.Recordset("es") = data_rec.Recordset("es")
         data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
         data_cli.Recordset("debe") = data_rec.Recordset("debe")
         data_cli.Recordset("concep") = data_rec.Recordset("concep")
         data_cli.Recordset("haber") = data_rec.Recordset("haber")
         data_cli.Recordset("cta") = data_rec.Recordset("cta")
         data_cli.Recordset("libro") = data_rec.Recordset("libro")
         data_cli.Recordset("iva") = data_rec.Recordset("iva")
         data_cli.Recordset.Update
         data_rec.Recordset.MoveNext
      Loop
   End If
End If
data_rec.DatabaseName = ""
data_rec.RecordSource = ""
data_rec.Refresh
Kill "c:\datos\recibe\envios.zip"
Shell (App.Path & "\borrar.bat"), vbMinimizedFocus
Timer7.Enabled = False
Timer2.Enabled = True

End Sub

Private Sub Timer8_Timer()
Xcontatime = Xcontatime + 1
If Xcontatime = 13 Then
   Xcontatime = 0
   If Data2.Recordset("siono") = "SI" Then
      XCuantosd = Date - Data1.Recordset("ult_env")
      If XCuantosd > 1 Then
         Timer3.Enabled = True
         Timer8.Enabled = False
      Else
         Timer2.Enabled = True
         Timer8.Enabled = False
      End If
   Else
      Timer8.Enabled = False
      Unload Me
   End If
End If

End Sub
