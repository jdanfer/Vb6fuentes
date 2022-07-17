VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form correo 
   BorderStyle     =   0  'None
   Caption         =   "Envio y Recepción"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   DrawStyle       =   5  'Transparent
   Icon            =   "correo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   1440
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   4200
      Top             =   840
   End
   Begin VB.Data data_busca 
      Caption         =   "data_busca"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_rec 
      Caption         =   "data_rec"
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   2280
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
      Enabled         =   0   'False
      Interval        =   100
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
      Height          =   300
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PARSEC0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   100
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sistema de Envío y Recepción automática de correo"
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
Attribute VB_Name = "correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Fecha As Date
Dim Archivo As String
Dim Remite As String
Dim Asunto As String
Dim Contra As String
Contra = "sapp1987"
'Envio de e-mail desde VB:
'1.- Adjuntar al proyecto los controles MAPI
'(ya sabes: Proyecto/Componentes y señalar Microsoft MAPI controls)
'2.- En tu formulario, coloca los controles MAPISession y MAPIMessages
'3.- Para enviar el mail:
Fecha = Data1.Recordset("ult_envio")
If Month(Fecha) > 9 Then
   If Day(Fecha) > 9 Then
      Archivo = Trim(Str(Year(Format(Fecha, "dd/mm/yyyy")))) + Trim(Str(Month(Format(Fecha, "dd/mm/yyyy")))) + Trim(Str(Day(Format(Fecha, "dd/mm/yyyy")))) + ".zip"
   Else
      Archivo = Trim(Str(Year(Format(Fecha, "dd/mm/yyyy")))) + Trim(Str(Month(Format(Fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(Fecha, "dd/mm/yyyy")))) + ".zip"
   End If
Else
   If Day(Fecha) > 9 Then
      Archivo = Trim(Str(Year(Format(Fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(Fecha, "dd/mm/yyyy")))) + Trim(Str(Day(Format(Fecha, "dd/mm/yyyy")))) + ".zip"
   Else
      Archivo = Trim(Str(Year(Format(Fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(Fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(Fecha, "dd/mm/yyyy")))) + ".zip"
   End If
End If
If Data1.Recordset("base") = 1 Then
   Remite = "sapp01@adinet.com.uy"
End If
If Data1.Recordset("base") = 2 Then
   Remite = "sapp003@adinet.com.uy"
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
MAPISession1.Password = Contra
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
Dim Fecha As Date
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
Fecha = Data1.Recordset("ult_envio")
If Month(Fecha) > 9 Then
   If Day(Fecha) > 9 Then
      Archivo = Trim(Str(Year(Format(Fecha, "dd/mm/yyyy")))) + Trim(Str(Month(Format(Fecha, "dd/mm/yyyy")))) + Trim(Str(Day(Format(Fecha, "dd/mm/yyyy")))) + ".zip"
   Else
      Archivo = Trim(Str(Year(Format(Fecha, "dd/mm/yyyy")))) + Trim(Str(Month(Format(Fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(Fecha, "dd/mm/yyyy")))) + ".zip"
   End If
Else
   If Day(Fecha) > 9 Then
      Archivo = Trim(Str(Year(Format(Fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(Fecha, "dd/mm/yyyy")))) + Trim(Str(Day(Format(Fecha, "dd/mm/yyyy")))) + ".zip"
   Else
      Archivo = Trim(Str(Year(Format(Fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Month(Format(Fecha, "dd/mm/yyyy")))) + "0" + Trim(Str(Day(Format(Fecha, "dd/mm/yyyy")))) + ".zip"
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

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\parsec0.mdb"

End Sub

Private Sub Timer1_Timer()
Dim Fecha As Date
Dim Archivo As String
Dim Remite As String
Dim Asunto As String
Dim nCanMsg As Integer
Dim cNomFic As String
Dim nX As Integer
Dim nY As Integer
Dim Rei As Integer

Timer1.Enabled = False
'Rei = ExitWindowsEx(2, 0&) 'Reinicia el Sistema
End

End Sub

Private Sub Timer2_Timer()
Dim Bajaarch As String
Dim Remibaj As String
Dim nCanMsg As Integer
Dim cNomFic As String
Dim nX As Integer
Dim nY As Integer
Dim Verarch As String
Bajaarch = "envios.zip"
'If Dir("c:\datos\recibe\envios.mdb") <> "" Then
'   FileCopy "c:\datos\envios.mdb", "c:\datos\envtot.mdb"
'   Kill ("c:\datos\envtot.mdb")

    If Data1.Recordset("base") = 1 Then
       Remibaj = "sapp01@adinet.com.uy"
    End If
    If Data1.Recordset("base") = 2 Then
       Remibaj = "sapp02@adinet.com.uy"
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
    MAPISession1.LogonUI = True
    MAPISession1.DownLoadMail = True
'    MAPISession1.SignOff
    MAPISession1.SignOn
    
    MAPIMessages1.SessionID = MAPISession1.SessionID
    MAPIMessages1.FetchUnreadOnly = True ' Solo los no leidos
    MAPIMessages1.FetchSorted = True ' ordenados segun llegada
    MAPIMessages1.Fetch ' obtiene el conjunto de mensajes
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
            Xquebase = 3
            If Xquebase = 1 Or _
                Xquebase = 2 Or _
                Xquebase = 3 Or _
                Xquebase = 4 Or _
                Xquebase = 6 Or _
                Xquebase = 8 Or _
                Xquebase = 9 Or _
                Xquebase = 10 Or _
                Xquebase = 11 Or _
                Xquebase = 13 Or _
                Xquebase = 15 Or _
                Xquebase = 16 Or _
                Xquebase = 17 Then
            Else
                Xquebase = 99
            End If
    
          MAPIMessages1.FetchUnreadOnly = False
          MAPIMessages1.Fetch ' obtiene el conjunto de mensajes
    
          nX = 0
          MAPIMessages1.MsgIndex = 0
          MAPIMessages1.Delete (mapMessageDelete)
    
          MAPISession1.SignOff
          Timer2.Enabled = False
          Timer5.Enabled = True
            
        Else
           If MAPIMessages1.MsgSubject = "Envio B1" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base1\" + cNomFic
              Next
              Xquebase = 1
           End If
           If MAPIMessages1.MsgSubject = "Envio B2" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base2\" + cNomFic
              Next
              Xquebase = 2
           End If
           If MAPIMessages1.MsgSubject = "Envio B4" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base4\" + cNomFic
              Next
              Xquebase = 4
           End If
           If MAPIMessages1.MsgSubject = "Envio B6" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base6\" + cNomFic
              Next
              Xquebase = 6
           End If
           If MAPIMessages1.MsgSubject = "Envio B8" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base8\" + cNomFic
              Next
              Xquebase = 8
           End If
           If MAPIMessages1.MsgSubject = "Envio B16" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base16\" + cNomFic
              Next
              Xquebase = 16
           End If
           If MAPIMessages1.MsgSubject = "Envio B10" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base10\" + cNomFic
              Next
              Xquebase = 10
           End If
           If MAPIMessages1.MsgSubject = "Envio B17" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base17\" + cNomFic
              Next
              Xquebase = 17
           End If
           If MAPIMessages1.MsgSubject = "Envio B13" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base13\" + cNomFic
              Next
              Xquebase = 13
           End If
           If MAPIMessages1.MsgSubject = "Envio B15" Then
    '      Si te interesa el texto del mensaje, está en MAPIMessages1.MsgNoteText
    '      Por cada archivo anexado al mensaje, extraerlo y copiarlo donde queramos
           Verarch = MAPIMessages1.AttachmentName
              For nY = 0 To MAPIMessages1.AttachmentCount - 1
                  MAPIMessages1.AttachmentIndex = nY
                  cNomFic = ExtraerNombreArchivo(MAPIMessages1.AttachmentName)
                  FileCopy MAPIMessages1.AttachmentPathName, "C:\Datos\Recibe\base15\" + cNomFic
              Next
              Xquebase = 15
           End If
           If Xquebase = 1 Or _
              Xquebase = 2 Or _
              Xquebase = 3 Or _
              Xquebase = 4 Or _
              Xquebase = 6 Or _
              Xquebase = 8 Or _
              Xquebase = 9 Or _
              Xquebase = 10 Or _
              Xquebase = 11 Or _
              Xquebase = 13 Or _
              Xquebase = 15 Or _
              Xquebase = 16 Or _
              Xquebase = 17 Then
           Else
              Xquebase = 99
           End If
    
           MAPIMessages1.FetchUnreadOnly = False
           MAPIMessages1.Fetch ' obtiene el conjunto de mensajes
    
           nX = 0
           MAPIMessages1.MsgIndex = 0
           MAPIMessages1.Delete (mapMessageDelete)
    
           MAPISession1.SignOff
           Timer2.Enabled = False
           Timer5.Enabled = True
      End If
'      MAPISession1.SignOff
    
    Else
       Xquebase = 99
       MAPISession1.SignOff
       Timer2.Enabled = False
       frm_inicia.Show
       frm_inicia.Timer1.Enabled = True
    
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
   If Data1.Recordset("base") <> 15 Then
      data_rec.RecordSource = "env_clib"
      data_rec.Refresh
      If data_rec.Recordset.RecordCount > 0 Then
         data_rec.Recordset.MoveFirst
         data_cli.DatabaseName = App.Path & "\sapp.mdb"
         data_cli.RecordSource = "clientes"
         data_cli.Refresh
         Do While Not data_rec.Recordset.EOF
             data_cli.Recordset.FindFirst "cl_codigo =" & data_rec.Recordset("cl_codigo")
             If Not data_cli.Recordset.NoMatch Then
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
   End If

   data_rec.RecordSource = "env_caja"
   data_rec.Refresh
   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.RecordSource = "Select * from caja"
   data_cli.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveFirst
      data_cli.Recordset.FindFirst "base =" & data_rec.Recordset("base") & " And fecha =#" & Format(data_rec.Recordset("fecha"), "yyyy/mm/dd") & "#"
      If data_cli.Recordset.NoMatch Then
         Do While Not data_rec.Recordset.EOF
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
         Loop
      End If
      data_rec.Recordset.MoveFirst
      Do While Not data_rec.Recordset.EOF
         data_rec.Recordset.Delete
         data_rec.Recordset.MoveNext
      Loop
   End If

''' Tesorería

   If Xquebase = 6 Or _
      Xquebase = 15 Or _
      Xquebase = 1 Then
      data_rec.RecordSource = "env_tes"
      data_rec.Refresh
      data_cli.DatabaseName = App.Path & "\tesorero.mdb"
      data_cli.RecordSource = "Select * from tesorero"
      data_cli.Refresh
      If data_rec.Recordset.RecordCount > 0 Then
         data_rec.Recordset.MoveFirst
         data_cli.Recordset.FindFirst "fecha =#" & data_rec.Recordset("fecha") & "# And usuario ='" & data_rec.Recordset("usuario") & "'"
         If Not data_cli.Recordset.NoMatch Then
         Else
            Do While Not data_rec.Recordset.EOF
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
               data_rec.Recordset.MoveNext
            Loop
         End If
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
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveFirst
      data_cli.RecordSource = "Select * from linmmdd"
      data_cli.Refresh
      data_cli.Recordset.FindFirst "fecha =#" & data_rec.Recordset("fecha") & "# And base =" & data_rec.Recordset("base")
      If Not data_cli.Recordset.NoMatch Then
         data_cli.Recordset.MoveNext
      Else
         Do While Not data_rec.Recordset.EOF
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
         Loop
      End If
      data_rec.Recordset.MoveFirst
      Do While Not data_rec.Recordset.EOF
         data_rec.Recordset.Delete
         data_rec.Recordset.MoveNext
      Loop
   End If
   Timer4.Enabled = False
   Timer6.Enabled = True

End Sub

Private Sub Timer5_Timer()
Dim Xnomb, Xbase As String
If Xquebase = 1 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 2 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 3 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 4 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 6 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 8 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 9 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 10 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 11 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 13 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 15 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 16 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase = 17 Then
   Xbase = "base" + Trim(Str(Xquebase))
End If
If Xquebase <> 99 Then
   If Xquebase <> 3 Then
      If Dir("c:\datos\recibe\" & Trim(Xbase) & "\envios.zip") <> "" Then
         Shell (App.Path & "\pkunzip -o c:\datos\recibe\" & Trim(Xbase) & "\envios.zip c:\datos\recibe"), vbNormalFocus
         Timer5.Enabled = False
         Timer4.Enabled = True
      Else
         Timer5.Enabled = False
         Timer2.Enabled = True
      End If
   Else
      Timer5.Enabled = False
      Timer2.Enabled = True
   End If
Else
   Timer5.Enabled = False
   Timer2.Enabled = True
End If

End Sub

Private Sub Timer6_Timer()
If Xquebase = 15 Then
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
'
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
If Xquebase = 15 Or _
   Xquebase = 13 Or _
   Xquebase = 6 Or _
   Xquebase = 1 Then
   data_rec.DatabaseName = "c:\datos\recibe"
   data_rec.RecordSource = "env_arq"
   data_rec.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveFirst
      data_cli.DatabaseName = App.Path & "\sapp.mdb"
      data_cli.RecordSource = "arqueo"
      data_cli.Refresh
      Do While Not data_rec.Recordset.EOF
        data_cli.Recordset.FindFirst "nrorec =" & data_rec.Recordset("nrorec") & " and matricula =" & data_rec.Recordset("matricula")
        If Not data_cli.Recordset.NoMatch Then
           data_cli.Recordset.Edit
           data_cli.Recordset("matricula") = data_rec.Recordset("matricula")
           data_cli.Recordset("nombre") = data_rec.Recordset("nombre")
           data_cli.Recordset("mes") = data_rec.Recordset("mes")
           data_cli.Recordset("ano") = data_rec.Recordset("ano")
           data_cli.Recordset("color") = data_rec.Recordset("color")
           data_cli.Recordset("cat") = data_rec.Recordset("cat")
           data_cli.Recordset("nomcat") = data_rec.Recordset("nomcat")
           data_cli.Recordset("arqueo") = data_rec.Recordset("arqueo")
           data_cli.Recordset("importe") = data_rec.Recordset("importe")
           data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
           data_cli.Recordset("nrorec") = data_rec.Recordset("nrorec")
           data_cli.Recordset("usuar") = data_rec.Recordset("usuar")
           data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
           data_cli.Recordset("cob") = data_rec.Recordset("cob")
           data_cli.Recordset("nomcob") = data_rec.Recordset("nomcob")
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
           data_cli.Recordset("nombre") = data_rec.Recordset("nombre")
           data_cli.Recordset("mes") = data_rec.Recordset("mes")
           data_cli.Recordset("ano") = data_rec.Recordset("ano")
           data_cli.Recordset("color") = data_rec.Recordset("color")
           data_cli.Recordset("cat") = data_rec.Recordset("cat")
           data_cli.Recordset("nomcat") = data_rec.Recordset("nomcat")
           data_cli.Recordset("arqueo") = data_rec.Recordset("arqueo")
           data_cli.Recordset("importe") = data_rec.Recordset("importe")
           data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
           data_cli.Recordset("nrorec") = data_rec.Recordset("nrorec")
           data_cli.Recordset("usuar") = data_rec.Recordset("usuar")
           data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
           data_cli.Recordset("cob") = data_rec.Recordset("cob")
           data_cli.Recordset("nomcob") = data_rec.Recordset("nomcob")
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
      Loop
   End If
End If
If Xquebase = 6 Then
   data_rec.RecordSource = "env_codc"
   data_rec.Refresh
   If data_rec.Recordset.RecordCount > 0 Then
      data_rec.Recordset.MoveFirst
      data_cli.DatabaseName = App.Path & "\sapp.mdb"
      data_cli.RecordSource = "cod_caja"
      data_cli.Refresh
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_cli.Recordset.Delete
         data_cli.Recordset.MoveNext
      Loop
      Do While Not data_rec.Recordset.EOF
         data_cli.Recordset.AddNew
         data_cli.Recordset("numero") = data_rec.Recordset("numero")
         data_cli.Recordset("nombre") = data_rec.Recordset("nombre")
         data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
         data_cli.Recordset("movimiento") = data_rec.Recordset("movimiento")
         data_cli.Recordset.Update
         data_rec.Recordset.MoveNext
      Loop
   End If
   
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
         data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
         data_cli.Recordset("iva") = data_rec.Recordset("iva")
         data_cli.Recordset.Update
         data_rec.Recordset.MoveNext
      Loop
   End If
End If

data_rec.DatabaseName = ""
data_rec.RecordSource = ""
data_rec.Refresh
Kill "c:\datos\recibe\base" & Trim(Str(Xquebase)) & "\envios.zip"
If Xquebase = 1 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb1") = Data1.Recordset("recb1") + 1
   Data1.Recordset.Update
End If
If Xquebase = 2 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb2") = Data1.Recordset("recb2") + 1
   Data1.Recordset.Update
End If
If Xquebase = 4 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb4") = Data1.Recordset("recb4") + 1
   Data1.Recordset.Update
End If
If Xquebase = 6 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb6") = Data1.Recordset("recb6") + 1
   Data1.Recordset.Update
End If
If Xquebase = 8 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb8") = Data1.Recordset("recb8") + 1
   Data1.Recordset.Update
End If
If Xquebase = 10 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb10") = Data1.Recordset("recb10") + 1
   Data1.Recordset.Update
End If
If Xquebase = 13 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb13") = Data1.Recordset("recb13") + 1
   Data1.Recordset.Update
End If
If Xquebase = 15 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb15") = Data1.Recordset("recb15") + 1
   Data1.Recordset.Update
End If
If Xquebase = 16 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb16") = Data1.Recordset("recb16") + 1
   Data1.Recordset.Update
End If
If Xquebase = 17 Then
   Data1.Recordset.Edit
   Data1.Recordset("recb17") = Data1.Recordset("recb17") + 1
   Data1.Recordset.Update
End If

data_rec.DatabaseName = ""
data_rec.RecordSource = ""
data_rec.Refresh
Shell (App.Path & "\borrar.bat"), vbMinimizedFocus
Timer6.Enabled = False
Timer2.Enabled = True

End Sub
