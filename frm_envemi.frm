VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_envemi 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar emisión..."
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_envemi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_env 
      Caption         =   "data_env"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   2880
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2760
      Top             =   1200
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
      Caption         =   "SALIR"
      Height          =   855
      Left            =   3840
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ACEPTAR"
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txt_a 
      Height          =   495
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txt_m 
      Height          =   495
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mes/Año de emisión a ENVIAR:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frm_envemi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Nombre As String

frm_envemi.MousePointer = 11
Nombre = "EMI"
If txt_m.Text > 9 Then
   Nombre = Nombre + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
Else
   Nombre = Nombre + "0" + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
End If
If txt_m.Text <> "" Then
   If txt_a.Text <> "" Then
      If data_env.Recordset.RecordCount > 0 Then
         data_env.Recordset.MoveFirst
         Do While Not data_env.Recordset.EOF
            data_env.Recordset.Delete
            data_env.Recordset.MoveNext
         Loop
      End If
      data_emi.RecordSource = "Select * from " & Nombre & " "
      data_emi.Refresh
      If data_emi.Recordset.RecordCount > 0 Then
         data_emi.Recordset.MoveFirst
         Do While Not data_emi.Recordset.EOF
            data_env.Recordset.AddNew
            data_env.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
            data_env.Recordset("nom_cnv") = data_emi.Recordset("nom_cnv")
            data_env.Recordset("cliente") = data_emi.Recordset("cliente")
            data_env.Recordset("apellidos") = data_emi.Recordset("apellidos")
            data_env.Recordset("tipocta") = data_emi.Recordset("tipocta")
            data_env.Recordset("ruc") = data_emi.Recordset("ruc")
            data_env.Recordset("cedula") = data_emi.Recordset("cedula")
            data_env.Recordset("fecha") = data_emi.Recordset("fecha")
            data_env.Recordset("tipodoc") = data_emi.Recordset("tipodoc")
            data_env.Recordset("tipo") = data_emi.Recordset("tipo")
            data_env.Recordset("moneda") = data_emi.Recordset("moneda")
            data_env.Recordset("debe_haber") = data_emi.Recordset("debe_haber")
            data_env.Recordset("origen") = data_emi.Recordset("origen")
            data_env.Recordset("operador") = data_emi.Recordset("operador")
            data_env.Recordset("hora") = data_emi.Recordset("hora")
            data_env.Recordset("nro_superv") = data_emi.Recordset("nro_superv")
            data_env.Recordset("nom_superv") = data_emi.Recordset("nom_superv")
            data_env.Recordset("nro_vende") = data_emi.Recordset("nro_vende")
            data_env.Recordset("nom_vende") = data_emi.Recordset("nom_vende")
            data_env.Recordset("zona") = data_emi.Recordset("zona")
            data_env.Recordset("dir_cli") = data_emi.Recordset("dir_cli")
            data_env.Recordset("loc_cli") = data_emi.Recordset("loc_cli")
            data_env.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
            data_env.Recordset("grupo") = data_emi.Recordset("grupo")
            data_env.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
            data_env.Recordset("fecha_nac") = data_emi.Recordset("fecha_nac")
            data_env.Recordset("documento") = data_emi.Recordset("documento")
            data_env.Recordset("importe") = data_emi.Recordset("importe")
            data_env.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
            data_env.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
            data_env.Recordset("mes") = data_emi.Recordset("mes")
            data_env.Recordset("ano") = data_emi.Recordset("ano")
            data_env.Recordset("color_rec") = data_emi.Recordset("color_rec")
            data_env.Recordset("tiquet") = data_emi.Recordset("tiquet")
            data_env.Recordset("servi") = data_emi.Recordset("servi")
            data_env.Recordset("deudas") = data_emi.Recordset("deudas")
            data_env.Recordset("iva") = data_emi.Recordset("iva")
            data_env.Recordset("total") = data_emi.Recordset("total")
            data_env.Recordset.Update
            data_emi.Recordset.MoveNext
         Loop
         data_env.DatabaseName = ""
         data_env.RecordSource = ""
         data_env.Refresh
         Shell "c:\Winzip\winzip32.exe" & " -a C:\datos\emi.zip" & " C:\datos\env_emi.mdb", vbMaximizedFocus
        MAPISession1.UserName = "sapp10@adinet.com.uy"
        MAPISession1.NewSession = True
        MAPISession1.DownLoadMail = False ' o false si no deseas recibir
        MAPISession1.SignOn
        MAPIMessages1.SessionID = MAPISession1.SessionID
        
        MAPIMessages1.MsgIndex = -1 ' nuevo mensaje
        MAPIMessages1.RecipDisplayName = "sappjorge@hotmail.com"
        
        MAPIMessages1.MsgSubject = "Emisión"
        MAPIMessages1.MsgNoteText = ""
        
        MAPIMessages1.AttachmentIndex = MAPIMessages1.AttachmentCount
        MAPIMessages1.AttachmentName = "emi.zip"
        MAPIMessages1.AttachmentPathName = "C:\Datos\emi.zip"
        MAPIMessages1.AttachmentPosition = MAPIMessages1.AttachmentIndex
        MAPIMessages1.AttachmentType = vbAttachTypeData
        
        MAPIMessages1.Send
        
        MAPISession1.SignOff
      
      
      End If
   End If
End If
frm_envemi.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()

'Winzip = fso.GetParentFolderName(Direcciondewinzip)

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
data_env.DatabaseName = "C:\Datos\env_emi.mdb"
data_env.RecordSource = "env_emi"
data_env.Refresh
data_emi.DatabaseName = App.Path & "\emisiones.mdb"
txt_m.Text = Month(Date)
txt_a.Text = Year(Date)

End Sub
