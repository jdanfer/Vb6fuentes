VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_resplla 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Respaldo de llamados"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6945
   Icon            =   "frm_resplla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_resp 
      Caption         =   "data_resp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_lla 
      Caption         =   "data_lla"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "PROCESAR..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "HASTA QUE FECHA?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frm_resplla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_resplla.MousePointer = 11
Command1.Enabled = False
If mh.Text <> "__/__/____" Then
   data_lla.RecordSource = "Select * from llamado where fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
   data_lla.Refresh
   If data_lla.Recordset.RecordCount > 0 Then
      data_lla.Recordset.MoveFirst
      Do While Not data_lla.Recordset.EOF
         data_resp.Recordset.AddNew
        data_resp.Recordset("nro") = data_lla.Recordset("nro")
        data_resp.Recordset("fecha") = data_lla.Recordset("fecha")
        data_resp.Recordset("hora") = data_lla.Recordset("hora")
        data_resp.Recordset("usuario") = data_lla.Recordset("usuario")
        data_resp.Recordset("matric") = data_lla.Recordset("matric")
        data_resp.Recordset("nombre") = data_lla.Recordset("nombre")
        data_resp.Recordset("edad") = data_lla.Recordset("edad")
        data_resp.Recordset("unied") = data_lla.Recordset("unied")
        data_resp.Recordset("categ") = data_lla.Recordset("categ")
        data_resp.Recordset("nomcat") = data_lla.Recordset("nomcat")
        data_resp.Recordset("ci") = data_lla.Recordset("ci")
        data_resp.Recordset("direcc") = data_lla.Recordset("direcc")
        data_resp.Recordset("telef") = data_lla.Recordset("telef")
        data_resp.Recordset("codzon") = data_lla.Recordset("codzon")
        data_resp.Recordset("base") = data_lla.Recordset("base")
        data_resp.Recordset("referen") = data_lla.Recordset("referen")
        data_resp.Recordset("motcon") = data_lla.Recordset("motcon")
        data_resp.Recordset("obsmot") = data_lla.Recordset("obsmot")
        data_resp.Recordset("codmot") = data_lla.Recordset("codmot")
        data_resp.Recordset("descol") = data_lla.Recordset("descol")
        data_resp.Recordset("movilpas") = data_lla.Recordset("movilpas")
        data_resp.Recordset("pend") = data_lla.Recordset("pend")
        data_resp.Recordset("fec_rea") = data_lla.Recordset("fec_rea")
        data_resp.Recordset("hor_rea") = data_lla.Recordset("hor_rea")
        data_resp.Recordset("fecpas") = data_lla.Recordset("fecpas")
        data_resp.Recordset("horpas") = data_lla.Recordset("horpas")
        data_resp.Recordset("fecsali") = data_lla.Recordset("fecsali")
        data_resp.Recordset("horsali") = data_lla.Recordset("horsali")
        data_resp.Recordset("fec_llega") = data_lla.Recordset("fec_llega")
        data_resp.Recordset("hor_llega") = data_lla.Recordset("hor_llega")
        data_resp.Recordset("fec_rea") = data_lla.Recordset("fec_rea")
        data_resp.Recordset("hor_rea") = data_lla.Recordset("hor_rea")
        data_resp.Recordset("diag") = data_lla.Recordset("diag")
        data_resp.Recordset("colormot") = data_lla.Recordset("colormot")
        data_resp.Recordset("codmed") = data_lla.Recordset("codmed")
        data_resp.Recordset("obs") = data_lla.Recordset("obs")
        data_resp.Recordset("nommed") = data_lla.Recordset("nommed")
        data_resp.Recordset("trasla") = data_lla.Recordset("trasla")
        data_resp.Recordset("lugar") = data_lla.Recordset("lugar")
        data_resp.Recordset("hsald") = data_lla.Recordset("hsald")
        data_resp.Recordset("hllega") = data_lla.Recordset("hllega")
        data_resp.Recordset("hzona") = data_lla.Recordset("hzona")
        data_resp.Recordset("movil_rea") = data_lla.Recordset("movil_rea")
        data_resp.Recordset("totdem") = data_lla.Recordset("totdem")
        data_resp.Recordset("totend") = data_lla.Recordset("totend")
        data_resp.Recordset("pasado") = data_lla.Recordset("pasado")
        data_resp.Recordset("realiza") = data_lla.Recordset("realiza")
        data_resp.Recordset("motmov") = data_lla.Recordset("motmov")
        data_resp.Recordset("cancela") = data_lla.Recordset("cancela")
        data_resp.Recordset("fec_cance") = data_lla.Recordset("fec_cance")
        data_resp.Recordset("hor_cance") = data_lla.Recordset("hor_cance")
        data_resp.Recordset("motcance") = data_lla.Recordset("motcance")
        data_resp.Recordset("movtras") = data_lla.Recordset("movtras")
        data_resp.Recordset("hh") = data_lla.Recordset("hh")
        data_resp.Recordset("mm") = data_lla.Recordset("mm")
        data_resp.Recordset("thh") = data_lla.Recordset("thh")
        data_resp.Recordset("tmm") = data_lla.Recordset("tmm")
        data_resp.Recordset("enfer") = data_lla.Recordset("enfer")
        data_resp.Recordset("activo") = data_lla.Recordset("activo")
        data_resp.Recordset("timsi") = data_lla.Recordset("timsi")
        data_resp.Recordset("timdes") = data_lla.Recordset("timdes")
        data_resp.Recordset("ncobr") = data_lla.Recordset("ncobr")
        data_resp.Recordset("dcobr") = data_lla.Recordset("dcobr")
        data_resp.Recordset.Update
        data_lla.Recordset.MoveNext
      Loop
      data_lla.Recordset.MoveFirst
      Do While Not data_lla.Recordset.EOF
         data_lla.Recordset.Delete
         data_lla.Recordset.MoveNext
      Loop
      MsgBox "Proceso terminado..."
   End If
End If
frm_resplla.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lla.RecordSource = "llamado"
data_lla.Refresh
data_resp.DatabaseName = App.Path & "\resplla.mdb"
data_resp.RecordSource = "resplla"
data_resp.Refresh

End Sub
