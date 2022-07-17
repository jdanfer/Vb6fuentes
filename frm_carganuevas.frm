VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_carganuevas 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar nuevas entregas"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frm_carganuevas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_arq 
      Caption         =   "data_arq"
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
      Top             =   2640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Terminar"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar..."
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
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para cargar nuevas entregas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6615
      Begin VB.TextBox txt_a 
         Alignment       =   1  'Right Justify
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
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txt_m 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label2 
         Caption         =   "Mes/Año de emisión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "FECHA DESDE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frm_carganuevas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Nomem As String
frm_carganuevas.MousePointer = 11
Nomem = "EMI"
If md.Text <> "__/__/____" Then
   If txt_m.Text <> "" Then
      If txt_m.Text < 10 Then
         Nomem = Nomem + "0" + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
      Else
         Nomem = Nomem + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
      End If
      data_emi.RecordSource = "Select * from " & Nomem & " where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "#"
      data_emi.Refresh
      If data_emi.Recordset.RecordCount > 0 Then
         data_emi.Recordset.MoveFirst
         Do While Not data_emi.Recordset.EOF
            data_arq.Recordset.AddNew
            data_arq.Recordset("matricula") = data_emi.Recordset("cliente")
            data_arq.Recordset("nombre") = Mid(data_emi.Recordset("apellidos"), 1, 20)
            data_arq.Recordset("mes") = data_emi.Recordset("mes")
            data_arq.Recordset("ano") = data_emi.Recordset("ano")
            data_arq.Recordset("color") = data_emi.Recordset("color_rec")
            data_arq.Recordset("cat") = data_emi.Recordset("cod_cnv")
            data_arq.Recordset("nomcat") = Mid(data_emi.Recordset("nom_cnv"), 1, 15)
            data_arq.Recordset("arqueo") = "C"
            data_arq.Recordset("importe") = data_emi.Recordset("importe")
            data_arq.Recordset("fecha") = Date
            data_arq.Recordset("nrorec") = data_emi.Recordset("documento")
            data_arq.Recordset("usuar") = frm_menu.labusua.Caption
            data_arq.Recordset("moneda") = data_emi.Recordset("moneda")
            data_arq.Recordset("cob") = data_emi.Recordset("nro_cobr")
            data_arq.Recordset("nomcob") = Mid(data_emi.Recordset("nom_cobr"), 1, 15)
            data_arq.Recordset("codzon") = data_emi.Recordset("grupo")
            data_arq.Recordset("codsup") = data_emi.Recordset("nro_superv")
            data_arq.Recordset("codpro") = data_emi.Recordset("nro_vende")
            data_arq.Recordset("tiquet") = data_emi.Recordset("tiquet")
            data_arq.Recordset("total") = data_emi.Recordset("total")
            data_arq.Recordset("varia") = 0
            data_arq.Recordset.Update
            data_emi.Recordset.MoveNext
         Loop
         MsgBox "Proceso terminado"
      Else
         MsgBox "No existen registros"
      End If
   End If
End If
frm_carganuevas.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_arq.DatabaseName = App.Path & "\sapp.mdb"
data_arq.RecordSource = "arqueo"
data_arq.Refresh
data_emi.DatabaseName = App.Path & "\emisiones.mdb"

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_m.SetFocus
End If

End Sub

Private Sub txt_a_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub txt_m_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_a.SetFocus
End If

End Sub
