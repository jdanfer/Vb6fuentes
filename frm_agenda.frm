VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_agenda 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "frm_agenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_agenda.frx":0442
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "frm_agenda.frx":0456
      TabIndex        =   12
      Top             =   4800
      Width           =   6735
   End
   Begin VB.CommandButton bbusca 
      Height          =   615
      Left            =   5040
      Picture         =   "frm_agenda.frx":0FF1
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton bcance 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3840
      Picture         =   "frm_agenda.frx":1433
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton bmodi 
      Height          =   615
      Left            =   2640
      Picture         =   "frm_agenda.frx":1875
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton bgraba 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1440
      Picture         =   "frm_agenda.frx":1CB7
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton balta 
      Height          =   615
      Left            =   240
      Picture         =   "frm_agenda.frx":20F9
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Datos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt_t 
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
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txt_d 
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
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txt_n 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   35
         TabIndex        =   2
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Teléfono/s:"
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
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Dirección:"
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
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Nombre:"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub balta_Click()
Frame1.Enabled = True
txt_n.SetFocus
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
   Text1.Text = Data1.Recordset("cl_codigo") + 1
Else
   Text1.Text = 1000
End If
Data1.Recordset.AddNew
balta.Enabled = False
bgraba.Enabled = True
bmodi.Enabled = False
bcance.Enabled = True
bbusca.Enabled = False
XAlta = 1

End Sub

Private Sub bcance_Click()
If XAlta = 1 Then
   Data1.Recordset.CancelUpdate
   balta.Enabled = True
   bgraba.Enabled = False
   bmodi.Enabled = True
   bcance.Enabled = False
   bbusca.Enabled = True
   XAlta = 0
   borraagen
   Frame1.Enabled = False
Else
   balta.Enabled = True
   bgraba.Enabled = False
   bmodi.Enabled = True
   bcance.Enabled = False
   bbusca.Enabled = True
   XAlta = 0
   borraagen
   Frame1.Enabled = False
End If
   
End Sub

Private Sub bgraba_Click()
If XAlta = 1 Then
   Data1.Recordset("cl_codigo") = Text1.Text
   Data1.Recordset("cl_apellid") = txt_n.Text
   Data1.Recordset("cl_nombre") = txt_t.Text
   Data1.Recordset("cl_direcci") = txt_d.Text
   Data1.Recordset("cl_etiquet") = 1
   Data1.Recordset.Update
   Data1.Refresh
   borraagen
   igualaagen
   XAlta = 0
   balta.Enabled = True
   bgraba.Enabled = False
   bmodi.Enabled = True
   bcance.Enabled = False
   bbusca.Enabled = True
   Frame1.Enabled = False
Else
   Data1.Recordset.Edit
   Data1.Recordset("cl_codigo") = Text1.Text
   Data1.Recordset("cl_apellid") = txt_n.Text
   Data1.Recordset("cl_nombre") = txt_t.Text
   Data1.Recordset("cl_direcci") = txt_d.Text
   Data1.Recordset("cl_etiquet") = 1
   Data1.Recordset.Update
   Data1.Refresh
   XAlta = 0
   borraagen
   igualaagen
   balta.Enabled = True
   bgraba.Enabled = False
   bmodi.Enabled = True
   bcance.Enabled = False
   bbusca.Enabled = True
   Frame1.Enabled = False
End If

End Sub

Private Sub bmodi_Click()
Frame1.Enabled = True
XAlta = 0
txt_n.SetFocus
balta.Enabled = False
bgraba.Enabled = True
bmodi.Enabled = False
bcance.Enabled = True
bbusca.Enabled = False

End Sub

Private Sub DBGrid1_DblClick()
borraagen
igualaagen

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.RecordSource = "select * from env_soc where cl_codigo <" & 10000 & " order by cl_codigo"
Data1.Refresh

End Sub

Private Sub txt_d_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_t.SetFocus
End If

End Sub

Private Sub txt_n_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_d.SetFocus
End If

End Sub

Private Sub txt_t_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Public Function borraagen()
txt_n.Text = ""
txt_d.Text = ""
txt_t.Text = ""

End Function

Public Function igualaagen()
txt_n.Text = Data1.Recordset("cl_apellid")
txt_d.Text = Data1.Recordset("cl_direcci")
txt_t.Text = Data1.Recordset("cl_nombre")
Text1.Text = Data1.Recordset("cl_codigo")

End Function
