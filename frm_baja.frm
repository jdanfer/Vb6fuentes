VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_baja 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Baja de socio"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "frm_baja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data data_abmb 
      Caption         =   "data_abmb"
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
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_mot 
      Caption         =   "data_mot"
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
      RecordSource    =   "SELECT * FROM motivos WHERE MC_NUMERO>=""B01"" AND mc_numero<=""B20"""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton btnacept 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Picture         =   "frm_baja.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin MSDBCtls.DBCombo cbomot 
      Bindings        =   "frm_baja.frx":0884
      Height          =   420
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   741
      _Version        =   393216
      Style           =   2
      ListField       =   "MC_DESC"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox txt_fecbaja 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
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
      BackColor       =   &H00FFFFC0&
      Caption         =   "MOTIVO DE BAJA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FECHA DE BAJA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frm_baja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnacept_Click()

frmabm.data_clientes.Recordset.Edit
frmabm.data_clientes.Recordset("estado") = 2
frmabm.data_clientes.Recordset("fecha_baja") = Format(txt_fecbaja.Text)
frmabm.data_clientes.Recordset.Update
frmabm.labestado.Caption = "BAJA"
frmabm.txt_fecbaj.Text = Format(txt_fecbaja.Text, "dd/mm/yyyy")

data_abmb.Recordset.AddNew
data_abmb.Recordset("cl_codigo") = frmabm.txt_mat.Caption
data_abmb.Recordset("cl_motivo") = cbomot.Text
data_abmb.Recordset("desc") = "BAJA"
data_abmb.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
data_abmb.Recordset("hora") = Format(Time, "HH:mm")
data_abmb.Recordset("usuario") = WElusuario
data_abmb.Recordset("convenio") = frmabm.txt_codcnv.Text
data_abmb.Recordset("base") = frmabm.data_parsec.Recordset("base")
data_abmb.Recordset.Update


'frm_baja.Hide
Unload Me

End Sub

Private Sub cbomot_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btnacept.SetFocus
End If

End Sub

Private Sub Form_Deactivate()

'frm_baja.Hide
'txt_fecbaja.Text = Format(Date, "dd/mm/yyyy")

End Sub

Private Sub Form_Load()
data_mot.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_mot.RecordSource = "MOTIVOS"
data_mot.Refresh
data_abmb.Connect = "odbc;dsn=" & Xconexrmt & ";"
'SelectLimit 10
data_abmb.RecordSource = "select * from abmsocio where cl_codigo =" & frmabm.txt_mat.Caption
data_abmb.Refresh
'SelectLimit 0
txt_fecbaja.Text = Format(Date, "dd/mm/yyyy")
cbomot.Text = "PROBLEMAS ECONOMICOS"


End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_baja.Hide

End Sub

Private Sub txt_fecbaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomot.SetFocus
End If

End Sub
