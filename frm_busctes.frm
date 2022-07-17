VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_busctes 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Buscar datos en caja tesorería"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "frm_busctes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Buscar"
      Top             =   4080
      Width           =   615
   End
   Begin VB.Data data_teso 
      Caption         =   "data_teso"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Picture         =   "frm_busctes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_busctes.frx":0B14
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "frm_busctes.frx":0B2C
      TabIndex        =   5
      Top             =   960
      Width           =   7215
   End
   Begin VB.TextBox txt_nrorub 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin MSMask.MaskEdBox mfecha 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "POR RUBRO..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "POR FECHA..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "BUSCAR POR:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   4560
      Picture         =   "frm_busctes.frx":1BA3
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frm_busctes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_busctes.Hide

End Sub

Private Sub Command2_Click()
If Option1.value = True Then
   If IsDate(mfecha.Text) = True Then
      data_teso.RecordSource = "select * from tesorero where fecha >=#" & Format(mfecha.Text, "yyyy/mm/dd") & "# and usuario =" & WNombase
      data_teso.Refresh
   Else
      MsgBox "Error en la fecha de búsqueda", vbCritical, "Mensaje"
      mfecha.SetFocus
   End If
End If
If Option2.value = True Then
   If IsNumeric(txt_nrorub.Text) = True Then
      data_teso.RecordSource = "select * from tesorero where cod_rub >=" & txt_nrorub.Text & " and usuario =" & WNombase
      data_teso.Refresh
   Else
      MsgBox "Error en el ingreso de RUBRO", vbCritical, "Mensaje"
      txt_nrorub.SetFocus
   End If
End If

End Sub

Private Sub DBGrid1_DblClick()
'    frm_teso.data_cajtes.Recordset.FindFirst "nromov =" & data_teso.Recordset("nromov")
    frm_teso.data_cajtes.RecordSource = "Select * from tesorero where nromov =" & data_teso.Recordset("nromov")
    frm_teso.data_cajtes.Refresh
    If frm_teso.data_cajtes.Recordset.RecordCount > 0 Then
       frm_teso.txt_base.Text = data_teso.Recordset("usuario")
       frm_teso.mfecha.Text = Format(data_teso.Recordset("fecha"), "dd/mm/yyyy")
       frm_teso.txt_hora.Text = Format(data_teso.Recordset("hora"), "HH:mm")
       frm_teso.txt_codrub.Text = data_teso.Recordset("cod_rub")
       frm_teso.dbcborub.Text = data_teso.Recordset("nom_rub")
       frm_teso.txt_debe.Text = data_teso.Recordset("cod_debe")
       frm_teso.txt_haber.Text = data_teso.Recordset("cod_haber")
       If data_teso.Recordset("moneda") = 2 Then
          frm_teso.cbomon.ListIndex = 1
          frm_teso.txt_imp2.Visible = True
          frm_teso.txt_tcam.Visible = True
          frm_teso.txt_imp2.Text = Format(data_teso.Recordset("saldou"), "Standard")
          frm_teso.txt_tcam.Text = Format(data_teso.Recordset("tcam"), "Standard")
       Else
          frm_teso.txt_imp2.Visible = False
          frm_teso.txt_tcam.Visible = False
          frm_teso.cbomon.ListIndex = 0
       End If
       frm_teso.cboiva.ListIndex = data_teso.Recordset("iva")
       frm_teso.txt_impiva.Text = Format(data_teso.Recordset("impiva"), "Standard")
       frm_teso.txt_con.Text = data_teso.Recordset("concep")
       frm_teso.txt_imp.Text = Format(data_teso.Recordset("monto"), "Standard")
       If IsNull(data_teso.Recordset("obs")) = False Then
          frm_teso.txt_obs.Text = data_teso.Recordset("obs")
       Else
          frm_teso.txt_obs.Text = ""
       End If
       frm_teso.labsaldop.Caption = Format(data_teso.Recordset("saldos"), "Standard")
       
       frm_busctes.Hide
    Else
       MsgBox "Error en la búsqueda", vbCritical, "Buscar"
       DBGrid1.SetFocus
    End If

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If


End Sub

Private Sub Form_Initialize()
'data_teso.Refresh
'data_teso.Recordset.MoveLast

End Sub

Private Sub Form_Load()
data_teso.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_teso.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' order by nromov DESC"

'data_teso.RecordSource = "tesorero"
data_teso.Refresh
'data_teso.Recordset.MoveLast

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command2.SetFocus
End If

End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfecha.SetFocus
End If

End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nrorub.SetFocus
End If
End Sub

Private Sub txt_nrorub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command2.SetFocus
End If

End Sub
