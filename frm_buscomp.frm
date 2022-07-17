VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_buscomp 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar compras"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_buscomp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9795
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   735
      Left            =   8760
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buscomp.frx":0442
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "frm_buscomp.frx":0456
      TabIndex        =   3
      Top             =   600
      Width           =   9495
   End
   Begin VB.TextBox tg_buscom 
      Height          =   360
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Por fecha"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Por descripción"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Doble click selecciona el registro."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   5055
   End
End
Attribute VB_Name = "frm_buscomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_compsto.data_comp.Recordset.FindFirst "grupo =" & Data1.Recordset("grupo")
If Not frm_compsto.data_comp.Recordset.NoMatch Then
    If IsNull(Data1.Recordset("grupo")) = False Then
       frm_compsto.t_nro.Text = Data1.Recordset("grupo")
    Else
       frm_compsto.t_nro.Text = 1
    End If
    If IsNull(Data1.Recordset("fecha")) = False Then
       frm_compsto.mfc.Text = Data1.Recordset("fecha")
    Else
       frm_compsto.mfc.Text = "__/__/____"
    End If
    If IsNull(Data1.Recordset("cliente")) = False Then
       frm_compsto.T_COD.Text = Data1.Recordset("cliente")
    Else
       frm_compsto.T_COD.Text = 0
    End If
    If IsNull(Data1.Recordset("nom_cnv")) = False Then
       frm_compsto.labdes.Caption = Data1.Recordset("nom_cnv")
    Else
       frm_compsto.labdes.Caption = ""
    End If
    If IsNull(Data1.Recordset("importe")) = False Then
       frm_compsto.t_precu.Text = Format(Data1.Recordset("importe"), "Standard")
    Else
       frm_compsto.t_precu.Text = 0
    End If
    If IsNull(Data1.Recordset("nro_superv")) = False Then
       frm_compsto.t_cant.Text = Data1.Recordset("nro_superv")
    Else
       frm_compsto.t_cant.Text = 0
    End If
    If IsNull(Data1.Recordset("moneda")) = False Then
       frm_compsto.cbomon.ListIndex = Data1.Recordset("moneda")
    Else
       frm_compsto.cbomon.ListIndex = 0
    End If
    If IsNull(Data1.Recordset("mes")) = False Then
       frm_compsto.cboiva.ListIndex = Data1.Recordset("mes")
    Else
       frm_compsto.cboiva.ListIndex = 0
    End If
    If IsNull(Data1.Recordset("nom_cobr")) = False Then
       frm_compsto.cbolab.Text = Data1.Recordset("nom_cobr")
    Else
       frm_compsto.cbolab.Text = ""
    End If
    If IsNull(Data1.Recordset("total")) = False Then
       frm_compsto.labtot.Caption = Format(Data1.Recordset("total"), "Standard")
    Else
       frm_compsto.labtot.Caption = 0
    End If
    If IsNull(Data1.Recordset("iva")) = False Then
       frm_compsto.labiva.Caption = Format(Data1.Recordset("iva"), "Standard")
    Else
       frm_compsto.labiva.Caption = 0
    End If
    Unload Me

End If

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\" & Trim(Xlabdd)
Data1.RecordSource = "provdeu"
Data1.Refresh


End Sub

Private Sub tg_buscom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Option1.Value = True Then
      Data1.RecordSource = "Select * from provdeu where nom_cnv >='" & tg_buscom.Text & "' order by nom_cnv"
      Data1.Refresh
   End If
   If Option2.Value = True Then
      Data1.RecordSource = "Select * from provdeu where fecha >=#" & Format(tg_buscom.Text, "yyyy/mm/dd") & "# order by fecha"
      Data1.Refresh
   End If
   DBGrid1.SetFocus
End If
   
End Sub
