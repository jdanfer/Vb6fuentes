VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_evaluacom 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluación de proveedor"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_evaluacom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9780
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "evalua"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_evaluacom.frx":0442
      Height          =   1695
      Left            =   240
      OleObjectBlob   =   "frm_evaluacom.frx":0456
      TabIndex        =   12
      Top             =   3240
      Width           =   9255
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_evaluacom.frx":0E29
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancelar acciòn"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_evaluacom.frx":13B3
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Grabar registro"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_evaluacom.frx":193D
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Editar registro"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_evaluacom.frx":1EC7
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Nuevo registro"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de la evaluación"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9255
      Begin VB.TextBox t_obs 
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1200
         Width           =   6615
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "Observación:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "FECHA:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label labcomd 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label labcomc 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Comercio:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   4080
      Picture         =   "frm_evaluacom.frx":2451
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frm_evaluacom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()
XAlta = 1
b_alta.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
Data1.RecordSource = "Select * from abmdesp order by nro"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Text1.Text = Data1.Recordset("nro") + 1
Else
   Text1.Text = 1
End If
Data1.RecordSource = "Select * from abmdesp where base =" & 99 & " and mat =" & labcomc.Caption
Data1.Refresh

mfec.Text = "__/__/____"
t_obs.Text = ""
mfec.SetFocus

End Sub

Private Sub b_cance_Click()
      mfec.Text = "__/__/____"
      t_obs.Text = ""
      b_alta.Enabled = True
      b_modif.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False

End Sub

Private Sub b_graba_Click()
If mfec.Text <> "__/__/____" Then
   If XAlta = 1 Then
      Data1.Recordset.AddNew
      Data1.Recordset("base") = 99
      Data1.Recordset("nro") = Text1.Text
      Data1.Recordset("mat") = labcomc.Caption
      Data1.Recordset("fecha") = mfec.Text
      Data1.Recordset("referen") = t_obs.Text
      Data1.Recordset.Update
      Data1.Refresh
      mfec.Text = "__/__/____"
      t_obs.Text = ""
      b_alta.Enabled = True
      b_modif.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      DBGrid1.SetFocus
   Else
      Data1.Recordset.Edit
      Data1.Recordset("nro") = Text1.Text
      Data1.Recordset("mat") = labcomc.Caption
      Data1.Recordset("fecha") = mfec.Text
      Data1.Recordset("referen") = t_obs.Text
      Data1.Recordset.Update
      Data1.Refresh
      mfec.Text = "__/__/____"
      t_obs.Text = ""
      b_alta.Enabled = True
      b_modif.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      DBGrid1.SetFocus
   End If
Else
   MsgBox "Ingrese fecha"
End If

End Sub

Private Sub b_modif_Click()
XAlta = 0
b_alta.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
t_obs.SetFocus

End Sub

Private Sub DBGrid1_DblClick()

If IsNull(Data1.Recordset("fecha")) = False Then
   mfec.Text = Data1.Recordset("fecha")
Else
   mfec.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("referen")) = False Then
   t_obs.Text = Data1.Recordset("referen")
Else
   t_obs.Text = ""
End If

End Sub

Private Sub Form_Load()
'Data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

If frm_labo.t_cod.Text <> "" Then
    labcomc.Caption = frm_labo.t_cod.Text
    labcomd.Caption = frm_labo.t_nom.Text
    Data1.RecordSource = "Select * from abmdesp where base =" & 99 & " and mat =" & labcomc.Caption
    Data1.Refresh
End If

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
