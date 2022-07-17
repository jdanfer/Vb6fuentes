VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_busper 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar personal"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7635
   Icon            =   "frm_busper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7635
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton B_C 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      Picture         =   "frm_busper.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      RecordSource    =   "Select * from tarjbrou order by ape1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_busper.frx":09CC
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "frm_busper.frx":09E0
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   3480
      Picture         =   "frm_busper.frx":1717
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "frm_busper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub B_C_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_personal.data_per.Recordset.FindFirst "cedula =" & Data1.Recordset("cedula")
If Not frm_personal.data_per.Recordset.NoMatch Then
   frm_personal.txt_ced.Text = Data1.Recordset("cedula")
   frm_personal.txt_cod.Text = Data1.Recordset("codver")
   frm_personal.txt_nom1.Text = Data1.Recordset("nom1")
   If IsNull(Data1.Recordset("nom2")) = False Then
      frm_personal.txt_nom2.Text = Data1.Recordset("nom2")
   Else
      frm_personal.txt_nom2.Text = ""
   End If
   frm_personal.txt_ape1.Text = Data1.Recordset("ape1")
   If IsNull(Data1.Recordset("ape2")) = False Then
      frm_personal.txt_ape2.Text = Data1.Recordset("ape2")
   Else
      frm_personal.txt_ape2.Text = ""
   End If
   If IsNull(Data1.Recordset("calle")) = False Then
      frm_personal.txt_dir.Text = Data1.Recordset("calle")
   Else
      frm_personal.txt_dir.Text = ""
   End If
   If IsNull(Data1.Recordset("codpos")) = False Then
      frm_personal.txt_codp.Text = Data1.Recordset("codpos")
   Else
      frm_personal.txt_codp.Text = 0
   End If
   If IsNull(Data1.Recordset("telpart")) = False Then
      frm_personal.txt_tel.Text = Data1.Recordset("telpart")
   Else
      frm_personal.txt_tel.Text = ""
   End If
   If IsNull(Data1.Recordset("localid")) = False Then
      frm_personal.txt_loc.Text = Data1.Recordset("localid")
   Else
      frm_personal.txt_loc.Text = ""
   End If
   If IsNull(Data1.Recordset("depto")) = False Then
      frm_personal.txt_dpto.Text = Data1.Recordset("depto")
   Else
      frm_personal.txt_dpto.Text = ""
   End If
   If IsNull(Data1.Recordset("fecing")) = False Then
      frm_personal.mfing.Text = Format(Data1.Recordset("fecing"), "dd/mm/yyyy")
   Else
      frm_personal.mfing.Text = "__/__/____"
   End If
   If IsNull(Data1.Recordset("rellab")) = False Then
      frm_personal.txt_rlab.Text = Data1.Recordset("rellab")
   Else
      frm_personal.txt_rlab.Text = 3
   End If
   If IsNull(Data1.Recordset("forpag")) = False Then
      frm_personal.txt_fpag.Text = Data1.Recordset("forpag")
   Else
      frm_personal.txt_fpag.Text = 1
   End If
   
   Unload Me
Else
   MsgBox "Verifique datos", vbCritical, "Mensaje"
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
