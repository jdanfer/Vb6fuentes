VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_bussue 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar sueldos ingresados"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "frm_bussue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Brou"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
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
      Left            =   6240
      Picture         =   "frm_bussue.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   3360
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_bussue.frx":09CC
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "frm_bussue.frx":09E0
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   3960
      Picture         =   "frm_bussue.frx":1BE7
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frm_bussue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_sueldos.Data1.Recordset.FindFirst "fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and ced =" & Data1.Recordset("ced")
If Not frm_sueldos.Data1.Recordset.NoMatch Then
   frm_sueldos.Label5.Caption = frm_sueldos.Data1.Recordset("nom1") + " " + frm_sueldos.Data1.Recordset("ape1")
   frm_sueldos.txt_ced.Text = frm_sueldos.Data1.Recordset("ced")
   frm_sueldos.txt_cod.Text = frm_sueldos.Data1.Recordset("codver")
   frm_sueldos.timp.Text = frm_sueldos.Data1.Recordset("importe")
   Unload Me
'   frm_sueldos.txt_ced.SetFocus
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from brou where fecha =#" & Format(frm_sueldos.mfec.Text, "yyyy/mm/dd") & "# order by ced"
Data1.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
