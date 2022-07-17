VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_histo 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial "
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frm_histo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btn_cerrar 
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
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFF00&
      Picture         =   "frm_histo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Data data_histo 
      Caption         =   "data_histo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_histo.frx":09CC
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "frm_histo.frx":09E5
      TabIndex        =   0
      Top             =   360
      Width           =   9855
   End
   Begin VB.Label labnomh 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label labmath 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "MOVIMIENTOS DEL SOCIO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4080
      Picture         =   "frm_histo.frx":1BFC
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frm_histo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


End Sub

Private Sub btn_cerrar_Click()
frm_histo.Hide
End Sub

Private Sub Form_Activate()
data_histo.RecordSource = "Select * from abmsocio where cl_codigo =" & frmabm.txt_mat.Caption
data_histo.Refresh
labmath.Caption = frmabm.txt_mat.Caption
labnomh.Caption = frmabm.txt_apellid.Text
btn_cerrar.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   frm_histo.Hide
End If

End Sub

Private Sub Form_Load()
data_histo.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_histo.RecordSource = "Select * from abmsocio where cl_codigo =" & frmabm.txt_mat.Caption
data_histo.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_histo.Hide

End Sub

