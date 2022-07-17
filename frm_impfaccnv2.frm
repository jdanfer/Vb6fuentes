VERSION 5.00
Begin VB.Form frm_impfaccnv2 
   BorderStyle     =   0  'None
   Caption         =   "Impresión"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5400
      Picture         =   "frm_impfaccnv2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   375
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox t_desde 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nro. de Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   480
      Picture         =   "frm_impfaccnv2.frx":058A
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "frm_impfaccnv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
data_inf.RecordSource = "lineas2"
data_inf.Refresh

If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If
If t_desde.Text <> "" Then

Else
   MsgBox "No ingresó número de documento a imprimir"
End If


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

t_desde.Text = frm_factconve22.labnrofac.Caption

data_inf.DatabaseName = App.Path & "\factura.mdb"

End Sub

Private Sub Form_Resize()
With Image1
     .Top = 0
     .Left = 0
     .Height = Me.Height
     .Width = Me.Width
     
End With
End Sub
