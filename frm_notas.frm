VERSION 5.00
Begin VB.Form frm_notas 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas "
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_notas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2880
      Width           =   8055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton bcierra 
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
      Left            =   7680
      Picture         =   "frm_notas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton bc 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_notas.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancelar acción"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton bg 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_notas.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Grabar"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton bn 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_notas.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo registro"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Width           =   8055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Nueva anotación:"
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
      TabIndex        =   9
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
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
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
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
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Socio:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   5160
      Picture         =   "frm_notas.frx":1A6A
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "frm_notas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bc_Click()
Text1.Enabled = True
Text2.Enabled = False
Text2.Text = ""
bn.Enabled = True
bg.Enabled = False
bc.Enabled = False
bcierra.Enabled = True

End Sub

Private Sub bcierra_Click()
Unload Me

End Sub

Private Sub bg_Click()
If Text2.Text <> "" Then
   Data1.Recordset.AddNew
   Data1.Recordset("cl_codigo") = Label2.Caption
   Data1.Recordset("fecha") = Date
   Data1.Recordset("anota") = Text2.Text
   Data1.Recordset.Update
   Data1.RecordSource = "Select * from tanota where cl_codigo =" & Label2.Caption & " order by fecha"
   Data1.Refresh
   Text1.Text = ""
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         Text1.Text = Text1.Text + "=======>" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
         Text1.Text = Text1.Text & " --" & Data1.Recordset("anota")
         Data1.Recordset.MoveNext
      Loop
   Else
      Text1.Text = ""
   End If
   Text1.Enabled = True
   Text2.Enabled = False
   Text2.Text = ""
   bn.Enabled = True
   bg.Enabled = False
   bc.Enabled = False
   bcierra.Enabled = True
Else
   MsgBox "Verifique el texto o cancele la operación", vbCritical, "Mensaje"
   Text2.SetFocus
End If
End Sub

Private Sub bn_Click()
Text1.Enabled = False
Text2.Enabled = True
Text2.SetFocus
bn.Enabled = False
bg.Enabled = True
bc.Enabled = True
bcierra.Enabled = False

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\anota.mdb"
Data1.RecordSource = "tanota"
Data1.Refresh
Label2.Caption = frmabm.txt_mat.Caption
Label3.Caption = frmabm.txt_apellid.Text
Data1.RecordSource = "Select * from tanota where cl_codigo =" & Label2.Caption & " order by fecha"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Text1.Text = Text1.Text + "======> " & Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
      Text1.Text = Text1.Text & " --" & Data1.Recordset("anota")
      Data1.Recordset.MoveNext
   Loop
Else
   Text1.Text = ""
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

