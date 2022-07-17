VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_opsconv 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Opciones de convenios"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check7 
      BackColor       =   &H00FF0000&
      Caption         =   "Sin complementos"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2700
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "frm_opsconv.frx":0000
      Height          =   600
      Left            =   3000
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1058
      _Version        =   393216
      Style           =   1
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cbogposap 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_opsconv.frx":0014
      Left            =   3240
      List            =   "frm_opsconv.frx":0027
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00FF0000&
      Caption         =   "SELECCIONAR CONVENIO"
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
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FF0000&
      Caption         =   "GRUPOS SAPP"
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
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FF0000&
      Caption         =   "SIN COMPLEMENTOS"
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
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FF0000&
      Caption         =   "SOLO COMPLEMENTOS"
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
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ComboBox cbomutcnv 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_opsconv.frx":005E
      Left            =   3240
      List            =   "frm_opsconv.frx":0080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF0000&
      Caption         =   "GRUPO MUTUAL"
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
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "TODOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "frm_opsconv.frx":00E1
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar"
      Top             =   4800
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      X1              =   0
      X2              =   7800
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      X1              =   0
      X2              =   7800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      X1              =   0
      X2              =   7800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Opciones para convenios"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   4560
      Picture         =   "frm_opsconv.frx":066B
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2655
   End
End
Attribute VB_Name = "frm_opsconv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.value = 1 Then
   Check2.value = 0
   Check3.value = 0
   Check4.value = 0
   Check5.value = 0
   Check6.value = 0
   Check7.Enabled = True
   Check7.value = 0
   Check7.Enabled = False
End If

End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
   Check1.value = 0
   Check3.value = 0
   Check4.value = 0
   Check5.value = 0
   Check6.value = 0
   Check7.Enabled = True
   Check7.value = 0
   Check7.Enabled = False

''   cbomutcnv.SetFocus
''   cbomutcnv.ListIndex = 0
End If

End Sub

Private Sub Check3_Click()
If Check3.value = 1 Then
   Check4.value = 0
End If

End Sub

Private Sub Check4_Click()
If Check4.value = 1 Then
   Check3.value = 0
End If

End Sub

Private Sub Check5_Click()
If Check5.value = 1 Then
   Check1.value = 0
   Check3.value = 0
   Check4.value = 0
   Check2.value = 0
   Check6.value = 0
   Check7.Enabled = True
'   cbogposap.SetFocus
'   cbogposap.ListIndex = 0
End If

End Sub

Private Sub Check6_Click()
If Check6.value = 1 Then
   Check1.value = 0
   Check3.value = 0
   Check4.value = 0
   Check5.value = 0
   Check2.value = 0
   DBCombo1.Visible = True
'   DBCombo1.SetFocus
   Check7.Enabled = True
   Check7.value = 0
   Check7.Enabled = False

End If

End Sub

Private Sub Command1_Click()
If Check1.value = 1 Then 'Todos
   Wopsconv = 1
   Wopsconvd = ""
Else
   If Check2.value = 1 Then 'Conv.Mutual
      Wopsconv = 2
      Wopsconvd = cbomutcnv.Text
      If Check3.value = 1 Then 'solo complementos
         Wopsconv = 5
      End If
      If Check4.value = 1 Then 'sin complemento
         Wopsconv = 6
      End If
   Else
      If Check5.value = 1 Then 'Convs.SAPP
         Wopsconv = 3
         Wopsconvd = cbogposap.Text
         If Check7.value = 1 Then
            Wopsconv = 9
         End If
      Else
         If Check6.value = 1 Then 'Selección
            Wopsconv = 4
            Wopsconvd = Data1.Recordset("cnv_codigo")
         Else
            Wopsconv = 1
            Wopsconvd = ""
         End If
      End If
   End If
End If
If cbomutcnv.Text <> "" Then
   If cbomutcnv.Text = "CCOU" Then
      Xledes = "C"
      Xlehas = "D"
   End If
   If cbomutcnv.Text = "H.EVANGELICO" Then
      Xledes = "E"
      Xlehas = "H"
   End If
   If cbomutcnv.Text = "UNIVERSAL" Then
      Xledes = "U"
      Xlehas = "X"
   End If
   If cbomutcnv.Text = "SMI" Then
      Xledes = "SM"
      Xlehas = "SX"
   End If
   If cbomutcnv.Text = "IMPASA" Then
      Xledes = "I"
      Xlehas = "J"
   End If
   If cbomutcnv.Text = "CASA DE GALICIA" Then
      Xledes = "C"
      Xlehas = "D"
   End If
   If cbomutcnv.Text = "CPS" Then
      Xledes = "CPS"
      Xlehas = "CPSSA"
   End If
   If cbomutcnv.Text = "SEMM" Then
      Xledes = "SEMM1"
      Xlehas = "SEMM2"
   End If
   If cbomutcnv.Text = "RET.MILITARES" Then
      Xledes = "RETMI"
      Xlehas = "RETMIL"
   End If
   
Else
   Xledes = ""
   Xlehas = ""
End If

Unload Me

End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBCombo1.ListField = "cnv_desc"
   DBCombo1.BoundColumn = "cnv_desc"
   If DBCombo1.Text <> "" Then
      Data1.Recordset.FindFirst "cnv_desc ='" & UCase(DBCombo1.Text) & "'"
      If Not Data1.Recordset.NoMatch Then
         DBCombo1.Text = Data1.Recordset("cnv_desc")
         DBCombo1.Height = 600
         DBCombo1.ListField = ""
         DBCombo1.BoundColumn = ""
         Command1.SetFocus
      Else
         DBCombo1.Height = 1650
         Data1.RecordSource = "Select * from convenio where cnv_desc >='" & DBCombo1.Text & "' order by cnv_desc"
         Data1.Refresh
      End If
   Else
      DBCombo1.Height = 1650
      Data1.RecordSource = "Select * from convenio order by cnv_desc"
      Data1.Refresh
   End If
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from convenio order by cnv_desc"
Data1.Refresh
If Wopsconv = 1 Then
   Check1.value = 1
End If
If Wopsconv = 2 Then
   Check2.value = 1
End If
If Wopsconv = 5 Then
   Check2.value = 1
   Check3.value = 1
End If
If Wopsconv = 6 Then
   Check2.value = 1
   Check4.value = 1
End If
If Wopsconv = 3 Then
   Check5.value = 1
End If
If Wopsconv = 4 Then
   Check6.value = 1
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
