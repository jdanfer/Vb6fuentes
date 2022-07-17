VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_opscob 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Opciones de cobrador"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF8080&
      Caption         =   "Seleccione cobrador:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_opscob.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   2280
      Width           =   495
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "frm_opscob.frx":058A
      Height          =   360
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      ListField       =   "CB_NOMBRE"
      Text            =   ""
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "TODOS los cobradores"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Opciones para selección de cobrador"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2520
      Picture         =   "frm_opscob.frx":059E
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "frm_opscob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.value = 1 Then
   Check2.value = 0
   DBCombo1.Enabled = False
End If

End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
   Check1.value = 0
   DBCombo1.Enabled = True
   DBCombo1.SetFocus
End If

End Sub

Private Sub Command1_Click()
If Check2.value = 1 Then
   If DBCombo1.Text <> "" Then
      Data1.Recordset.FindFirst "cb_nombre ='" & DBCombo1.Text & "'"
      If Not Data1.Recordset.NoMatch Then
         Wopscob = Data1.Recordset("cb_numero")
         Wopscobd = Data1.Recordset("cb_nombre")
         Unload Me
      Else
         MsgBox "Cobrador no encontrado", vbInformation, "Cobradores"
         DBCombo1.SetFocus
      End If
   Else
      MsgBox "Seleccione cobrador", vbInformation, "Cobradores"
      DBCombo1.SetFocus
   End If
Else
   Wopscob = 0
   Wopscobd = ""
   Unload Me
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select * from cobrador order by cb_nombre"
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
