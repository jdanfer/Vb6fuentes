VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_opspro 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Seleccion de promotor"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
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
   ScaleHeight     =   3015
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_opspro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aceptar"
      Top             =   2280
      Width           =   615
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "frm_opspro.frx":058A
      Height          =   360
      Left            =   3600
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      ListField       =   "VN_NOMBRE"
      Text            =   ""
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Seleccione Promotor:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TODOS los promotores"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   8280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   8280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Selección datos de promotor"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm_opspro.frx":059E
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "frm_opspro"
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
If Check1.value = 1 Then
   Wopspro = 0
   Wopsprod = ""
   Unload Me
Else
   If Check2.value = 1 Then
      If DBCombo1.Text <> "" Then
         Data1.Recordset.FindFirst "vn_nombre ='" & DBCombo1.Text & "'"
         If Not Data1.Recordset.NoMatch Then
            Wopspro = Data1.Recordset("vn_numero")
            Wopsprod = Data1.Recordset("vn_nombre")
            Unload Me
         Else
            MsgBox "Promotor no encontrado", vbInformation, "Promotores"
            DBCombo1.SetFocus
         End If
      Else
         MsgBox "Seleccione Promotor", vbInformation, "Promotores"
         DBCombo1.SetFocus
      End If
   Else
      Wopspro = 0
      Wopsprod = ""
      Unload Me
   End If
End If
      
End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from vendedor order by vn_nombre"
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
