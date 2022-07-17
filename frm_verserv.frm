VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_verserv 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicios"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "frm_verserv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7485
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "estudios"
      Top             =   2040
      Visible         =   0   'False
      Width           =   3135
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
      Left            =   6480
      Picture         =   "frm_verserv.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3000
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_verserv.frx":09CC
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "frm_verserv.frx":09E0
      TabIndex        =   2
      Top             =   480
      Width           =   6975
   End
   Begin VB.TextBox txt_c 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Con doble click SELECCIONA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Consulta por descripción:"
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
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4320
      Picture         =   "frm_verserv.frx":13C7
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "frm_verserv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_vtasxgpo.t_cod.Text = Data1.Recordset("codest")
Unload Me

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
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

Private Sub txt_c_KeyPress(KeyAscii As Integer)
Data1.RecordSource = "select * from estudios where descrip >='" & UCase(txt_c.Text) & "' order by descrip"
Data1.Refresh

End Sub
