VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_verpres 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos en los registros"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8595
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8595
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frm_verpres.frx":0000
      Height          =   1455
      Left            =   120
      OleObjectBlob   =   "frm_verpres.frx":0014
      TabIndex        =   3
      Top             =   2640
      Width           =   8295
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_verpres.frx":0D2B
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frm_verpres.frx":0D3F
      TabIndex        =   1
      Top             =   240
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   4200
      Width           =   855
   End
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
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Procesos de archivos"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3615
   End
End
Attribute VB_Name = "frm_verpres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\abmpres.mdb"
Data1.RecordSource = "Select * from prestamo where cedula =" & frm_prestamo.tced.Text
Data1.Refresh
Data2.DatabaseName = App.Path & "\abmpres.mdb"
Data2.RecordSource = "Select * from prestamo where cedula =" & 0
Data2.Refresh

End Sub
