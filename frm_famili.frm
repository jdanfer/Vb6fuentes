VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_famili 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familias"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5820
   Icon            =   "frm_famili.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_famili.frx":058A
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "frm_famili.frx":059E
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frm_famili"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "familias"
Data1.Refresh

End Sub
