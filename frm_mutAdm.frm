VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_mutAdm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mutualistas (Adm)"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6330
   Icon            =   "frm_mutAdm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_mutAdm.frx":058A
      Height          =   2655
      Left            =   240
      OleObjectBlob   =   "frm_mutAdm.frx":059E
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2520
      Picture         =   "frm_mutAdm.frx":0F81
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2175
   End
End
Attribute VB_Name = "frm_mutAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "ca_adm"
Data1.Refresh


End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub
