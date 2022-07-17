VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_hist 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de envíos"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   Icon            =   "frm_hist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\sappmys\sappmysql\enviosrs.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "envioshist"
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_hist.frx":058A
      Height          =   3015
      Left            =   240
      OleObjectBlob   =   "frm_hist.frx":059E
      TabIndex        =   2
      Top             =   480
      Width           =   8895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FUNCIONARIO:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frm_hist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\enviosrs.mdb"
Data1.RecordSource = "Select * from envioshist where nro =" & frm_enviosueldos.Data1.Recordset("nro") & " order by fecha"
Data1.Refresh
Label2.Caption = frm_enviosueldos.Data1.Recordset("nombre")

End Sub
