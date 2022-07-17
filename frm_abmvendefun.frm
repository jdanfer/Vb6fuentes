VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_abmvendefun 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendedores (Funcionarios)"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   Icon            =   "frm_abmvendefun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_abmvendefun.frx":058A
      Height          =   2895
      Left            =   240
      OleObjectBlob   =   "frm_abmvendefun.frx":059E
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "vende_func"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "frm_abmvendefun.frx":0F81
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2535
   End
End
Attribute VB_Name = "frm_abmvendefun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub
