VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset("ced") = Text1.Text
Data1.Recordset("ape1") = Text2.Text
Data1.Recordset("ape2") = Text3.Text
Data1.Recordset("nom1") = Text4.Text
Data1.Recordset("nom2") = Text5.Text
Data1.Recordset.Update
MsgBox "Grabado"


End Sub

Private Sub Form_Load()
Data1.DatabaseName = ""
Data1.Connect = "ODBC;DSN=personal;"
Data1.RecordSource = "personas"
Data1.Refresh

End Sub
