VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      DataField       =   "qr"
      DataSource      =   "Data1"
      Height          =   1575
      Left            =   600
      ScaleHeight     =   1515
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\sappmys\sappmyspru\imagen.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "qr"
      Top             =   3600
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Picture1.Picture = LoadPicture(App.Path & "\qr.bmp")
Data1.Recordset.Update

End Sub
