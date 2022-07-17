VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_histdesp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de movimientos del llamado"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_histdesp.frx":0000
      Height          =   2655
      Left            =   360
      OleObjectBlob   =   "frm_histdesp.frx":0014
      TabIndex        =   0
      Top             =   360
      Width           =   10815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "LLAMADO NRO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   7800
      Picture         =   "frm_histdesp.frx":1093
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
End
Attribute VB_Name = "frm_histdesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Caption = frm_largador.txt_mat.Text
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

If frm_largador.txt_nro.Text <> "" Then
   Label2.Caption = frm_largador.txt_nro.Text
   Data1.RecordSource = "Select * from abmdespa where idllamado =" & frm_largador.txt_nro.Text
Else
   Data1.RecordSource = "Select * from abmdespa where idllamado =" & 0
End If
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
