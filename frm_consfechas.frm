VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_consfechas 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de fechas de especialista"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frm_consfechas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton b_cierra 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_consfechas.frx":0442
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "frm_consfechas.frx":0456
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "FECHAS DEL ESPECIALISTA SELECCIONADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frm_consfechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_cierra_Click()
Unload Me

End Sub

Private Sub DBGrid1_Click()

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.RecordSource = "Select * from fechasesp where base =" & frm_espec.txt_base.Text & " and cod ='" & frm_espec.txt_cod.Text & "' order by fecha"
Data1.Refresh

End Sub
