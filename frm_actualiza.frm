VERSION 5.00
Begin VB.Form frm_actualiza 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Actualización"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   Icon            =   "frm_actualiza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_versionlocal 
      Caption         =   "data_versionlocal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data adoversion 
      Caption         =   "adoversion"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   360
      Top             =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aguarde, se está actualizando el sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   8295
   End
End
Attribute VB_Name = "frm_actualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frm_actualiza.MousePointer = 11
adoversion.Connect = "odbc;dsn=sappnew;"
adoversion.RecordSource = "select * from version"
adoversion.Refresh

data_versionlocal.DatabaseName = App.Path & "\ctrf.mdb"
data_versionlocal.RecordSource = "ctrf"
data_versionlocal.Refresh

Timer1.Enabled = True
'If Dir(App.Path & "\SAPP\sapp.zip") <> "" Then
'   Kill (App.Path & "\SAPP\sapp.zip")
'End If

Shell ("c:\arch\unzip.exe /O " & App.Path & "\SAPP\sapp.zip"), vbMaximizedFocus
Shell ("copy /Y " & App.Path & "\SAPP\*.* " & App.Path), vbMaximizedFocus
Shell ("del " & App.Path & "\SAPP\*.*"), vbMaximizedFocus

data_versionlocal.Recordset.Edit
data_versionlocal.Recordset("ultfecact") = adoversion.Recordset("fecha")
data_versionlocal.Recordset.Update
frm_actualiza.MousePointer = 0
MsgBox "Proceso TERMINADO. Ya puede volver a iniciar el sistema."
End


End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

End Sub
