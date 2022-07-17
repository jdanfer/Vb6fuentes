VERSION 5.00
Begin VB.Form frm_mensajesvar 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Mensaje"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   Icon            =   "frm_mensajesvar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Paciente con COVID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Picture         =   "frm_mensajesvar.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Paciente con covid en domicilio"
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO, realizarán carta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Picture         =   "frm_mensajesvar.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SÍ, realizan carta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Picture         =   "frm_mensajesvar.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3600
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8400
      Picture         =   "frm_mensajesvar.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cerrar mensaje"
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
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
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Atención!!"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frm_mensajesvar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If XAlta <> 1 Then
   XAlta = 0
End If

Unload Me

End Sub

Private Sub Command2_Click()
Wopscob = 2
Data2.RecordSource = "cartasnosapp"
Data2.Refresh
Data2.Recordset.AddNew
Data2.Recordset("fecha") = Date
Data2.Recordset("usuario") = WElusuario
If frm_largador.txt_ced.Text <> "" Then
   If frm_largador.t_codced.Text <> "" Then
      Data2.Recordset("cedula") = frm_largador.txt_ced.Text & "-" & frm_largador.t_codced.Text
   Else
      Data2.Recordset("cedula") = frm_largador.txt_ced.Text & "-0"
   End If
End If
If frm_largador.txt_nomb.Text <> "" Then
   Data2.Recordset("nombre") = Mid(frm_largador.txt_nomb.Text, 1, 80)
End If
If frm_largador.txt_cat.Text <> "" Then
   Data2.Recordset("convenio") = frm_largador.txt_cat.Text
End If
Data2.Recordset("opcion") = "SI"
If frm_largador.txt_nro.Text <> "" Then
   Data2.Recordset("nrolla") = Val(frm_largador.txt_nro.Text)
Else
   Data2.Recordset("nrolla") = 0
End If
Data2.Recordset("hora_lla") = frm_largador.txt_hora.Text
Data2.Recordset.Update

Unload Me

End Sub

Private Sub Command3_Click()
Wopscob = 1
Data2.RecordSource = "cartasnosapp"
Data2.Refresh
Data2.Recordset.AddNew
Data2.Recordset("fecha") = Date
Data2.Recordset("usuario") = WElusuario
If frm_largador.txt_ced.Text <> "" Then
   If frm_largador.t_codced.Text <> "" Then
      Data2.Recordset("cedula") = frm_largador.txt_ced.Text & "-" & frm_largador.t_codced.Text
   Else
      Data2.Recordset("cedula") = frm_largador.txt_ced.Text & "-0"
   End If
End If
If frm_largador.txt_nomb.Text <> "" Then
   Data2.Recordset("nombre") = Mid(frm_largador.txt_nomb.Text, 1, 80)
End If
If frm_largador.txt_cat.Text <> "" Then
   Data2.Recordset("convenio") = frm_largador.txt_cat.Text
End If
Data2.Recordset("opcion") = "NO"
If frm_largador.txt_nro.Text <> "" Then
   Data2.Recordset("nrolla") = Val(frm_largador.txt_nro.Text)
Else
   Data2.Recordset("nrolla") = 0
End If
Data2.Recordset("hora_lla") = frm_largador.txt_hora.Text

Data2.Recordset.Update
Data2.Refresh

Unload Me

End Sub

Private Sub Command4_Click()
Wopscob = 2
Data2.RecordSource = "cartasnosapp"
Data2.Refresh
Data2.Recordset.AddNew
Data2.Recordset("fecha") = Date
Data2.Recordset("usuario") = WElusuario
If frm_largador.txt_ced.Text <> "" Then
   If frm_largador.t_codced.Text <> "" Then
      Data2.Recordset("cedula") = frm_largador.txt_ced.Text & "-" & frm_largador.t_codced.Text
   Else
      Data2.Recordset("cedula") = frm_largador.txt_ced.Text & "-0"
   End If
End If
If frm_largador.txt_nomb.Text <> "" Then
   Data2.Recordset("nombre") = Mid(frm_largador.txt_nomb.Text, 1, 80)
End If
If frm_largador.txt_cat.Text <> "" Then
   Data2.Recordset("convenio") = frm_largador.txt_cat.Text
End If
Data2.Recordset("opcion") = "SI"
If frm_largador.txt_nro.Text <> "" Then
   Data2.Recordset("nrolla") = Val(frm_largador.txt_nro.Text)
Else
   Data2.Recordset("nrolla") = 0
End If
Data2.Recordset("hora_lla") = frm_largador.txt_hora.Text
Data2.Recordset("codigo") = "COVID"

Data2.Recordset.Update
Data2.Refresh
Xestaok = 19

Unload Me

End Sub

Private Sub Form_Load()
If Wopspro = 99 Then
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command1.Visible = False
Else
   Command2.Visible = False
   Command3.Visible = False
   Command1.Visible = True
End If
Data1.DatabaseName = App.path & "\mensa.mdb"
Data1.RecordSource = "mensaje"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   If IsNull(Data1.Recordset("text")) = False Then
      Label1.Caption = Data1.Recordset("text")
   Else
      Label1.Caption = "Atención!!!"
   End If
Else
   Label1.Caption = "Atención!!!"
End If
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
If XAlta = 19 Or XAlta = 17 Then
   Command4.Visible = False
End If

End Sub

Private Sub Timer1_Timer()
If Wopspro = 99 Then
Else
   Command1.Visible = True
End If
Timer1.Enabled = False

End Sub
