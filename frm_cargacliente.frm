VERSION 5.00
Begin VB.Form frm_cargacliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar cliente"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   Icon            =   "frm_cargacliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_old 
      Caption         =   "data_old"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "CEDULA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "MATRICULA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frm_cargacliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False

If Text1.Text <> "" Then
   If Trim(Text2.Text) = "" Then
      data_old.RecordSource = "select * from clientes where cl_codigo =" & Text1.Text
   Else
      data_old.RecordSource = "select * from clientes where cl_cedula =" & Text2.Text
   End If
   data_old.Refresh
   If data_old.Recordset.RecordCount > 0 Then
      data_old.Recordset.MoveFirst
      If Trim(Text2.Text) = "" Then
         data_cli.RecordSource = "select * from clientes where cl_codigo =" & Text1.Text
      Else
         data_cli.RecordSource = "select * from clientes where cl_cedula =" & Text2.Text
      End If
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         MsgBox "Ya existe el cliente", vbExclamation
      Else
         data_cli.Recordset.AddNew
         data_cli.Recordset("estado") = data_old.Recordset("estado")
         data_cli.Recordset("cl_codigo") = data_old.Recordset("cl_codigo")
         data_cli.Recordset("cl_codconv") = data_old.Recordset("cl_codconv")
         data_cli.Recordset("cl_nomconv") = data_old.Recordset("cl_nomconv")
         data_cli.Recordset("cl_apellid") = data_old.Recordset("cl_apellid")
         data_cli.Recordset("cl_ruc") = data_old.Recordset("cl_ruc")
         data_cli.Recordset("cl_cedula") = data_old.Recordset("cl_cedula")
         data_cli.Recordset("cl_codced") = data_old.Recordset("cl_codced")
         data_cli.Recordset("cl_fnac") = data_old.Recordset("cl_fnac")
         data_cli.Recordset("cl_codruta") = data_old.Recordset("cl_codruta")
         data_cli.Recordset("cl_direcci") = data_old.Recordset("cl_direcci")
         data_cli.Recordset("cl_entre") = data_old.Recordset("cl_entre")
         data_cli.Recordset("cl_dpto") = data_old.Recordset("cl_dpto")
         data_cli.Recordset("cl_referen") = data_old.Recordset("cl_referen")
         data_cli.Recordset("cl_grupo") = data_old.Recordset("cl_grupo")
         data_cli.Recordset("cl_zona") = data_old.Recordset("cl_zona")
         data_cli.Recordset("cl_sexo") = data_old.Recordset("cl_sexo")
         data_cli.Recordset("cl_telefon") = data_old.Recordset("cl_telefon")
         data_cli.Recordset("cl_dircobr") = data_old.Recordset("cl_dircobr")
         data_cli.Recordset("cl_nombre") = data_old.Recordset("cl_nombre")
         data_cli.Recordset("cl_socmnom") = data_old.Recordset("cl_socmnom")
         data_cli.Recordset("cl_nrosocm") = data_old.Recordset("cl_nrosocm")
         data_cli.Recordset("cl_fecing") = data_old.Recordset("cl_fecing")
         data_cli.Recordset("fecha_baja") = data_old.Recordset("fecha_baja")
         data_cli.Recordset("cl_nrovend") = data_old.Recordset("cl_nrovend")
         data_cli.Recordset("cl_nomvend") = data_old.Recordset("cl_nomvend")
         data_cli.Recordset("cl_nrocobr") = data_old.Recordset("cl_nrocobr")
         data_cli.Recordset("cl_nomcobr") = data_old.Recordset("cl_nomcobr")
         data_cli.Recordset("cl_forpago") = data_old.Recordset("cl_forpago")
         data_cli.Recordset("cl_descpag") = data_old.Recordset("cl_descpag")
         data_cli.Recordset("cl_diacobr") = data_old.Recordset("cl_diacobr")
         data_cli.Recordset("fecha_sys") = data_old.Recordset("fecha_sys")
         data_cli.Recordset("cl_decuota") = data_old.Recordset("cl_decuota")
         data_cli.Recordset("fecha_reac") = data_old.Recordset("fecha_reac")
         data_cli.Recordset("saldo_chc2") = data_old.Recordset("saldo_chc2")
         data_cli.Recordset.Update
         MsgBox "Registro grabado correctamente"
      End If
   Else
      MsgBox "No se encuentra socio"
   End If
Else
   If Trim(Text2.Text) <> "" Then
        If Trim(Text2.Text) = "" Then
           data_old.RecordSource = "select * from clientes where cl_codigo =" & Text1.Text
        Else
           data_old.RecordSource = "select * from clientes where cl_cedula =" & Text2.Text
        End If
        data_old.Refresh
        If data_old.Recordset.RecordCount > 0 Then
           data_old.Recordset.MoveFirst
           If Trim(Text2.Text) = "" Then
              data_cli.RecordSource = "select * from clientes where cl_codigo =" & Text1.Text
           Else
              data_cli.RecordSource = "select * from clientes where cl_cedula =" & Text2.Text
           End If
           data_cli.Refresh
           If data_cli.Recordset.RecordCount > 0 Then
              MsgBox "Ya existe el cliente", vbExclamation
           Else
              data_cli.Recordset.AddNew
              data_cli.Recordset("estado") = data_old.Recordset("estado")
              data_cli.Recordset("cl_codigo") = data_old.Recordset("cl_codigo")
              data_cli.Recordset("cl_codconv") = data_old.Recordset("cl_codconv")
              data_cli.Recordset("cl_nomconv") = data_old.Recordset("cl_nomconv")
              data_cli.Recordset("cl_apellid") = data_old.Recordset("cl_apellid")
              data_cli.Recordset("cl_ruc") = data_old.Recordset("cl_ruc")
              data_cli.Recordset("cl_cedula") = data_old.Recordset("cl_cedula")
              data_cli.Recordset("cl_codced") = data_old.Recordset("cl_codced")
              data_cli.Recordset("cl_fnac") = data_old.Recordset("cl_fnac")
              data_cli.Recordset("cl_codruta") = data_old.Recordset("cl_codruta")
              data_cli.Recordset("cl_direcci") = data_old.Recordset("cl_direcci")
              data_cli.Recordset("cl_entre") = data_old.Recordset("cl_entre")
              data_cli.Recordset("cl_dpto") = data_old.Recordset("cl_dpto")
              data_cli.Recordset("cl_referen") = data_old.Recordset("cl_referen")
              data_cli.Recordset("cl_grupo") = data_old.Recordset("cl_grupo")
              data_cli.Recordset("cl_zona") = data_old.Recordset("cl_zona")
              data_cli.Recordset("cl_sexo") = data_old.Recordset("cl_sexo")
              data_cli.Recordset("cl_telefon") = data_old.Recordset("cl_telefon")
              data_cli.Recordset("cl_dircobr") = data_old.Recordset("cl_dircobr")
              data_cli.Recordset("cl_nombre") = data_old.Recordset("cl_nombre")
              data_cli.Recordset("cl_socmnom") = data_old.Recordset("cl_socmnom")
              data_cli.Recordset("cl_nrosocm") = data_old.Recordset("cl_nrosocm")
              data_cli.Recordset("cl_fecing") = data_old.Recordset("cl_fecing")
              data_cli.Recordset("fecha_baja") = data_old.Recordset("fecha_baja")
              data_cli.Recordset("cl_nrovend") = data_old.Recordset("cl_nrovend")
              data_cli.Recordset("cl_nomvend") = data_old.Recordset("cl_nomvend")
              data_cli.Recordset("cl_nrocobr") = data_old.Recordset("cl_nrocobr")
              data_cli.Recordset("cl_nomcobr") = data_old.Recordset("cl_nomcobr")
              data_cli.Recordset("cl_forpago") = data_old.Recordset("cl_forpago")
              data_cli.Recordset("cl_descpag") = data_old.Recordset("cl_descpag")
              data_cli.Recordset("cl_diacobr") = data_old.Recordset("cl_diacobr")
              data_cli.Recordset("fecha_sys") = data_old.Recordset("fecha_sys")
              data_cli.Recordset("cl_decuota") = data_old.Recordset("cl_decuota")
              data_cli.Recordset("fecha_reac") = data_old.Recordset("fecha_reac")
              data_cli.Recordset("saldo_chc2") = data_old.Recordset("saldo_chc2")
              data_cli.Recordset.Update
              MsgBox "Registro grabado correctamente"
           End If
        Else
           MsgBox "No se encuentra socio"
        End If
   
   End If

End If
Command1.Enabled = True

End Sub

Private Sub Form_Load()
data_old.Connect = "odbc;dsn=sappcli;"

data_cli.Connect = "odbc;dsn=sappnew;"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub Text1_LostFocus()
If Text1.Text <> "" Then
   data_old.RecordSource = "select * from clientes where cl_codigo =" & Text1.Text
   data_old.Refresh
   If data_old.Recordset.RecordCount > 0 Then
      If IsNull(data_old.Recordset("cl_apellid")) = False Then
         Label2.Caption = data_old.Recordset("cl_apellid")
      Else
         Label2.Caption = ""
      End If
   Else
      Label2.Caption = ""
   End If
Else
   Label2.Caption = ""
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub Text2_LostFocus()
If Trim(Text2.Text) <> "" Then
   data_old.RecordSource = "select * from clientes where cl_cedula =" & Text2.Text
   data_old.Refresh
   If data_old.Recordset.RecordCount > 0 Then
      If IsNull(data_old.Recordset("cl_apellid")) = False Then
         Label2.Caption = data_old.Recordset("cl_apellid")
      Else
         Label2.Caption = ""
      End If
   Else
      Label2.Caption = ""
   End If
Else
   Label2.Caption = ""
End If

End Sub
