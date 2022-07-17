VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Data data_env 
      Caption         =   "data_env"
      Connect         =   "Access"
      DatabaseName    =   "D:\sapprespaldo\sapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENTES"
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   "Z:\sapp\sapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENTES"
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   2880
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
data_cli.Recordset.MoveFirst
Do While Not data_cli.Recordset.EOF
     data_env.Recordset.AddNew
     data_env.Recordset("estado") = data_cli.Recordset("estado")
     data_env.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
     data_env.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
     data_env.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
     data_env.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
     If IsNull(data_cli.Recordset("cl_cedula")) = False Then
        data_env.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
     Else
        data_env.Recordset("cl_cedula") = 0
     End If
     If IsNull(data_cli.Recordset("cl_codced")) = False Then
        data_env.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
     Else
        data_env.Recordset("cl_codced") = 0
     End If
     If IsNull(data_cli.Recordset("cl_fnac")) = False Then
        data_env.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
     End If
     If IsNull(data_cli.Recordset("cl_edad")) = False Then
        data_env.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
     Else
        data_env.Recordset("cl_edad") = 0
     End If
     If IsNull(data_cli.Recordset("cl_uniedad")) = False Then
        data_env.Recordset("cl_uniedad") = data_cli.Recordset("cl_uniedad")
     Else
        data_env.Recordset("cl_uniedad") = "A"
     End If
     If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
        data_env.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
     Else
        data_env.Recordset("cl_ultmesp") = 0
     End If
     If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
        data_env.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
     Else
        data_env.Recordset("cl_ultanop") = 0
     End If
     If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
        data_env.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa")
     Else
        data_env.Recordset("cl_atrasoa") = 0
     End If
     If IsNull(data_cli.Recordset("saldo_cc")) = False Then
        data_env.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
     Else
        data_env.Recordset("saldo_cc") = 0
     End If
     If IsNull(data_cli.Recordset("cl_direcci")) = False Then
        data_env.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
     Else
        data_env.Recordset("cl_direcci") = ""
     End If
     If IsNull(data_cli.Recordset("cl_entre")) = False Then
        data_env.Recordset("cl_entre") = data_cli.Recordset("cl_entre")
     Else
        data_env.Recordset("cl_entre") = ""
     End If
     If IsNull(data_cli.Recordset("cl_grupo")) = False Then
        data_env.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
        data_env.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
     Else
        data_env.Recordset("cl_grupo") = 0
        data_env.Recordset("cl_zona") = ""
     End If
     data_env.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
     If IsNull(data_cli.Recordset("cl_telefon")) = False Then
        data_env.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
     Else
        data_env.Recordset("cl_telefon") = ""
     End If
     If IsNull(data_cli.Recordset("cl_dircobr")) = False Then
        data_env.Recordset("cl_dircobr") = data_cli.Recordset("cl_dircobr")
     Else
        data_env.Recordset("cl_dircobr") = ""
     End If
     If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
        data_env.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
     Else
        data_env.Recordset("cl_socmnom") = ""
     End If
     If IsNull(data_cli.Recordset("cl_socmnro")) = False Then
        data_env.Recordset("cl_socmnro") = data_cli.Recordset("cl_socmnro")
     Else
        data_env.Recordset("cl_socmnro") = ""
     End If
     If IsNull(data_cli.Recordset("cl_fecing")) = False Then
        data_env.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
     End If
     If IsNull(data_cli.Recordset("fecha_baja")) = False Then
        data_env.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
     End If
     If IsNull(data_cli.Recordset("cl_nrovend")) = False Then
        data_env.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
        data_env.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
     Else
        data_env.Recordset("cl_nrovend") = 799
        data_env.Recordset("cl_nomvend") = "*TODOS"
     End If
     If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
        data_env.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
        data_env.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
     Else
        data_env.Recordset("cl_nrocobr") = 0
        data_env.Recordset("cl_nomcobr") = "*TODOS"
     End If
     If IsNull(data_cli.Recordset("cl_forpago")) = False Then
        data_env.Recordset("cl_forpago") = data_cli.Recordset("cl_forpago")
        data_env.Recordset("cl_descpag") = data_cli.Recordset("cl_descpag")
     Else
        data_env.Recordset("cl_forpago") = 1
        data_env.Recordset("cl_descpag") = "Abono Mensual"
     End If
     If IsNull(data_cli.Recordset("cl_diacobr")) = False Then
        data_env.Recordset("cl_diacobr") = data_cli.Recordset("cl_diacobr")
     Else
        data_env.Recordset("cl_diacobr") = ""
     End If
     If IsNull(data_cli.Recordset("tit_tarj")) = False Then
        data_env.Recordset("tit_tarj") = data_cli.Recordset("tit_tarj")
     Else
        data_env.Recordset("tit_tarj") = ""
     End If
     If IsNull(data_cli.Recordset("cl_nrotarj")) = False Then
        data_env.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
     Else
        data_env.Recordset("cl_nrotarj") = 0
     End If
     If IsNull(data_cli.Recordset("ci_tarj")) = False Then
        data_env.Recordset("ci_tarj") = data_cli.Recordset("ci_tarj")
     Else
        data_env.Recordset("ci_tarj") = 0
     End If
     If IsNull(data_cli.Recordset("codcitarj")) = False Then
        data_env.Recordset("codcitarj") = data_cli.Recordset("codcitarj")
     Else
        data_env.Recordset("codcitarj") = 0
     End If
     If IsNull(data_cli.Recordset("cl_tjemi_c")) = False Then
        data_env.Recordset("cl_tjemi_c") = data_cli.Recordset("cl_tjemi_c")
     Else
        data_env.Recordset("cl_tjemi_c") = 0
     End If
     If IsNull(data_cli.Recordset("cl_tjemi_n")) = False Then
        data_env.Recordset("cl_tjemi_n") = data_cli.Recordset("cl_tjemi_n")
     Else
        data_env.Recordset("cl_tjemi_n") = 0
     End If
     If IsNull(data_cli.Recordset("cl_tj_venc")) = False Then
        data_env.Recordset("cl_tj_venc") = data_cli.Recordset("cl_tj_venc")
     End If
     If IsNull(data_cli.Recordset("fecha_sys")) = False Then
        data_env.Recordset("fecha_sys") = data_cli.Recordset("fecha_sys")
     End If
     If IsNull(data_cli.Recordset("fecha_modi")) = False Then
        data_env.Recordset("fecha_modi") = data_cli.Recordset("fecha_modi")
     End If
     data_env.Recordset.Update
     data_cli.Recordset.MoveNext
Loop
MsgBox "Terminado"

End Sub

