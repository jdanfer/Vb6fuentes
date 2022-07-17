VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_envyrec 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envío y Recepción de datos"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "frm_envyrec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   3120
      Top             =   3720
   End
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
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2640
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton b_sale 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Picture         =   "frm_envyrec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton b_acep 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Picture         =   "frm_envyrec.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Envío y Recepción de datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Procesos desde disquete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Data data_envorec 
         Caption         =   "data_envorec"
         Connect         =   "dBASE IV;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_datos 
         Caption         =   "data_datos"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Recepción de información"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Envío de información"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3255
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "FECHA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_envyrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
Dim Xhastaf As Date
Xhastaf = Date - 1
FileCopy "c:\datos\vacios\env_clia.dbf", "C:\datos\env_clia.dbf"

FileCopy "c:\datos\vacios\env_clib.dbf", "C:\datos\env_clib.dbf"

FileCopy "c:\datos\vacios\env_clim.dbf", "C:\datos\env_clim.dbf"

FileCopy "c:\datos\vacios\env_lin.dbf", "C:\datos\env_lin.dbf"

FileCopy "c:\datos\vacios\env_caja.dbf", "C:\datos\env_caja.dbf"

If Check1.Value = 1 Then
Else
   data_envorec.Connect = "Access"
   data_envorec.DatabaseName = "C:\datos\envios.mdb"
End If

If mfec.Text <> "__/__/____" Then
   frm_envyrec.MousePointer = 11
   If Option1.Value = True Then
      data_envorec.RecordSource = "env_clia"
      data_envorec.Refresh
      If data_envorec.Recordset.RecordCount > 0 Then
         data_envorec.Recordset.MoveFirst
         Do While Not data_envorec.Recordset.EOF
            data_envorec.Recordset.Delete
            data_envorec.Recordset.MoveNext
         Loop
      End If
      If Data1.Recordset("base") = 3 Then
         data_datos.RecordSource = "Select * from clientes where fecha_sys >=#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
         data_datos.Refresh
      Else
         data_datos.RecordSource = "Select * from clientes where fecha_sys >=#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
         data_datos.Refresh
      End If
      If data_datos.Recordset.RecordCount > 0 Then
         data_datos.Recordset.MoveFirst
         Do While Not data_datos.Recordset.EOF
            data_envorec.Recordset.AddNew
            data_envorec.Recordset("cl_codigo") = data_datos.Recordset("cl_codigo")
            data_envorec.Recordset("cl_apellid") = data_datos.Recordset("cl_apellid")
            data_envorec.Recordset("cl_direcci") = data_datos.Recordset("cl_direcci")
            data_envorec.Recordset("cl_localid") = data_datos.Recordset("cl_localid")
            data_envorec.Recordset("cl_dpto") = data_datos.Recordset("cl_dpto")
            data_envorec.Recordset("cl_celular") = data_datos.Recordset("cl_celular")
            data_envorec.Recordset("cl_telefon") = data_datos.Recordset("cl_telefon")
            data_envorec.Recordset("cl_cedula") = data_datos.Recordset("cl_cedula")
            data_envorec.Recordset("cl_codced") = data_datos.Recordset("cl_codced")
            data_envorec.Recordset("cl_fnac") = data_datos.Recordset("cl_fnac")
            data_envorec.Recordset("cl_nrovend") = data_datos.Recordset("cl_nrovend")
            data_envorec.Recordset("cl_nomvend") = data_datos.Recordset("cl_nomvend")
            data_envorec.Recordset("cl_fecing") = data_datos.Recordset("cl_fecing")
            data_envorec.Recordset("cl_forpago") = data_datos.Recordset("cl_forpago")
            data_envorec.Recordset("cl_descpag") = data_datos.Recordset("cl_descpag")
            data_envorec.Recordset("cl_atrasoa") = data_datos.Recordset("cl_atrasoa")
            data_envorec.Recordset("cl_atrasop") = data_datos.Recordset("cl_atrasop")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_nom_sup") = data_datos.Recordset("cl_nom_sup")
            data_envorec.Recordset("cl_grupo") = data_datos.Recordset("cl_grupo")
            data_envorec.Recordset("cl_zona") = data_datos.Recordset("cl_zona")
            data_envorec.Recordset("saldo_cc") = data_datos.Recordset("saldo_cc")
            data_envorec.Recordset("cl_nrocobr") = data_datos.Recordset("cl_nrocobr")
            data_envorec.Recordset("cl_nomcobr") = data_datos.Recordset("cl_nomcobr")
            data_envorec.Recordset("cl_entre") = data_datos.Recordset("cl_entre")
            data_envorec.Recordset("cl_sexo") = data_datos.Recordset("cl_sexo")
            data_envorec.Recordset("cl_nrotarj") = data_datos.Recordset("cl_nrotarj")
            data_envorec.Recordset("cl_tjemi_c") = data_datos.Recordset("cl_tjemi_c")
            data_envorec.Recordset("cl_tjemi_n") = data_datos.Recordset("cl_tjemi_n")
            data_envorec.Recordset("cl_tj_venc") = data_datos.Recordset("cl_tj_venc")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_codconv") = data_datos.Recordset("cl_codconv")
            data_envorec.Recordset("cl_nomconv") = data_datos.Recordset("cl_nomconv")
            data_envorec.Recordset("cl_socnro") = data_datos.Recordset("cl_socnro")
            data_envorec.Recordset("cl_ultmesp") = data_datos.Recordset("cl_ultmesp")
            data_envorec.Recordset("cl_ultanop") = data_datos.Recordset("cl_ultanop")
            data_envorec.Recordset("cl_socmnro") = data_datos.Recordset("cl_socmnro")
            data_envorec.Recordset("cl_socmnom") = data_datos.Recordset("cl_socmnom")
            data_envorec.Recordset("cl_nrosocm") = data_datos.Recordset("cl_nrosocm")
            data_envorec.Recordset("fecha_baja") = data_datos.Recordset("fecha_baja")
            data_envorec.Recordset("hora_baja") = data_datos.Recordset("hora_baja")
            data_envorec.Recordset("usu_baja") = data_datos.Recordset("usu_baja")
            data_envorec.Recordset("estado") = data_datos.Recordset("estado")
            data_envorec.Recordset("fecha_reac") = data_datos.Recordset("fecha_reac")
            data_envorec.Recordset("hora_reac") = data_datos.Recordset("hora_reac")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("usu_reac") = data_datos.Recordset("usu_reac")
            data_envorec.Recordset("cl_base") = data_datos.Recordset("cl_base")
            data_envorec.Recordset("cl_dircobr") = data_datos.Recordset("cl_dircobr")
            data_envorec.Recordset("cl_edad") = data_datos.Recordset("cl_edad")
            data_envorec.Recordset("cl_uniedad") = data_datos.Recordset("cl_uniedad")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_hon_pes") = data_datos.Recordset("cl_hon_pes")
            data_envorec.Recordset("mesproxemi") = data_datos.Recordset("mesproxemi")
            data_envorec.Recordset("anoproxemi") = data_datos.Recordset("anoproxemi")
            data_envorec.Recordset("cl_diacobr") = data_datos.Recordset("cl_diacobr")
            data_envorec.Recordset("diacobro") = data_datos.Recordset("diacobro")
            data_envorec.Recordset("fecha_modi") = data_datos.Recordset("fecha_modi")
            data_envorec.Recordset("info_debit") = data_datos.Recordset("info_debit")
            data_envorec.Recordset("fecha_sys") = data_datos.Recordset("fecha_sys")
            data_envorec.Recordset("tit_tarj") = data_datos.Recordset("tit_tarj")
            data_envorec.Recordset("ci_tarj") = data_datos.Recordset("ci_tarj")
            data_envorec.Recordset("cl_emite") = data_datos.Recordset("cl_emite")
            data_envorec.Recordset("codcitarj") = data_datos.Recordset("codcitarj")
            data_envorec.Recordset.Update
            data_datos.Recordset.MoveNext
        Loop
      End If
      data_envorec.RecordSource = "env_clib"
      data_envorec.Refresh
      If data_envorec.Recordset.RecordCount > 0 Then
         data_envorec.Recordset.MoveFirst
         Do While Not data_envorec.Recordset.EOF
            data_envorec.Recordset.Delete
            data_envorec.Recordset.MoveNext
         Loop
      End If
      If Data1.Recordset("base") = 3 Then
         data_datos.RecordSource = "Select * from clientes where fecha_baja >=#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
         data_datos.Refresh
      Else
         data_datos.RecordSource = "Select * from clientes where fecha_baja >=#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
         data_datos.Refresh
      End If
      If data_datos.Recordset.RecordCount > 0 Then
         data_datos.Recordset.MoveFirst
         Do While Not data_datos.Recordset.EOF
            data_envorec.Recordset.AddNew
            data_envorec.Recordset("cl_codigo") = data_datos.Recordset("cl_codigo")
            data_envorec.Recordset("cl_apellid") = data_datos.Recordset("cl_apellid")
            data_envorec.Recordset("cl_direcci") = data_datos.Recordset("cl_direcci")
            data_envorec.Recordset("cl_localid") = data_datos.Recordset("cl_localid")
            data_envorec.Recordset("cl_dpto") = data_datos.Recordset("cl_dpto")
            data_envorec.Recordset("cl_celular") = data_datos.Recordset("cl_celular")
            data_envorec.Recordset("cl_telefon") = data_datos.Recordset("cl_telefon")
            data_envorec.Recordset("cl_cedula") = data_datos.Recordset("cl_cedula")
            data_envorec.Recordset("cl_codced") = data_datos.Recordset("cl_codced")
            data_envorec.Recordset("cl_fnac") = data_datos.Recordset("cl_fnac")
            data_envorec.Recordset("cl_nrovend") = data_datos.Recordset("cl_nrovend")
            data_envorec.Recordset("cl_nomvend") = data_datos.Recordset("cl_nomvend")
            data_envorec.Recordset("cl_fecing") = data_datos.Recordset("cl_fecing")
            data_envorec.Recordset("cl_forpago") = data_datos.Recordset("cl_forpago")
            data_envorec.Recordset("cl_descpag") = data_datos.Recordset("cl_descpag")
            data_envorec.Recordset("cl_atrasoa") = data_datos.Recordset("cl_atrasoa")
            data_envorec.Recordset("cl_atrasop") = data_datos.Recordset("cl_atrasop")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_nom_sup") = data_datos.Recordset("cl_nom_sup")
            data_envorec.Recordset("cl_grupo") = data_datos.Recordset("cl_grupo")
            data_envorec.Recordset("cl_zona") = data_datos.Recordset("cl_zona")
            data_envorec.Recordset("saldo_cc") = data_datos.Recordset("saldo_cc")
            data_envorec.Recordset("cl_nrocobr") = data_datos.Recordset("cl_nrocobr")
            data_envorec.Recordset("cl_nomcobr") = data_datos.Recordset("cl_nomcobr")
            data_envorec.Recordset("cl_entre") = data_datos.Recordset("cl_entre")
            data_envorec.Recordset("cl_sexo") = data_datos.Recordset("cl_sexo")
            data_envorec.Recordset("cl_nrotarj") = data_datos.Recordset("cl_nrotarj")
            data_envorec.Recordset("cl_tjemi_c") = data_datos.Recordset("cl_tjemi_c")
            data_envorec.Recordset("cl_tjemi_n") = data_datos.Recordset("cl_tjemi_n")
            data_envorec.Recordset("cl_tj_venc") = data_datos.Recordset("cl_tj_venc")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_codconv") = data_datos.Recordset("cl_codconv")
            data_envorec.Recordset("cl_nomconv") = data_datos.Recordset("cl_nomconv")
            data_envorec.Recordset("cl_socnro") = data_datos.Recordset("cl_socnro")
            data_envorec.Recordset("cl_ultmesp") = data_datos.Recordset("cl_ultmesp")
            data_envorec.Recordset("cl_ultanop") = data_datos.Recordset("cl_ultanop")
            data_envorec.Recordset("cl_socmnro") = data_datos.Recordset("cl_socmnro")
            data_envorec.Recordset("cl_socmnom") = data_datos.Recordset("cl_socmnom")
            data_envorec.Recordset("cl_nrosocm") = data_datos.Recordset("cl_nrosocm")
            data_envorec.Recordset("fecha_baja") = data_datos.Recordset("fecha_baja")
            data_envorec.Recordset("hora_baja") = data_datos.Recordset("hora_baja")
            data_envorec.Recordset("usu_baja") = data_datos.Recordset("usu_baja")
            data_envorec.Recordset("estado") = data_datos.Recordset("estado")
            data_envorec.Recordset("fecha_reac") = data_datos.Recordset("fecha_reac")
            data_envorec.Recordset("hora_reac") = data_datos.Recordset("hora_reac")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("usu_reac") = data_datos.Recordset("usu_reac")
            data_envorec.Recordset("cl_base") = data_datos.Recordset("cl_base")
            data_envorec.Recordset("cl_dircobr") = data_datos.Recordset("cl_dircobr")
            data_envorec.Recordset("cl_edad") = data_datos.Recordset("cl_edad")
            data_envorec.Recordset("cl_uniedad") = data_datos.Recordset("cl_uniedad")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_hon_pes") = data_datos.Recordset("cl_hon_pes")
            data_envorec.Recordset("mesproxemi") = data_datos.Recordset("mesproxemi")
            data_envorec.Recordset("anoproxemi") = data_datos.Recordset("anoproxemi")
            data_envorec.Recordset("cl_diacobr") = data_datos.Recordset("cl_diacobr")
            data_envorec.Recordset("diacobro") = data_datos.Recordset("diacobro")
            data_envorec.Recordset("fecha_modi") = data_datos.Recordset("fecha_modi")
            data_envorec.Recordset("fecha_baja") = data_datos.Recordset("fecha_baja")
            data_envorec.Recordset("info_debit") = data_datos.Recordset("info_debit")
            data_envorec.Recordset("fecha_sys") = data_datos.Recordset("fecha_sys")
            data_envorec.Recordset("tit_tarj") = data_datos.Recordset("tit_tarj")
            data_envorec.Recordset("ci_tarj") = data_datos.Recordset("ci_tarj")
            data_envorec.Recordset("cl_emite") = data_datos.Recordset("cl_emite")
            data_envorec.Recordset("codcitarj") = data_datos.Recordset("codcitarj")
            data_envorec.Recordset.Update
            data_datos.Recordset.MoveNext
        Loop
      End If
      data_envorec.RecordSource = "select * from env_clim"
      data_envorec.Refresh
      If data_envorec.Recordset.RecordCount > 0 Then
         data_envorec.Recordset.MoveFirst
         Do While Not data_envorec.Recordset.EOF
            data_envorec.Recordset.Delete
            data_envorec.Recordset.MoveNext
         Loop
      End If
      If Data1.Recordset("base") = 3 Then
         data_datos.RecordSource = "Select * from clientes where fecha_modi >=#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
         data_datos.Refresh
      Else
         data_datos.RecordSource = "Select * from clientes where fecha_modi >=#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
         data_datos.Refresh
      End If
      If data_datos.Recordset.RecordCount > 0 Then
         data_datos.Recordset.MoveFirst
         Do While Not data_datos.Recordset.EOF
            data_envorec.Recordset.AddNew
            data_envorec.Recordset("cl_codigo") = data_datos.Recordset("cl_codigo")
            data_envorec.Recordset("cl_apellid") = data_datos.Recordset("cl_apellid")
            data_envorec.Recordset("cl_direcci") = data_datos.Recordset("cl_direcci")
            data_envorec.Recordset("cl_localid") = data_datos.Recordset("cl_localid")
            data_envorec.Recordset("cl_dpto") = data_datos.Recordset("cl_dpto")
            data_envorec.Recordset("cl_celular") = data_datos.Recordset("cl_celular")
            data_envorec.Recordset("cl_telefon") = data_datos.Recordset("cl_telefon")
            data_envorec.Recordset("cl_cedula") = data_datos.Recordset("cl_cedula")
            data_envorec.Recordset("cl_codced") = data_datos.Recordset("cl_codced")
            data_envorec.Recordset("cl_fnac") = data_datos.Recordset("cl_fnac")
            data_envorec.Recordset("cl_nrovend") = data_datos.Recordset("cl_nrovend")
            data_envorec.Recordset("cl_nomvend") = data_datos.Recordset("cl_nomvend")
            data_envorec.Recordset("cl_fecing") = data_datos.Recordset("cl_fecing")
            data_envorec.Recordset("cl_forpago") = data_datos.Recordset("cl_forpago")
            data_envorec.Recordset("cl_descpag") = data_datos.Recordset("cl_descpag")
            data_envorec.Recordset("cl_atrasoa") = data_datos.Recordset("cl_atrasoa")
            data_envorec.Recordset("cl_atrasop") = data_datos.Recordset("cl_atrasop")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_nom_sup") = data_datos.Recordset("cl_nom_sup")
            data_envorec.Recordset("cl_grupo") = data_datos.Recordset("cl_grupo")
            data_envorec.Recordset("cl_zona") = data_datos.Recordset("cl_zona")
            data_envorec.Recordset("saldo_cc") = data_datos.Recordset("saldo_cc")
            data_envorec.Recordset("cl_nrocobr") = data_datos.Recordset("cl_nrocobr")
            data_envorec.Recordset("cl_nomcobr") = data_datos.Recordset("cl_nomcobr")
            data_envorec.Recordset("cl_entre") = data_datos.Recordset("cl_entre")
            data_envorec.Recordset("cl_sexo") = data_datos.Recordset("cl_sexo")
            data_envorec.Recordset("cl_nrotarj") = data_datos.Recordset("cl_nrotarj")
            data_envorec.Recordset("cl_tjemi_c") = data_datos.Recordset("cl_tjemi_c")
            data_envorec.Recordset("cl_tjemi_n") = data_datos.Recordset("cl_tjemi_n")
            data_envorec.Recordset("cl_tj_venc") = data_datos.Recordset("cl_tj_venc")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_codconv") = data_datos.Recordset("cl_codconv")
            data_envorec.Recordset("cl_nomconv") = data_datos.Recordset("cl_nomconv")
            data_envorec.Recordset("cl_socnro") = data_datos.Recordset("cl_socnro")
            data_envorec.Recordset("cl_ultmesp") = data_datos.Recordset("cl_ultmesp")
            data_envorec.Recordset("cl_ultanop") = data_datos.Recordset("cl_ultanop")
            data_envorec.Recordset("cl_socmnro") = data_datos.Recordset("cl_socmnro")
            data_envorec.Recordset("cl_socmnom") = data_datos.Recordset("cl_socmnom")
            data_envorec.Recordset("cl_nrosocm") = data_datos.Recordset("cl_nrosocm")
            data_envorec.Recordset("fecha_baja") = data_datos.Recordset("fecha_baja")
            data_envorec.Recordset("hora_baja") = data_datos.Recordset("hora_baja")
            data_envorec.Recordset("usu_baja") = data_datos.Recordset("usu_baja")
            data_envorec.Recordset("estado") = data_datos.Recordset("estado")
            data_envorec.Recordset("fecha_reac") = data_datos.Recordset("fecha_reac")
            data_envorec.Recordset("hora_reac") = data_datos.Recordset("hora_reac")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("usu_reac") = data_datos.Recordset("usu_reac")
            data_envorec.Recordset("cl_base") = data_datos.Recordset("cl_base")
            data_envorec.Recordset("cl_dircobr") = data_datos.Recordset("cl_dircobr")
            data_envorec.Recordset("cl_edad") = data_datos.Recordset("cl_edad")
            data_envorec.Recordset("cl_uniedad") = data_datos.Recordset("cl_uniedad")
            data_envorec.Recordset("cl_nro_sup") = data_datos.Recordset("cl_nro_sup")
            data_envorec.Recordset("cl_hon_pes") = data_datos.Recordset("cl_hon_pes")
            data_envorec.Recordset("mesproxemi") = data_datos.Recordset("mesproxemi")
            data_envorec.Recordset("anoproxemi") = data_datos.Recordset("anoproxemi")
            data_envorec.Recordset("cl_diacobr") = data_datos.Recordset("cl_diacobr")
            data_envorec.Recordset("diacobro") = data_datos.Recordset("diacobro")
            data_envorec.Recordset("fecha_modi") = data_datos.Recordset("fecha_modi")
            data_envorec.Recordset("info_debit") = data_datos.Recordset("info_debit")
            data_envorec.Recordset("fecha_baja") = data_datos.Recordset("fecha_baja")
            data_envorec.Recordset("fecha_sys") = data_datos.Recordset("fecha_sys")
            data_envorec.Recordset("tit_tarj") = data_datos.Recordset("tit_tarj")
            data_envorec.Recordset("ci_tarj") = data_datos.Recordset("ci_tarj")
            data_envorec.Recordset("cl_emite") = data_datos.Recordset("cl_emite")
            data_envorec.Recordset("codcitarj") = data_datos.Recordset("codcitarj")
            data_envorec.Recordset.Update
            data_datos.Recordset.MoveNext
        Loop
      End If
      data_envorec.RecordSource = "env_caja"
      data_envorec.Refresh
      If data_envorec.Recordset.RecordCount > 0 Then
         data_envorec.Recordset.MoveFirst
         Do While Not data_envorec.Recordset.EOF
            data_envorec.Recordset.Delete
            data_envorec.Recordset.MoveNext
         Loop
      End If
      If Data1.Recordset("base") = 3 Then
         data_datos.RecordSource = "Select * from caja where fecha >=#" & Format(mfec.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xhastaf, "yyyy/mm/dd") & "#"
         data_datos.Refresh
      Else
         data_datos.RecordSource = "Select * from caja where fecha >=#" & Format(mfec.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xhastaf, "yyyy/mm/dd") & "# and base =" & frm_menu.data_parse.Recordset("Base")
         data_datos.Refresh
      End If
      If data_datos.Recordset.RecordCount > 0 Then
         data_datos.Recordset.MoveFirst
         Do While Not data_datos.Recordset.EOF
            data_envorec.Recordset.AddNew
            data_envorec.Recordset("fecha") = data_datos.Recordset("fecha")
            data_envorec.Recordset("numero") = data_datos.Recordset("numero")
            data_envorec.Recordset("moneda") = data_datos.Recordset("moneda")
            data_envorec.Recordset("nombre") = data_datos.Recordset("nombre")
            data_envorec.Recordset("movimiento") = data_datos.Recordset("movimiento")
            data_envorec.Recordset("imp_fact") = data_datos.Recordset("imp_fact")
            data_envorec.Recordset("nrorub") = data_datos.Recordset("nrorub")
            data_envorec.Recordset("rubro") = data_datos.Recordset("rubro")
            data_envorec.Recordset("documento") = data_datos.Recordset("documento")
            data_envorec.Recordset("observ") = data_datos.Recordset("observ")
            data_envorec.Recordset("saldo") = data_datos.Recordset("saldo")
            data_envorec.Recordset("usuario") = data_datos.Recordset("usuario")
            data_envorec.Recordset("hora") = data_datos.Recordset("hora")
            data_envorec.Recordset("saldo_user") = data_datos.Recordset("saldo_user")
            data_envorec.Recordset("base") = data_datos.Recordset("base")
            data_envorec.Recordset("cod_serv") = data_datos.Recordset("cod_serv")
            data_envorec.Recordset("nom_serv") = data_datos.Recordset("nom_serv")
            data_envorec.Recordset("cod_socio") = data_datos.Recordset("cod_socio")
            data_envorec.Recordset("nom_socio") = data_datos.Recordset("nom_socio")
            data_envorec.Recordset("caja_mesp") = data_datos.Recordset("caja_mesp")
            data_envorec.Recordset("caja_anop") = data_datos.Recordset("caja_anop")
            data_envorec.Recordset("imp_iva") = data_datos.Recordset("imp_iva")
            data_envorec.Recordset("opiva") = data_datos.Recordset("opiva")
            data_envorec.Recordset.Update
            data_datos.Recordset.MoveNext
         Loop
      End If
      data_envorec.RecordSource = "env_lin"
      data_envorec.Refresh
      If data_envorec.Recordset.RecordCount > 0 Then
         data_envorec.Recordset.MoveFirst
         Do While Not data_envorec.Recordset.EOF
            data_envorec.Recordset.Delete
            data_envorec.Recordset.MoveNext
         Loop
      End If
      If Data1.Recordset("base") = 3 Then
         data_datos.RecordSource = "Select * from linmmdd where fecha >=#" & Format(mfec.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xhastaf, "yyyy/mm/dd") & "#"
         data_datos.Refresh
      Else
         data_datos.RecordSource = "Select * from linmmdd where fecha >=#" & Format(mfec.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xhastaf, "yyyy/mm/dd") & "# and base =" & frm_menu.data_parse.Recordset("base")
         data_datos.Refresh
      End If
      If data_datos.Recordset.RecordCount > 0 Then
         data_datos.Recordset.MoveFirst
         Do While Not data_datos.Recordset.EOF
            data_envorec.Recordset.AddNew
            data_envorec.Recordset("reg_cab") = data_datos.Recordset("reg_cab")
            data_envorec.Recordset("tipo_mov") = data_datos.Recordset("tipo_mov")
            data_envorec.Recordset("factura") = data_datos.Recordset("factura")
            data_envorec.Recordset("tipo") = data_datos.Recordset("tipo")
            data_envorec.Recordset("realizada") = data_datos.Recordset("realizada")
            data_envorec.Recordset("fecha") = data_datos.Recordset("fecha")
            data_envorec.Recordset("cod_cli") = data_datos.Recordset("cod_cli")
            data_envorec.Recordset("nom_cli") = data_datos.Recordset("nom_cli")
            data_envorec.Recordset("cod_prod") = data_datos.Recordset("cod_prod")
            data_envorec.Recordset("nom_prod") = data_datos.Recordset("nom_prod")
            data_envorec.Recordset("cantidad") = data_datos.Recordset("cantidad")
            data_envorec.Recordset("moneda") = data_datos.Recordset("moneda")
            data_envorec.Recordset("costo_prod") = data_datos.Recordset("costo_prod")
            data_envorec.Recordset("operador") = data_datos.Recordset("operador")
            data_envorec.Recordset("hora") = data_datos.Recordset("hora")
            data_envorec.Recordset("nro_flia") = data_datos.Recordset("nro_flia")
            data_envorec.Recordset("nom_flia") = data_datos.Recordset("nom_flia")
            data_envorec.Recordset("costo") = data_datos.Recordset("costo")
            data_envorec.Recordset("grupo") = data_datos.Recordset("grupo")
            data_envorec.Recordset("zona") = data_datos.Recordset("zona")
            data_envorec.Recordset("linea") = data_datos.Recordset("linea")
            data_envorec.Recordset("convenio") = data_datos.Recordset("convenio")
            data_envorec.Recordset("rub_cont") = data_datos.Recordset("rub_cont")
            data_envorec.Recordset("arancel") = data_datos.Recordset("arancel")
            data_envorec.Recordset("usa_timbre") = data_datos.Recordset("usa_timbre")
            data_envorec.Recordset("imp_timbre") = data_datos.Recordset("imp_timbre")
            data_envorec.Recordset("ced_socio") = data_datos.Recordset("ced_socio")
            data_envorec.Recordset("tot_lin") = data_datos.Recordset("tot_lin")
            data_envorec.Recordset("fact") = data_datos.Recordset("fact")
            data_envorec.Recordset("rub_nomb") = data_datos.Recordset("rub_nomb")
            data_envorec.Recordset("nro_med_a") = data_datos.Recordset("nro_med_a")
            data_envorec.Recordset("nom_med_a") = data_datos.Recordset("nom_med_a")
            data_envorec.Recordset("precio_est") = data_datos.Recordset("precio_est")
            data_envorec.Recordset("mes_paga") = data_datos.Recordset("mes_paga")
            data_envorec.Recordset("ano_paga") = data_datos.Recordset("ano_paga")
            data_envorec.Recordset("base") = data_datos.Recordset("base")
            data_envorec.Recordset("imp_iva") = data_datos.Recordset("imp_iva")
            data_envorec.Recordset("ruc") = data_datos.Recordset("ruc")
            data_envorec.Recordset.Update
            data_datos.Recordset.MoveNext
         Loop
      End If
      data_envorec.DatabaseName = ""
      data_envorec.RecordSource = ""
      data_envorec.Refresh
      If Check1.Value = 1 Then
         If Dir("c:\datos\envios.zip") <> "" Then
            Kill "c:\datos\envios.zip"
         End If
'         Shell (App.Path & "\pkunzip -o c:\datos\recibe\envios.zip c:\datos\recibe"), vbNormalFocus
         Shell (App.Path & "\pkzip -a c:\datos\envios.zip c:\datos\env_*.*"), vbNormalFocus
         Timer1.Enabled = True
      Else
         If Dir("c:\datos\envios.zip") <> "" Then
            Kill "c:\datos\envios.zip"
         End If
'         Shell (App.Path & "\pkunzip -o c:\datos\recibe\envios.zip c:\datos\recibe"), vbNormalFocus
         Shell (App.Path & "\pkzip c:\datos\envios.zip c:\datos\envios.mdb"), vbNormalFocus
         Timer1.Enabled = True
      End If
   End If
   
   If Option2.Value = True Then
      Command1_Click
   Else
'      frm_envyrec.MousePointer = 0
'      MsgBox "Proceso finalizado", vbInformation, "Mensaje"
   End If
End If
End Sub

Private Sub b_sale_Click()
Unload Me

End Sub

Private Sub Command1_Click()
If Check1.Value = 1 Then
   MsgBox "Inserte el disquete para procesar...", vbInformation, "Mensaje"
End If
   If Option2.Value = True Then
      If Check1.Value = 1 Then
        If Dir("a:\envios.zip") <> "" Then
           If Dir("c:\datos\recibe\envios.zip") <> "" Then
              Kill "c:\datos\recibe\envios.zip"
           End If
           FileCopy "a:\envios.zip", "c:\datos\recibe\envios.zip"
           Shell (App.Path & "\pkunzip -o c:\datos\recibe\envios.zip c:\datos\recibe"), vbNormalFocus
        Else
           MsgBox "No existe disquete para procesar", vbCritical, "Mensaje"
           End
        End If
      Else
        If Dir("c:\datos\recibe\envios.zip") <> "" Then
           Shell (App.Path & "\pkunzip -o c:\datos\recibe\envios.zip c:\datos\recibe"), vbNormalFocus
        Else
           MsgBox "No existe archivo para procesar", vbCritical, "Mensaje"
           End
        End If
      End If
      MsgBox "Archivo descomprimido...continuar...", vbInformation, "Recibe información"
      frm_envyrec.MousePointer = 11
      If Data1.Recordset("base") = 6 Then
        If Check1.Value = 1 Then
           data_envorec.DatabaseName = "C:\datos\recibe"
        Else
           data_envorec.Connect = "Access"
           data_envorec.DatabaseName = "c:\datos\recibe\envios.mdb"
        End If
      Else
        If Check1.Value = 1 Then
           data_envorec.DatabaseName = "C:\datos\recibe"
        Else
           data_envorec.Connect = "Access"
           data_envorec.DatabaseName = "c:\datos\recibe\envios.mdb"
        End If
        data_envorec.RecordSource = "select * from env_clia order by fecha_sys"
        data_envorec.Refresh
        If data_envorec.Recordset.RecordCount > 0 Then
           data_envorec.Recordset.MoveFirst
           data_datos.RecordSource = "Select * from clientes order by cl_codigo"
           data_datos.Refresh
           Do While Not data_envorec.Recordset.EOF
              data_datos.Recordset.FindFirst "cl_codigo =" & data_envorec.Recordset("cl_codigo")
              If Not data_datos.Recordset.NoMatch Then
                 data_envorec.Recordset.MoveNext
              Else
                 data_datos.Recordset.AddNew
                 data_datos.Recordset("cl_codigo") = data_envorec.Recordset("cl_codigo")
                 data_datos.Recordset("cl_apellid") = data_envorec.Recordset("cl_apellid")
                 data_datos.Recordset("cl_direcci") = data_envorec.Recordset("cl_direcci")
                 data_datos.Recordset("cl_localid") = data_envorec.Recordset("cl_localid")
                 data_datos.Recordset("cl_dpto") = data_envorec.Recordset("cl_dpto")
                 data_datos.Recordset("cl_celular") = data_envorec.Recordset("cl_celular")
                 data_datos.Recordset("cl_telefon") = data_envorec.Recordset("cl_telefon")
                 data_datos.Recordset("cl_cedula") = data_envorec.Recordset("cl_cedula")
                 data_datos.Recordset("cl_codced") = data_envorec.Recordset("cl_codced")
                 data_datos.Recordset("cl_fnac") = data_envorec.Recordset("cl_fnac")
                 data_datos.Recordset("cl_nrovend") = data_envorec.Recordset("cl_nrovend")
                 data_datos.Recordset("cl_nomvend") = data_envorec.Recordset("cl_nomvend")
                 data_datos.Recordset("cl_fecing") = data_envorec.Recordset("cl_fecing")
                 data_datos.Recordset("cl_forpago") = data_envorec.Recordset("cl_forpago")
                 data_datos.Recordset("cl_descpag") = data_envorec.Recordset("cl_descpag")
                 data_datos.Recordset("cl_atrasoa") = data_envorec.Recordset("cl_atrasoa")
                 data_datos.Recordset("cl_atrasop") = data_envorec.Recordset("cl_atrasop")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_nom_sup") = data_envorec.Recordset("cl_nom_sup")
                 data_datos.Recordset("cl_grupo") = data_envorec.Recordset("cl_grupo")
                 data_datos.Recordset("cl_zona") = data_envorec.Recordset("cl_zona")
                 data_datos.Recordset("saldo_cc") = data_envorec.Recordset("saldo_cc")
                 data_datos.Recordset("cl_nrocobr") = data_envorec.Recordset("cl_nrocobr")
                 data_datos.Recordset("cl_nomcobr") = data_envorec.Recordset("cl_nomcobr")
                 data_datos.Recordset("cl_entre") = data_envorec.Recordset("cl_entre")
                 data_datos.Recordset("cl_sexo") = data_envorec.Recordset("cl_sexo")
                 data_datos.Recordset("cl_nrotarj") = data_envorec.Recordset("cl_nrotarj")
                 data_datos.Recordset("cl_tjemi_c") = data_envorec.Recordset("cl_tjemi_c")
                 data_datos.Recordset("cl_tjemi_n") = data_envorec.Recordset("cl_tjemi_n")
                 data_datos.Recordset("cl_tj_venc") = data_envorec.Recordset("cl_tj_venc")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_codconv") = data_envorec.Recordset("cl_codconv")
                 data_datos.Recordset("cl_nomconv") = data_envorec.Recordset("cl_nomconv")
                 data_datos.Recordset("cl_socnro") = data_envorec.Recordset("cl_socnro")
                 data_datos.Recordset("cl_ultmesp") = data_envorec.Recordset("cl_ultmesp")
                 data_datos.Recordset("cl_ultanop") = data_envorec.Recordset("cl_ultanop")
                 data_datos.Recordset("cl_socmnro") = data_envorec.Recordset("cl_socmnro")
                 data_datos.Recordset("cl_socmnom") = data_envorec.Recordset("cl_socmnom")
                 data_datos.Recordset("cl_nrosocm") = data_envorec.Recordset("cl_nrosocm")
                 data_datos.Recordset("fecha_baja") = data_envorec.Recordset("fecha_baja")
                 data_datos.Recordset("hora_baja") = data_envorec.Recordset("hora_baja")
                 data_datos.Recordset("usu_baja") = data_envorec.Recordset("usu_baja")
                 data_datos.Recordset("estado") = data_envorec.Recordset("estado")
                 data_datos.Recordset("fecha_reac") = data_envorec.Recordset("fecha_reac")
                 data_datos.Recordset("hora_reac") = data_envorec.Recordset("hora_reac")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("usu_reac") = data_envorec.Recordset("usu_reac")
                 data_datos.Recordset("cl_base") = data_envorec.Recordset("cl_base")
                 data_datos.Recordset("cl_dircobr") = data_envorec.Recordset("cl_dircobr")
                 data_datos.Recordset("cl_edad") = data_envorec.Recordset("cl_edad")
                 data_datos.Recordset("cl_uniedad") = data_envorec.Recordset("cl_uniedad")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_hon_pes") = data_envorec.Recordset("cl_hon_pes")
                 data_datos.Recordset("mesproxemi") = data_envorec.Recordset("mesproxemi")
                 data_datos.Recordset("anoproxemi") = data_envorec.Recordset("anoproxemi")
                 data_datos.Recordset("cl_diacobr") = data_envorec.Recordset("cl_diacobr")
                 data_datos.Recordset("diacobro") = data_envorec.Recordset("diacobro")
                 data_datos.Recordset("fecha_modi") = data_envorec.Recordset("fecha_modi")
                 data_datos.Recordset("info_debit") = data_envorec.Recordset("info_debit")
                 data_datos.Recordset("fecha_sys") = data_envorec.Recordset("fecha_sys")
                 data_datos.Recordset("tit_tarj") = data_envorec.Recordset("tit_tarj")
                 data_datos.Recordset("ci_tarj") = data_envorec.Recordset("ci_tarj")
                 data_datos.Recordset("cl_emite") = data_envorec.Recordset("cl_emite")
                 data_datos.Recordset("codcitarj") = data_envorec.Recordset("codcitarj")
                 data_datos.Recordset.Update
                 data_envorec.Recordset.MoveNext
              End If
           Loop
        End If
        data_envorec.RecordSource = "select * from env_clib order by fecha_baja"
        data_envorec.Refresh
        If data_envorec.Recordset.RecordCount > 0 Then
           data_envorec.Recordset.MoveFirst
           data_datos.RecordSource = "Select * from clientes order by cl_codigo"
           data_datos.Refresh
           Do While Not data_envorec.Recordset.EOF
              data_datos.Recordset.FindFirst "cl_codigo =" & data_envorec.Recordset("cl_codigo")
              If Not data_datos.Recordset.NoMatch Then
                 data_datos.Recordset.Edit
                 data_datos.Recordset("cl_codigo") = data_envorec.Recordset("cl_codigo")
                 data_datos.Recordset("cl_apellid") = data_envorec.Recordset("cl_apellid")
                 data_datos.Recordset("cl_direcci") = data_envorec.Recordset("cl_direcci")
                 data_datos.Recordset("cl_localid") = data_envorec.Recordset("cl_localid")
                 data_datos.Recordset("cl_dpto") = data_envorec.Recordset("cl_dpto")
                 data_datos.Recordset("cl_celular") = data_envorec.Recordset("cl_celular")
                 data_datos.Recordset("cl_telefon") = data_envorec.Recordset("cl_telefon")
                 data_datos.Recordset("cl_cedula") = data_envorec.Recordset("cl_cedula")
                 data_datos.Recordset("cl_codced") = data_envorec.Recordset("cl_codced")
                 data_datos.Recordset("cl_fnac") = data_envorec.Recordset("cl_fnac")
                 data_datos.Recordset("cl_nrovend") = data_envorec.Recordset("cl_nrovend")
                 data_datos.Recordset("cl_nomvend") = data_envorec.Recordset("cl_nomvend")
                 data_datos.Recordset("cl_fecing") = data_envorec.Recordset("cl_fecing")
                 data_datos.Recordset("cl_forpago") = data_envorec.Recordset("cl_forpago")
                 data_datos.Recordset("cl_descpag") = data_envorec.Recordset("cl_descpag")
                 data_datos.Recordset("cl_atrasoa") = data_envorec.Recordset("cl_atrasoa")
                 data_datos.Recordset("cl_atrasop") = data_envorec.Recordset("cl_atrasop")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_nom_sup") = data_envorec.Recordset("cl_nom_sup")
                 data_datos.Recordset("cl_grupo") = data_envorec.Recordset("cl_grupo")
                 data_datos.Recordset("cl_zona") = data_envorec.Recordset("cl_zona")
                 data_datos.Recordset("saldo_cc") = data_envorec.Recordset("saldo_cc")
                 data_datos.Recordset("cl_nrocobr") = data_envorec.Recordset("cl_nrocobr")
                 data_datos.Recordset("cl_nomcobr") = data_envorec.Recordset("cl_nomcobr")
                 data_datos.Recordset("cl_entre") = data_envorec.Recordset("cl_entre")
                 data_datos.Recordset("cl_sexo") = data_envorec.Recordset("cl_sexo")
                 data_datos.Recordset("cl_nrotarj") = data_envorec.Recordset("cl_nrotarj")
                 data_datos.Recordset("cl_tjemi_c") = data_envorec.Recordset("cl_tjemi_c")
                 data_datos.Recordset("cl_tjemi_n") = data_envorec.Recordset("cl_tjemi_n")
                 data_datos.Recordset("cl_tj_venc") = data_envorec.Recordset("cl_tj_venc")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_codconv") = data_envorec.Recordset("cl_codconv")
                 data_datos.Recordset("cl_nomconv") = data_envorec.Recordset("cl_nomconv")
                 data_datos.Recordset("cl_socnro") = data_envorec.Recordset("cl_socnro")
                 data_datos.Recordset("cl_ultmesp") = data_envorec.Recordset("cl_ultmesp")
                 data_datos.Recordset("cl_ultanop") = data_envorec.Recordset("cl_ultanop")
                 data_datos.Recordset("cl_socmnro") = data_envorec.Recordset("cl_socmnro")
                 data_datos.Recordset("cl_socmnom") = data_envorec.Recordset("cl_socmnom")
                 data_datos.Recordset("cl_nrosocm") = data_envorec.Recordset("cl_nrosocm")
                 data_datos.Recordset("fecha_baja") = data_envorec.Recordset("fecha_baja")
                 data_datos.Recordset("hora_baja") = data_envorec.Recordset("hora_baja")
                 data_datos.Recordset("usu_baja") = data_envorec.Recordset("usu_baja")
                 data_datos.Recordset("estado") = data_envorec.Recordset("estado")
                 data_datos.Recordset("fecha_reac") = data_envorec.Recordset("fecha_reac")
                 data_datos.Recordset("hora_reac") = data_envorec.Recordset("hora_reac")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("usu_reac") = data_envorec.Recordset("usu_reac")
                 data_datos.Recordset("cl_base") = data_envorec.Recordset("cl_base")
                 data_datos.Recordset("cl_dircobr") = data_envorec.Recordset("cl_dircobr")
                 data_datos.Recordset("cl_edad") = data_envorec.Recordset("cl_edad")
                 data_datos.Recordset("cl_uniedad") = data_envorec.Recordset("cl_uniedad")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_hon_pes") = data_envorec.Recordset("cl_hon_pes")
                 data_datos.Recordset("mesproxemi") = data_envorec.Recordset("mesproxemi")
                 data_datos.Recordset("anoproxemi") = data_envorec.Recordset("anoproxemi")
                 data_datos.Recordset("cl_diacobr") = data_envorec.Recordset("cl_diacobr")
                 data_datos.Recordset("diacobro") = data_envorec.Recordset("diacobro")
                 data_datos.Recordset("fecha_modi") = data_envorec.Recordset("fecha_modi")
                 data_datos.Recordset("info_debit") = data_envorec.Recordset("info_debit")
                 data_datos.Recordset("fecha_sys") = data_envorec.Recordset("fecha_sys")
                 data_datos.Recordset("tit_tarj") = data_envorec.Recordset("tit_tarj")
                 data_datos.Recordset("ci_tarj") = data_envorec.Recordset("ci_tarj")
                 data_datos.Recordset("cl_emite") = data_envorec.Recordset("cl_emite")
                 data_datos.Recordset("codcitarj") = data_envorec.Recordset("codcitarj")
                 data_datos.Recordset.Update
                 data_envorec.Recordset.MoveNext
              Else
                 data_envorec.Recordset.MoveNext
              End If
           Loop
        End If
        data_envorec.RecordSource = "select * from env_clim order by fecha_modi"
        data_envorec.Refresh
        If data_envorec.Recordset.RecordCount > 0 Then
           data_envorec.Recordset.MoveFirst
           data_datos.RecordSource = "Select * from clientes order by cl_codigo"
           data_datos.Refresh
           Do While Not data_envorec.Recordset.EOF
              data_datos.Recordset.FindFirst "cl_codigo =" & data_envorec.Recordset("cl_codigo")
              If Not data_datos.Recordset.NoMatch Then
                 data_datos.Recordset.Edit
                 data_datos.Recordset("cl_codigo") = data_envorec.Recordset("cl_codigo")
                 data_datos.Recordset("cl_apellid") = data_envorec.Recordset("cl_apellid")
                 data_datos.Recordset("cl_direcci") = data_envorec.Recordset("cl_direcci")
                 data_datos.Recordset("cl_localid") = data_envorec.Recordset("cl_localid")
                 data_datos.Recordset("cl_dpto") = data_envorec.Recordset("cl_dpto")
                 data_datos.Recordset("cl_celular") = data_envorec.Recordset("cl_celular")
                 data_datos.Recordset("cl_telefon") = data_envorec.Recordset("cl_telefon")
                 data_datos.Recordset("cl_cedula") = data_envorec.Recordset("cl_cedula")
                 data_datos.Recordset("cl_codced") = data_envorec.Recordset("cl_codced")
                 data_datos.Recordset("cl_fnac") = data_envorec.Recordset("cl_fnac")
                 data_datos.Recordset("cl_nrovend") = data_envorec.Recordset("cl_nrovend")
                 data_datos.Recordset("cl_nomvend") = data_envorec.Recordset("cl_nomvend")
                 data_datos.Recordset("cl_fecing") = data_envorec.Recordset("cl_fecing")
                 data_datos.Recordset("cl_forpago") = data_envorec.Recordset("cl_forpago")
                 data_datos.Recordset("cl_descpag") = data_envorec.Recordset("cl_descpag")
                 data_datos.Recordset("cl_atrasoa") = data_envorec.Recordset("cl_atrasoa")
                 data_datos.Recordset("cl_atrasop") = data_envorec.Recordset("cl_atrasop")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_nom_sup") = data_envorec.Recordset("cl_nom_sup")
                 data_datos.Recordset("cl_grupo") = data_envorec.Recordset("cl_grupo")
                 data_datos.Recordset("cl_zona") = data_envorec.Recordset("cl_zona")
                 data_datos.Recordset("saldo_cc") = data_envorec.Recordset("saldo_cc")
                 data_datos.Recordset("cl_nrocobr") = data_envorec.Recordset("cl_nrocobr")
                 data_datos.Recordset("cl_nomcobr") = data_envorec.Recordset("cl_nomcobr")
                 data_datos.Recordset("cl_entre") = data_envorec.Recordset("cl_entre")
                 data_datos.Recordset("cl_sexo") = data_envorec.Recordset("cl_sexo")
                 data_datos.Recordset("cl_nrotarj") = data_envorec.Recordset("cl_nrotarj")
                 data_datos.Recordset("cl_tjemi_c") = data_envorec.Recordset("cl_tjemi_c")
                 data_datos.Recordset("cl_tjemi_n") = data_envorec.Recordset("cl_tjemi_n")
                 data_datos.Recordset("cl_tj_venc") = data_envorec.Recordset("cl_tj_venc")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_codconv") = data_envorec.Recordset("cl_codconv")
                 data_datos.Recordset("cl_nomconv") = data_envorec.Recordset("cl_nomconv")
                 data_datos.Recordset("cl_socnro") = data_envorec.Recordset("cl_socnro")
                 data_datos.Recordset("cl_ultmesp") = data_envorec.Recordset("cl_ultmesp")
                 data_datos.Recordset("cl_ultanop") = data_envorec.Recordset("cl_ultanop")
                 data_datos.Recordset("cl_socmnro") = data_envorec.Recordset("cl_socmnro")
                 data_datos.Recordset("cl_socmnom") = data_envorec.Recordset("cl_socmnom")
                 data_datos.Recordset("cl_nrosocm") = data_envorec.Recordset("cl_nrosocm")
                 data_datos.Recordset("fecha_baja") = data_envorec.Recordset("fecha_baja")
                 data_datos.Recordset("hora_baja") = data_envorec.Recordset("hora_baja")
                 data_datos.Recordset("usu_baja") = data_envorec.Recordset("usu_baja")
                 data_datos.Recordset("estado") = data_envorec.Recordset("estado")
                 data_datos.Recordset("fecha_reac") = data_envorec.Recordset("fecha_reac")
                 data_datos.Recordset("hora_reac") = data_envorec.Recordset("hora_reac")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("usu_reac") = data_envorec.Recordset("usu_reac")
                 data_datos.Recordset("cl_base") = data_envorec.Recordset("cl_base")
                 data_datos.Recordset("cl_dircobr") = data_envorec.Recordset("cl_dircobr")
                 data_datos.Recordset("cl_edad") = data_envorec.Recordset("cl_edad")
                 data_datos.Recordset("cl_uniedad") = data_envorec.Recordset("cl_uniedad")
                 data_datos.Recordset("cl_nro_sup") = data_envorec.Recordset("cl_nro_sup")
                 data_datos.Recordset("cl_hon_pes") = data_envorec.Recordset("cl_hon_pes")
                 data_datos.Recordset("mesproxemi") = data_envorec.Recordset("mesproxemi")
                 data_datos.Recordset("anoproxemi") = data_envorec.Recordset("anoproxemi")
                 data_datos.Recordset("cl_diacobr") = data_envorec.Recordset("cl_diacobr")
                 data_datos.Recordset("diacobro") = data_envorec.Recordset("diacobro")
                 data_datos.Recordset("fecha_modi") = data_envorec.Recordset("fecha_modi")
                 data_datos.Recordset("info_debit") = data_envorec.Recordset("info_debit")
                 data_datos.Recordset("fecha_sys") = data_envorec.Recordset("fecha_sys")
                 data_datos.Recordset("tit_tarj") = data_envorec.Recordset("tit_tarj")
                 data_datos.Recordset("ci_tarj") = data_envorec.Recordset("ci_tarj")
                 data_datos.Recordset("cl_emite") = data_envorec.Recordset("cl_emite")
                 data_datos.Recordset("codcitarj") = data_envorec.Recordset("codcitarj")
                 data_datos.Recordset.Update
                 data_envorec.Recordset.MoveNext
              Else
                 data_envorec.Recordset.MoveNext
              End If
           Loop
        End If
      End If
      data_envorec.RecordSource = "select * from env_caja"
      data_envorec.Refresh
      If data_envorec.Recordset.RecordCount > 0 Then
         data_envorec.Recordset.MoveFirst
         data_datos.RecordSource = "Select * from caja order by fecha,base"
         data_datos.Refresh
         data_datos.Recordset.FindFirst "fecha =#" & Format(data_envorec.Recordset("fecha"), "yyyy/mm/dd") & "# and base =" & data_envorec.Recordset("base")
         If Not data_datos.Recordset.NoMatch Then
            MsgBox "Ya existe una fecha ingresada con ésta base, Verifique!!", vbCritical
         Else
            Do While Not data_envorec.Recordset.EOF
               data_datos.Recordset.AddNew
               data_datos.Recordset("fecha") = data_envorec.Recordset("fecha")
               data_datos.Recordset("numero") = data_envorec.Recordset("numero")
               data_datos.Recordset("moneda") = data_envorec.Recordset("moneda")
               data_datos.Recordset("nombre") = data_envorec.Recordset("nombre")
               data_datos.Recordset("movimiento") = data_envorec.Recordset("movimiento")
               data_datos.Recordset("imp_fact") = data_envorec.Recordset("imp_fact")
               data_datos.Recordset("nrorub") = data_envorec.Recordset("nrorub")
               data_datos.Recordset("rubro") = data_envorec.Recordset("rubro")
               data_datos.Recordset("documento") = data_envorec.Recordset("documento")
               data_datos.Recordset("observ") = data_envorec.Recordset("observ")
               data_datos.Recordset("saldo") = data_envorec.Recordset("saldo")
               data_datos.Recordset("usuario") = data_envorec.Recordset("usuario")
               data_datos.Recordset("hora") = data_envorec.Recordset("hora")
               data_datos.Recordset("saldo_user") = data_envorec.Recordset("saldo_user")
               data_datos.Recordset("base") = data_envorec.Recordset("base")
               data_datos.Recordset("cod_serv") = data_envorec.Recordset("cod_serv")
               data_datos.Recordset("nom_serv") = data_envorec.Recordset("nom_serv")
               data_datos.Recordset("cod_socio") = data_envorec.Recordset("cod_socio")
               data_datos.Recordset("nom_socio") = data_envorec.Recordset("nom_socio")
               data_datos.Recordset("caja_mesp") = data_envorec.Recordset("caja_mesp")
               data_datos.Recordset("caja_anop") = data_envorec.Recordset("caja_anop")
               data_datos.Recordset("imp_iva") = data_envorec.Recordset("imp_iva")
               data_datos.Recordset("opiva") = data_envorec.Recordset("opiva")
               data_datos.Recordset.Update
               data_envorec.Recordset.MoveNext
            Loop
         End If
      End If
      data_envorec.RecordSource = "select * from env_lin"
      data_envorec.Refresh
      If data_envorec.Recordset.RecordCount > 0 Then
         data_envorec.Recordset.MoveFirst
         data_datos.RecordSource = "Select * from linmmdd order by fecha,base"
         data_datos.Refresh
         data_datos.Recordset.FindFirst "fecha =#" & Format(data_envorec.Recordset("fecha"), "yyyy/mm/dd") & "# and base =" & data_envorec.Recordset("base")
         If Not data_datos.Recordset.NoMatch Then
            MsgBox "Ya existe una fecha ingresada con ésta base, Verifique!!", vbCritical
         Else
            Do While Not data_envorec.Recordset.EOF
              data_datos.Recordset.AddNew
              data_datos.Recordset("reg_cab") = data_envorec.Recordset("reg_cab")
              data_datos.Recordset("tipo_mov") = data_envorec.Recordset("tipo_mov")
              data_datos.Recordset("factura") = data_envorec.Recordset("factura")
              data_datos.Recordset("tipo") = data_envorec.Recordset("tipo")
              data_datos.Recordset("realizada") = data_envorec.Recordset("realizada")
              data_datos.Recordset("fecha") = data_envorec.Recordset("fecha")
              data_datos.Recordset("cod_cli") = data_envorec.Recordset("cod_cli")
              data_datos.Recordset("nom_cli") = data_envorec.Recordset("nom_cli")
              data_datos.Recordset("cod_prod") = data_envorec.Recordset("cod_prod")
              data_datos.Recordset("nom_prod") = data_envorec.Recordset("nom_prod")
              data_datos.Recordset("cantidad") = data_envorec.Recordset("cantidad")
              data_datos.Recordset("moneda") = data_envorec.Recordset("moneda")
              data_datos.Recordset("costo_prod") = data_envorec.Recordset("costo_prod")
              data_datos.Recordset("operador") = data_envorec.Recordset("operador")
              data_datos.Recordset("hora") = data_envorec.Recordset("hora")
              data_datos.Recordset("nro_flia") = data_envorec.Recordset("nro_flia")
              data_datos.Recordset("nom_flia") = data_envorec.Recordset("nom_flia")
              data_datos.Recordset("costo") = data_envorec.Recordset("costo")
              data_datos.Recordset("grupo") = data_envorec.Recordset("grupo")
              data_datos.Recordset("zona") = data_envorec.Recordset("zona")
              data_datos.Recordset("linea") = data_envorec.Recordset("linea")
              data_datos.Recordset("convenio") = data_envorec.Recordset("convenio")
              data_datos.Recordset("rub_cont") = data_envorec.Recordset("rub_cont")
              data_datos.Recordset("arancel") = data_envorec.Recordset("arancel")
              data_datos.Recordset("usa_timbre") = data_envorec.Recordset("usa_timbre")
              data_datos.Recordset("imp_timbre") = data_envorec.Recordset("imp_timbre")
              data_datos.Recordset("ced_socio") = data_envorec.Recordset("ced_socio")
              data_datos.Recordset("tot_lin") = data_envorec.Recordset("tot_lin")
              data_datos.Recordset("fact") = data_envorec.Recordset("fact")
              data_datos.Recordset("rub_nomb") = data_envorec.Recordset("rub_nomb")
              data_datos.Recordset("nro_med_a") = data_envorec.Recordset("nro_med_a")
              data_datos.Recordset("nom_med_a") = data_envorec.Recordset("nom_med_a")
              data_datos.Recordset("precio_est") = data_envorec.Recordset("precio_est")
              data_datos.Recordset("mes_paga") = data_envorec.Recordset("mes_paga")
              data_datos.Recordset("ano_paga") = data_envorec.Recordset("ano_paga")
              data_datos.Recordset("base") = data_envorec.Recordset("base")
              data_datos.Recordset("imp_iva") = data_envorec.Recordset("imp_iva")
              data_datos.Recordset("ruc") = data_envorec.Recordset("ruc")
              data_datos.Recordset.Update
              data_envorec.Recordset.MoveNext
           Loop
         End If
      End If
      frm_envyrec.MousePointer = 0
      MsgBox "Proceso terminado"
      
   End If


End Sub

Private Sub Form_Load()
Dim mifecha As Date
mifecha = Date - 1
mfec.Text = Format(mifecha, "dd/mm/yyyy")

data_datos.DatabaseName = App.Path & "\sapp.mdb"
data_envorec.DatabaseName = "c:\datos"
Data1.DatabaseName = App.Path & "\parse.mdb"
Data1.RecordSource = "parsec0"
Data1.Refresh

End Sub

Private Sub Timer1_Timer()
If Check1.Value = 1 Then
   MsgBox "Inserte un disquete para copiar envios.zip", vbExclamation, "Mensaje"
   FileCopy "c:\datos\envios.zip", "A:\envios.zip"
End If
frm_envyrec.MousePointer = 0
MsgBox "Proceso finalizado", vbInformation, "Mensaje"

Timer1.Enabled = False

End Sub
