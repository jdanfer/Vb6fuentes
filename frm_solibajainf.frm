VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_solibajainf 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5880
   Icon            =   "frm_solibajainf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2220
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3480
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      Picture         =   "frm_solibajainf.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos para informe"
      Height          =   3135
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Por fecha de terminado"
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
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Solo funciona para la opción de Terminados"
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Incluir acciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Resumen"
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
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Detalle"
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
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_solibajainf.frx":0B14
         Left            =   1800
         List            =   "frm_solibajainf.frx":0B27
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Opción:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Rango de fechas:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4560
      Picture         =   "frm_solibajainf.frx":0B90
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frm_solibajainf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MaskEdBox2_Change()

End Sub

Private Sub Command1_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
Dim Numdesde, Numale As Long
Dim Numhasta As Long

MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh
Command1.Enabled = False
If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
   If Combo1.ListIndex = 0 Then
      Data1.RecordSource = "Select * from solic_bajas where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
      Data1.Refresh
   Else
      If Combo1.ListIndex = 1 Then
         Data1.RecordSource = "Select * from solic_bajas where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
         Data1.Refresh
      Else
         If Combo1.ListIndex = 2 Then
            If Check2.Value = 1 Then
               Data1.RecordSource = "Select * from solic_bajas where fechafin >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fechafin <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
            Else
               Data1.RecordSource = "Select * from solic_bajas where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
            End If
            Data1.Refresh
         Else
            If Combo1.ListIndex = 3 Then
               If Check2.Value = 1 Then
                  Data1.RecordSource = "Select * from solic_bajas where fechafin >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fechafin <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
               Else
                  Data1.RecordSource = "Select * from solic_bajas where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
               End If
               Data1.Refresh
            Else
               Data1.RecordSource = "Select * from solic_bajas where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
               Data1.Refresh
            End If
         End If
      End If
   End If
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_nrocobr") = Data1.Recordset("id")
         data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
         data_inf.Recordset("cl_codigo") = Data1.Recordset("matricula")
         data_inf.Recordset("cl_cedula") = Data1.Recordset("cedula")
         data_inf.Recordset("cl_codced") = Data1.Recordset("codced")
         data_inf.Recordset("cl_codconv") = Mid(Data1.Recordset("convenio"), 1, 6)
         data_inf.Recordset("cl_apellid") = Mid(Data1.Recordset("nombre"), 1, 60)
         data_inf.Recordset("cl_telefon") = Data1.Recordset("telefono")
         data_inf.Recordset("cl_dpto") = Data1.Recordset("celular")
         data_inf.Recordset("cl_nombre") = Mid(Data1.Recordset("otrotel"), 1, 30)
         data_inf.Recordset("cl_descpag") = Data1.Recordset("hora1")
         data_inf.Recordset("cl_ruc") = Data1.Recordset("hora2")
         data_inf.Recordset("cl_nomvend") = Data1.Recordset("origen")
         data_inf.Recordset("cl_nomconv") = Data1.Recordset("motivo")
         data_inf.Recordset("cl_nomcobr") = Data1.Recordset("resultado")
         If IsNull(Data1.Recordset("fechafin")) = False Then
            If Check2.Value = 1 Then
               data_inf.Recordset("cl_fecing") = Data1.Recordset("fechafin")
            End If
            data_inf.Recordset("cl_fnac") = Data1.Recordset("fechafin")
         End If
         If IsNull(Data1.Recordset("terminado")) = False Then
            If Data1.Recordset("terminado") = 1 Then
               data_inf.Recordset("cl_forpago") = 1
               data_inf.Recordset("cl_nom_sup") = "TERMINADO"
            Else
               data_inf.Recordset("cl_forpago") = 0
               data_inf.Recordset("cl_nom_sup") = "PENDIENTE"
            End If
         Else
            data_inf.Recordset("cl_forpago") = 0
            data_inf.Recordset("cl_nom_sup") = "PENDIENTE"
         End If
         If IsNull(Data1.Recordset("contrato")) = False Then
            If Data1.Recordset("contrato") = 1 Then
               data_inf.Recordset("cl_tipcli") = "SI"
            Else
               data_inf.Recordset("cl_tipcli") = "NO"
            End If
         Else
            data_inf.Recordset("cl_tipcli") = "NO"
         End If
         
         If Check1.Value = 1 Then
            Data2.RecordSource = "select * from solbaja_acc where idid =" & Data1.Recordset("id")
            Data2.Refresh
            If Data2.Recordset.RecordCount > 0 Then
               Data2.Recordset.MoveFirst
               Do While Not Data2.Recordset.EOF
                  If IsNull(data_inf.Recordset("info_debit")) = False Then
                     data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & Data2.Recordset("fecha") & "--" & Data2.Recordset("hora") & "--" & Data2.Recordset("accion")
                  Else
                     data_inf.Recordset("info_debit") = Data2.Recordset("fecha") & "--" & Data2.Recordset("hora") & "--" & Data2.Recordset("accion")
                  End If
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & chr(13) & chr(10) & "----------------------------" & chr(13) & chr(10)
                  Data2.Recordset.MoveNext
               Loop
            End If
         End If
         data_inf.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
      data_inf.RecordSource = "Select * from infcli order by cl_fecing"
      data_inf.Refresh
      MsgBox "Proceso terminado"
      If Combo1.ListIndex = 1 Then
         If Check1.Value = 1 Then
            cr1.ReportFileName = App.path & "\infsolbajapendd.rpt"
         Else
            cr1.ReportFileName = App.path & "\infsolbajapend.rpt"
         End If
         cr1.ReportTitle = "Informe de solicitudes pendientes ingresadas desde: " & md.Text & " hasta:" & mh.Text
         cr1.Action = 1
      Else
         If Combo1.ListIndex = 2 Then
            If Check1.Value = 1 Then
               cr1.ReportFileName = App.path & "\infsolbajatermd.rpt"
            Else
               cr1.ReportFileName = App.path & "\infsolbajaterm.rpt"
            End If
            cr1.ReportTitle = "Informe de solicitudes terminadas desde: " & md.Text & " hasta:" & mh.Text
            cr1.Action = 1
         Else
            If Combo1.ListIndex = 3 Then
               If Check1.Value = 1 Then
                  cr1.ReportFileName = App.path & "\infsolbajatermred.rpt"
               Else
                  cr1.ReportFileName = App.path & "\infsolbajatermre.rpt"
               End If
               cr1.ReportTitle = "Informe de solicitudes terminadas por recupero desde: " & md.Text & " hasta:" & mh.Text
               cr1.Action = 1
            Else
               If Combo1.ListIndex = 4 Then
                  If Check1.Value = 1 Then
                     cr1.ReportFileName = App.path & "\infsolbajatotmotd.rpt"
                  Else
                     cr1.ReportFileName = App.path & "\infsolbajatotmot.rpt"
                  End If
                  cr1.ReportTitle = "Informe de solicitudes por motivo de baja ingresadas desde: " & md.Text & " hasta:" & mh.Text
                  cr1.Action = 1
               Else
                  cr1.ReportFileName = App.path & "\infsolbajatotd.rpt"
                  cr1.ReportTitle = "Informe total de solicitudes ingresadas desde: " & md.Text & " hasta:" & mh.Text
                  cr1.Action = 1
               End If
            End If
         End If
      End If
   End If
Else
   MsgBox "No ingresó fechas"
End If
Command1.Enabled = True


End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_inf.DatabaseName = App.path & "\informes.mdb"


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
