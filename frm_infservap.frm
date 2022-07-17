VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infservap 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de servicios a A.P. y 2da.Opinión Médica"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infservap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Data data_reg 
      Caption         =   "data_reg"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2460
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   960
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      Picture         =   "frm_infservap.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infservap.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Procesar"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Opciones de informe"
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infservap.frx":0F56
         Left            =   1440
         List            =   "frm_infservap.frx":0F60
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Todo"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pendiente"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Realizado"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Value           =   -1  'True
         Width           =   3735
      End
      Begin MSMask.MaskEdBox mfhh 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfdd 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Opción:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "RANGO de FECHAS:"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1800
      Picture         =   "frm_infservap.frx":0F86
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1695
   End
End
Attribute VB_Name = "frm_infservap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frm_infservap.MousePointer = 11
data_inf.RecordSource = "infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If
If Combo1.ListIndex = 1 Then
   If WElusuario = "SPEREZ" Or WElusuario = "YULIANAD" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Or WElusuario = "MARCELOM" Or WElusuario = "DARIOH" Then
      data_reg.Connect = "ODBC;DSN=" & Xconexrmt & ";"
   Else
      MsgBox "Usuario no habilitado"
      End
   End If
Else
   data_reg.Connect = "ODBC;DSN=sappespecial;"
End If

If mfdd.Text <> "__/__/____" And mfhh.Text <> "__/__/____" Then
   If Option1.Value = True Then
      If Combo1.ListIndex = 1 Then
         data_reg.RecordSource = "Select * from segunda_op where fecha >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "# and finfecha is not null"
         data_reg.Refresh
      Else
         data_reg.RecordSource = "Select * from env_soc where cl_fultmov >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and cl_fultmov <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "#"
         data_reg.Refresh
      End If
   Else
      If Combo1.ListIndex = 1 Then
         If Option2.Value = True Then
            data_reg.RecordSource = "Select * from segunda_op where fecha >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "# and finfecha is null"
            data_reg.Refresh
         Else
            data_reg.RecordSource = "Select * from segunda_op where fecha >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "#"
            data_reg.Refresh
         End If
      Else
         data_reg.RecordSource = "Select * from env_soc where cl_fnac >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "#"
         data_reg.Refresh
      End If
   End If
   If data_reg.Recordset.RecordCount > 0 Then
      data_reg.Recordset.MoveFirst
      If Option1.Value = True Then
         Do While Not data_reg.Recordset.EOF
            If Combo1.ListIndex = 1 Then
            
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_fnac") = data_reg.Recordset("fecha")
               data_inf.Recordset("cl_ruc") = data_reg.Recordset("hora")
               data_inf.Recordset("info_debit") = Mid(data_reg.Recordset("detalle"), 1, 130)
               data_inf.Recordset("cl_descpag") = Mid(data_reg.Recordset("soc_nom"), 1, 25)
               data_inf.Recordset("cl_nrovend") = data_reg.Recordset("base")
               data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("usuario")
               data_inf.Recordset("cl_fultmov") = data_reg.Recordset("finfecha")
               data_inf.Recordset("cl_fax") = data_reg.Recordset("finhora")
               If IsNull(data_reg.Recordset("finobs")) = False Then
                  data_inf.Recordset("cl_email") = Mid(data_reg.Recordset("finobs"), 1, 80)
               End If
               If Val(data_reg.Recordset("confop")) = 0 Then
                  data_inf.Recordset("cl_zona") = "CONFORME"
               Else
                  If Val(data_reg.Recordset("confop")) = 1 Then
                     data_inf.Recordset("cl_zona") = "CON DEMORA"
                  Else
                     If Val(data_reg.Recordset("confop")) = 2 Then
                        data_inf.Recordset("cl_zona") = "NO CONFORME"
                     End If
                  End If
               End If
'               data_inf.Recordset("cl_nomcobr") = data_reg.Recordset("cl_nomcobr")
               data_inf.Recordset.Update
               data_reg.Recordset.MoveNext
            Else
                If IsNull(data_reg.Recordset("cl_fultmov")) = True Then
                   data_reg.Recordset.MoveNext
                Else
                   data_inf.Recordset.AddNew
                   data_inf.Recordset("cl_fnac") = data_reg.Recordset("cl_fnac")
                   data_inf.Recordset("cl_ruc") = data_reg.Recordset("cl_ruc")
                   data_inf.Recordset("info_debit") = Mid(data_reg.Recordset("info_debit"), 1, 130)
                   data_inf.Recordset("cl_descpag") = Mid(data_reg.Recordset("cl_descpag"), 1, 25)
                   data_inf.Recordset("cl_nrovend") = data_reg.Recordset("cl_nrovend")
                   data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("cl_nom_sup")
                   data_inf.Recordset("cl_fultmov") = data_reg.Recordset("cl_fultmov")
                   data_inf.Recordset("cl_fax") = data_reg.Recordset("cl_fax")
                   data_inf.Recordset("cl_email") = Mid(data_reg.Recordset("cl_email"), 1, 80)
                   data_inf.Recordset("cl_zona") = data_reg.Recordset("cl_zona")
                   data_inf.Recordset("cl_nomcobr") = data_reg.Recordset("cl_nomcobr")
                   data_inf.Recordset.Update
                   data_reg.Recordset.MoveNext
                End If
            End If
         Loop
         data_inf.RecordSource = "Select * from infcli"
         data_inf.Refresh
         frm_infservap.MousePointer = 0
         cr1.ReportFileName = App.path & "\infregrea.rpt"
         cr1.ReportTitle = "REGISTROS CUMPLIDOS DESDE: " & mfdd.Text & " HASTA: " & mfhh.Text
         cr1.Action = 1
         
      End If
      If Option2.Value = True Then
         Do While Not data_reg.Recordset.EOF
            If Combo1.ListIndex = 1 Then
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_fnac") = data_reg.Recordset("fecha")
               data_inf.Recordset("cl_ruc") = data_reg.Recordset("hora")
               data_inf.Recordset("info_debit") = Mid(data_reg.Recordset("detalle"), 1, 130)
               data_inf.Recordset("cl_descpag") = Mid(data_reg.Recordset("soc_nom"), 1, 25)
               data_inf.Recordset("cl_nrovend") = data_reg.Recordset("base")
               data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("usuario")
               data_inf.Recordset("cl_fultmov") = data_reg.Recordset("finfecha")
               data_inf.Recordset("cl_fax") = data_reg.Recordset("finhora")
               If IsNull(data_reg.Recordset("finobs")) = False Then
                  data_inf.Recordset("cl_email") = Mid(data_reg.Recordset("finobs"), 1, 80)
               End If
               If Val(data_reg.Recordset("confop")) = 0 Then
                  data_inf.Recordset("cl_zona") = "CONFORME"
               Else
                  If Val(data_reg.Recordset("confop")) = 1 Then
                     data_inf.Recordset("cl_zona") = "CON DEMORA"
                  Else
                     If Val(data_reg.Recordset("confop")) = 2 Then
                        data_inf.Recordset("cl_zona") = "NO CONFORME"
                     End If
                  End If
               End If
'               data_inf.Recordset("cl_nomcobr") = data_reg.Recordset("cl_nomcobr")
               data_inf.Recordset.Update
               data_reg.Recordset.MoveNext
            
            Else
               If IsNull(data_reg.Recordset("cl_fultmov")) = False Then
                  data_reg.Recordset.MoveNext
               Else
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_fnac") = data_reg.Recordset("cl_fnac")
                  data_inf.Recordset("cl_ruc") = data_reg.Recordset("cl_ruc")
                  data_inf.Recordset("info_debit") = Mid(data_reg.Recordset("info_debit"), 1, 130)
                  data_inf.Recordset("cl_descpag") = Mid(data_reg.Recordset("cl_descpag"), 1, 25)
                  data_inf.Recordset("cl_nrovend") = data_reg.Recordset("cl_nrovend")
                  data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("cl_nom_sup")
                  data_inf.Recordset.Update
                  data_reg.Recordset.MoveNext
               End If
            End If
         Loop
         data_inf.RecordSource = "Select * from infcli"
         data_inf.Refresh
         frm_infservap.MousePointer = 0
         cr1.ReportFileName = App.path & "\infregnor.rpt"
         cr1.ReportTitle = "REGISTROS SIN CUMPLIR DESDE: " & mfdd.Text & " HASTA: " & mfhh.Text
         cr1.Action = 1
      
      End If
      If Option3.Value = True Then
         Do While Not data_reg.Recordset.EOF
            If Combo1.ListIndex = 1 Then
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_fnac") = data_reg.Recordset("fecha")
               data_inf.Recordset("cl_ruc") = data_reg.Recordset("hora")
               data_inf.Recordset("info_debit") = Mid(data_reg.Recordset("detalle"), 1, 130)
               data_inf.Recordset("cl_descpag") = Mid(data_reg.Recordset("soc_nom"), 1, 25)
               data_inf.Recordset("cl_nrovend") = data_reg.Recordset("base")
               data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("usuario")
               data_inf.Recordset("cl_fultmov") = data_reg.Recordset("finfecha")
               data_inf.Recordset("cl_fax") = data_reg.Recordset("finhora")
               If IsNull(data_reg.Recordset("finobs")) = False Then
                  data_inf.Recordset("cl_email") = Mid(data_reg.Recordset("finobs"), 1, 80)
               End If
               If Val(data_reg.Recordset("confop")) = 0 Then
                  data_inf.Recordset("cl_zona") = "CONFORME"
               Else
                  If Val(data_reg.Recordset("confop")) = 1 Then
                     data_inf.Recordset("cl_zona") = "CON DEMORA"
                  Else
                     If Val(data_reg.Recordset("confop")) = 2 Then
                        data_inf.Recordset("cl_zona") = "NO CONFORME"
                     End If
                  End If
               End If
'               data_inf.Recordset("cl_nomcobr") = data_reg.Recordset("cl_nomcobr")
               data_inf.Recordset.Update
               data_reg.Recordset.MoveNext
            
            Else
                data_inf.Recordset.AddNew
                data_inf.Recordset("cl_fnac") = data_reg.Recordset("cl_fnac")
                data_inf.Recordset("cl_ruc") = data_reg.Recordset("cl_ruc")
                data_inf.Recordset("info_debit") = Mid(data_reg.Recordset("info_debit"), 1, 130)
                data_inf.Recordset("cl_descpag") = Mid(data_reg.Recordset("cl_descpag"), 1, 25)
                data_inf.Recordset("cl_nrovend") = data_reg.Recordset("cl_nrovend")
                data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("cl_nom_sup")
                data_inf.Recordset("cl_fultmov") = data_reg.Recordset("cl_fultmov")
                data_inf.Recordset("cl_fax") = data_reg.Recordset("cl_fax")
                data_inf.Recordset("cl_email") = Mid(data_reg.Recordset("cl_email"), 1, 80)
                data_inf.Recordset("cl_zona") = data_reg.Recordset("cl_zona")
                data_inf.Recordset("cl_nomcobr") = data_reg.Recordset("cl_nomcobr")
                data_inf.Recordset.Update
                data_reg.Recordset.MoveNext
            End If
         Loop
         frm_infservap.MousePointer = 0
         data_inf.RecordSource = "Select * from infcli"
         data_inf.Refresh
         cr1.ReportFileName = App.path & "\infregreat.rpt"
         cr1.ReportTitle = "REGISTROS TOTALES DESDE: " & mfdd.Text & " HASTA: " & mfhh.Text
         cr1.Action = 1
      
      End If
      frm_infservap.MousePointer = 0
   End If
Else
   MsgBox "Debe ingresar fechas"
End If

frm_infservap.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

'data_reg.DatabaseName = App.Path & "\sapp.mdb"
data_reg.Connect = "ODBC;DSN=sappespecial;"

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

Private Sub mfdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfhh.SetFocus
End If

End Sub
