VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infpsoc 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes solicitudes a Padr�n Social"
   ClientHeight    =   4245
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
   Icon            =   "frm_infpsoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
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
      Top             =   2880
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
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
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
      Picture         =   "frm_infpsoc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infpsoc.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Procesar"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Opciones de informe"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Todas las asistencias"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Asistencias pendientes"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Asistencias realizadas"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
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
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "RANGO de FECHAS:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1560
      Picture         =   "frm_infpsoc.frx":0F56
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1455
   End
End
Attribute VB_Name = "frm_infpsoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frm_infreg.MousePointer = 11
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

Command1.Enabled = False
MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh

If mfdd.Text <> "__/__/____" And mfhh.Text <> "__/__/____" Then
   If Option1.Value = True Then
      If WElusuario = "PALONZO" Or WElusuario = "JFERNAN" Or WElusuario = "NVANEZCA" Or WElusuario = "MCOSTA" Or WElusuario = "MSANCHEZ" Or WElusuario = "NOELIAR" Or XWeltipoU = "USUARIOS ADM" Then
         data_reg.RecordSource = "Select * from env_soc where cl_ter_vto =" & 6 & " And cl_fultmov >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and cl_fultmov <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "#"
         data_reg.Refresh
      Else
         data_reg.RecordSource = "Select * from env_soc where cl_ter_vto =" & 6 & " And cl_fultmov >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and cl_fultmov <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "# and cl_nom_sup ='" & WElusuario & "'"
         data_reg.Refresh
      End If
   Else
      If WElusuario = "PALONZO" Or WElusuario = "JFERNAN" Or WElusuario = "NVANEZCA" Or WElusuario = "MCOSTA" Or WElusuario = "MSANCHEZ" Or WElusuario = "NOELIAR" Or XWeltipoU = "USUARIOS ADM" Then
         data_reg.RecordSource = "Select * from env_soc where cl_ter_vto =" & 6 & " And cl_fnac >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "#"
         data_reg.Refresh
      Else
         data_reg.RecordSource = "Select * from env_soc where cl_ter_vto =" & 6 & " And cl_fnac >=#" & Format(mfdd.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mfhh.Text, "yyyy/mm/dd") & "# and cl_nom_sup ='" & WElusuario & "'"
         data_reg.Refresh
      End If
   End If
   If data_reg.Recordset.RecordCount > 0 Then
      data_reg.Recordset.MoveFirst
      If Option1.Value = True Then
         Do While Not data_reg.Recordset.EOF
            If IsNull(data_reg.Recordset("cl_fultmov")) = True Then
               data_reg.Recordset.MoveNext
            Else
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_fnac") = data_reg.Recordset("cl_fnac")
               data_inf.Recordset("cl_ruc") = data_reg.Recordset("cl_ruc")
               data_inf.Recordset("info_debit") = data_reg.Recordset("info_debit")
               data_inf.Recordset("cl_descpag") = data_reg.Recordset("cl_descpag")
               data_inf.Recordset("cl_nrovend") = data_reg.Recordset("cl_nrovend")
               data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("cl_nom_sup")
               data_inf.Recordset("cl_fultmov") = data_reg.Recordset("cl_fultmov")
               data_inf.Recordset("cl_fax") = data_reg.Recordset("cl_fax")
               data_inf.Recordset("cl_email") = data_reg.Recordset("cl_email")
               data_inf.Recordset("cl_zona") = data_reg.Recordset("cl_zona")
               data_inf.Recordset("cl_nomcobr") = data_reg.Recordset("cl_nomcobr")
               data_inf.Recordset.Update
               data_reg.Recordset.MoveNext
            End If
         Loop
         data_inf.RecordSource = "Select * from infcli"
         data_inf.Refresh
         cr1.ReportFileName = App.path & "\infregrea.rpt"
         cr1.ReportTitle = "SOLICITUDES CUMPLIDAS POR PADRON SOCIAL A USUARIOS DESDE: " & mfdd.Text & " HASTA: " & mfhh.Text
         cr1.Action = 1
         
      End If
      If Option2.Value = True Then
         Do While Not data_reg.Recordset.EOF
            If IsNull(data_reg.Recordset("cl_fultmov")) = False Then
               data_reg.Recordset.MoveNext
            Else
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_fnac") = data_reg.Recordset("cl_fnac")
               data_inf.Recordset("cl_ruc") = data_reg.Recordset("cl_ruc")
               data_inf.Recordset("info_debit") = data_reg.Recordset("info_debit")
               data_inf.Recordset("cl_descpag") = data_reg.Recordset("cl_descpag")
               data_inf.Recordset("cl_nrovend") = data_reg.Recordset("cl_nrovend")
               data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("cl_nom_sup")
               data_inf.Recordset.Update
               data_reg.Recordset.MoveNext
            End If
         Loop
         data_inf.RecordSource = "Select * from infcli"
         data_inf.Refresh
         cr1.ReportFileName = App.path & "\infregnor.rpt"
         cr1.ReportTitle = "SOLICITUDES SIN CUMPLIR POR PADRON SOCIAL DESDE: " & mfdd.Text & " HASTA: " & mfhh.Text
         cr1.Action = 1
      
      End If
      If Option3.Value = True Then
         Do While Not data_reg.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_fnac") = data_reg.Recordset("cl_fnac")
            data_inf.Recordset("cl_ruc") = data_reg.Recordset("cl_ruc")
            data_inf.Recordset("info_debit") = data_reg.Recordset("info_debit")
            data_inf.Recordset("cl_descpag") = data_reg.Recordset("cl_descpag")
            data_inf.Recordset("cl_nrovend") = data_reg.Recordset("cl_nrovend")
            data_inf.Recordset("cl_nom_sup") = data_reg.Recordset("cl_nom_sup")
            data_inf.Recordset("cl_fultmov") = data_reg.Recordset("cl_fultmov")
            data_inf.Recordset("cl_fax") = data_reg.Recordset("cl_fax")
            data_inf.Recordset("cl_email") = data_reg.Recordset("cl_email")
            data_inf.Recordset("cl_zona") = data_reg.Recordset("cl_zona")
            data_inf.Recordset("cl_nomcobr") = data_reg.Recordset("cl_nomcobr")
            data_inf.Recordset.Update
            data_reg.Recordset.MoveNext
         Loop
         data_inf.RecordSource = "Select * from infcli"
         data_inf.Refresh
         cr1.ReportFileName = App.path & "\infregreat.rpt"
         cr1.ReportTitle = "SOLICITUDES TOTALES REALIZADAS A PADRON SOCIAL DESDE: " & mfdd.Text & " HASTA: " & mfhh.Text
         cr1.Action = 1
      
      End If
      frm_infreg.MousePointer = 0
   End If
Else
   MsgBox "Debe ingresar fechas"
End If
Command1.Enabled = True
frm_infreg.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

data_reg.Connect = "odbc;dsn=" & Xconexrmt & ";"
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
