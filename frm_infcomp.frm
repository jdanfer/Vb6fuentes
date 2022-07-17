VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infcomp 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de compras"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infcomp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "FORMATO"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   5895
      Begin VB.OptionButton Option4 
         BackColor       =   &H00800000&
         Caption         =   "RESUMEN"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00800000&
         Caption         =   "DETALLE"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Data data_cab 
      Caption         =   "data_cab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_comp 
      Caption         =   "data_comp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3360
      Top             =   3720
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
      Left            =   5520
      Picture         =   "frm_infcomp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infcomp.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.OptionButton Option2 
         BackColor       =   &H00800000&
         Caption         =   "Por laboratorio/comercio"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Por ITEM"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Value           =   -1  'True
         Width           =   2895
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   5880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "FECHAS:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   3840
      Picture         =   "frm_infcomp.frx":0F56
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "frm_infcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_infcomp.MousePointer = 11

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"

data_inf.RecordSource = "infvtas"
data_inf.Refresh

If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
         data_comp.RecordSource = "Select * from lineascomp where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and grupo in (3)"
      Else
         data_comp.RecordSource = "Select * from lineascomp where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and grupo not in (3)"
      End If
      data_comp.Refresh
      If data_comp.Recordset.RecordCount > 0 Then
         data_comp.Recordset.MoveFirst
         Do While Not data_comp.Recordset.EOF
            data_cab.RecordSource = "Select * from compras where nrobol =" & data_comp.Recordset("codbol")
            data_cab.Refresh
            If data_cab.Recordset.RecordCount > 0 Then
               data_inf.Recordset.AddNew
               data_inf.Recordset("fecha") = data_comp.Recordset("fecha")
               data_inf.Recordset("cod_prod") = data_comp.Recordset("codprod")
               data_inf.Recordset("nom_prod") = Mid(data_comp.Recordset("coddesc"), 1, 50)
               data_inf.Recordset("nom_cli") = Mid(data_cab.Recordset("nomcomer"), 1, 30)
               data_inf.Recordset("cod_cli") = data_cab.Recordset("codcomer")
               data_inf.Recordset("costo_prod") = data_comp.Recordset("precuni")
               data_inf.Recordset("cantidad") = data_comp.Recordset("cant")
               data_inf.Recordset("imp_iva") = 0
               data_inf.Recordset("tot_lin") = data_comp.Recordset("totprod")
               data_inf.Recordset.Update
            End If
            data_comp.Recordset.MoveNext
         Loop
         frm_infcomp.MousePointer = 0
         MsgBox "Terminado"
         If Option1.value = True Then
            If Option3.value = True Then
               cr1.ReportFileName = App.Path & "\infcompit.rpt"
            Else
               cr1.ReportFileName = App.Path & "\infcompitn.rpt"
            End If
            cr1.ReportTitle = "Informe de compras por productos desde: " & mfd.Text & " Hasta: " & mfh.Text
            cr1.Action = 1
         End If
         If Option2.value = True Then
            cr1.ReportFileName = App.Path & "\infcompla.rpt"
            cr1.ReportTitle = "Informe de compras por comercio desde: " & mfd.Text & " Hasta: " & mfh.Text
            cr1.Action = 1
         End If
      End If
   End If
End If
frm_infcomp.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "infvtas"
data_inf.Refresh
data_comp.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cab.Connect = "odbc;dsn=" & Xconexrmt & ";"


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
