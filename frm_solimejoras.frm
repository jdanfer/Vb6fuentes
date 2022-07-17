VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_solimejoras 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de solicitudes de mejora"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_solimejoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5310
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1200
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1920
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4560
      Picture         =   "frm_solimejoras.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_solimejoras.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   3720
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Datos para el informe"
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Ordenar por SOLICITANTE"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   3255
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Solicitudes pendientes"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Solicitues terminadas"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Total de solicitudes"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   3255
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   480
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
         BackColor       =   &H0080FFFF&
         Caption         =   "FECHAS:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1680
      Picture         =   "frm_solimejoras.frx":0F56
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frm_solimejoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

If mfd.Text <> "__/__/____" Then
   If Option1.value = True Then
      Data1.RecordSource = "Select * from solaudito where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
      Data1.Refresh
   Else
      If Option2.value = True Then
         Data1.RecordSource = "Select * from solaudito where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and val1 =" & 1
         Data1.Refresh
      Else
         If Option3.value = True Then
            Data1.RecordSource = "Select * from solaudito where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and val1 <" & 1
            Data1.Refresh
         Else
            Data1.RecordSource = "Select * from solaudito where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
            Data1.Refresh
         End If
      End If
   End If
   data_inf.RecordSource = "infcli"
   data_inf.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_codigo") = Data1.Recordset("cl_codigo")
         data_inf.Recordset("cl_fnac") = Data1.Recordset("cl_fnac")
         data_inf.Recordset("cl_codced") = Val(Data1.Recordset("cl_etiquet"))
         If Val(Data1.Recordset("cl_etiquet")) = 0 Then
            data_inf.Recordset("cl_nombre") = "AMOBLAMIENTOS"
         Else
            If Val(Data1.Recordset("cl_etiquet")) = 1 Then
               data_inf.Recordset("cl_nombre") = "EQUIPOS"
            Else
               If Val(Data1.Recordset("cl_etiquet")) = 2 Then
                  data_inf.Recordset("cl_nombre") = "INSTALACIONES"
               Else
                  If Val(Data1.Recordset("cl_etiquet")) = 3 Then
                     data_inf.Recordset("cl_nombre") = "OTROS"
                  Else
                     If Val(Data1.Recordset("cl_etiquet")) = 4 Then
                        data_inf.Recordset("cl_nombre") = "PROCEDIMIENTO"
                     Else
                        If Val(Data1.Recordset("cl_etiquet")) = 5 Then
                           data_inf.Recordset("cl_nombre") = "SUGERENCIA"
                        Else
                           data_inf.Recordset("cl_nombre") = "S/R"
                        End If
                     End If
                  End If
               End If
            End If
         End If
         data_inf.Recordset("info_debit") = Mid(Data1.Recordset("info_debit"), 1, 130)
         data_inf.Recordset("cl_nomvend") = Mid(Data1.Recordset("cl_nomvend"), 1, 25)
         data_inf.Recordset("cl_cedula") = Data1.Recordset("cl_schqmn")
         If IsNull(Data1.Recordset("cl_schqmn")) = False Then
            If Data1.Recordset("cl_schqmn") = 0 Then
               data_inf.Recordset("cl_nomcobr") = "ACEPTADO"
            Else
               If Data1.Recordset("cl_schqmn") = 1 Then
                  data_inf.Recordset("cl_nomcobr") = "A ESTUDIO"
               Else
                  If Data1.Recordset("cl_schqmn") = 2 Then
                     data_inf.Recordset("cl_nomcobr") = "NO ACEPTADO"
                  Else
                     data_inf.Recordset("cl_nomcobr") = "S/R"
                  End If
               End If
            End If
         Else
            data_inf.Recordset("cl_nomcobr") = "S/R"
         End If
         data_inf.Recordset("cl_nrovend") = Data1.Recordset("val1")
         If IsNull(Data1.Recordset("val1")) = False Then
            If Data1.Recordset("val1") = 1 Then
               data_inf.Recordset("cl_telefon") = "TERMINADO"
            Else
               data_inf.Recordset("cl_telefon") = "PENDIENTE"
            End If
         Else
            data_inf.Recordset("cl_telefon") = "PENDIENTE"
         End If
         data_inf.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
      MsgBox "Proceso Terminado"
      
      data_inf.RecordSource = "Select * from infcli"
      data_inf.Refresh
      If Check1.value = 1 Then
         cr1.ReportFileName = App.Path & "\infsolimejor2.rpt"
      Else
         cr1.ReportFileName = App.Path & "\infsolimejor.rpt"
      End If
      If Option1.value = True Then
         cr1.ReportTitle = "Informe de solicitudes de Mejora (TODAS). FECHA: " & mfd.Text & " HASTA: " & mfh.Text
      Else
         If Option2.value = True Then
            cr1.ReportTitle = "Informe de solicitudes de Mejora (TERMINADAS). FECHA: " & mfd.Text & " HASTA: " & mfh.Text
         Else
            If Option3.value = True Then
               cr1.ReportTitle = "Informe de solicitudes de Mejora (PENDIENTES). FECHA: " & mfd.Text & " HASTA: " & mfh.Text
            Else
               cr1.ReportTitle = "Informe de solicitudes de Mejora. FECHA: " & mfd.Text & " HASTA: " & mfh.Text
            End If
         End If
      End If
      cr1.Action = 1
   End If
Else
   MsgBox "Ingrese fecha"
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.Path & "\informes.mdb"

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
