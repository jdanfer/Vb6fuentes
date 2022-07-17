VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infmetasnew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de METAS"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7005
   Icon            =   "frm_infmetasnew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frm_infmetasnew.frx":058A
   ScaleHeight     =   4260
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   2760
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   5280
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      Picture         =   "frm_infmetasnew.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      Picture         =   "frm_infmetasnew.frx":13DE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Datos para el informe"
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
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_infmetasnew.frx":1968
         Left            =   2160
         List            =   "frm_infmetasnew.frx":197E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "Informe de NO realizadas"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   4575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grupo Mutual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seleccione Meta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FECHAS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label labd 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label labm 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label laba 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   5160
      Picture         =   "frm_infmetasnew.frx":19BE
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frm_infmetasnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Regcli As New ADODB.Recordset
Public Regmeta As New ADODB.Recordset
Public Regbusca As New ADODB.Recordset
Public Sqlcli, Sqlmeta, Sqlconsmeta As String

Private Sub Command1_Click()
Dim Xlin, XCol As Long
Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet
Dim Xced As Long
Dim Xcantced, Xcanttot, Xnrocan As Integer
Xcantced = 0
Xcanttot = 0
Xnrocan = 1
Xlin = 1
XCol = 1
Dim Xlabrir As New Excel.Application
Dim Xarchtex As String

Command1.Enabled = False
Command2.Enabled = False
frm_infmetasnew.MousePointer = 99

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data1.Recordset.Delete
      Data1.Recordset.MoveNext
   Loop
End If
If mfd.Text <> "__/__/____" And mfh.Text <> "__/__/____" Then
   If Combo1.Text = "CTROL.RECIEN NACIDO" Or _
      Combo1.Text = "CTROL.1ER.AÑO DE VIDA" Or _
      Combo1.Text = "CTROL.2DO.AÑO DE VIDA" Or _
      Combo1.Text = "CTROL.3ER.AÑO DE VIDA" Or _
      Combo1.Text = "CTROL.4TO.AÑO DE VIDA" Or _
      Combo1.Text = "CTROL.5TO.AÑO DE VIDA" Then
      If Check1.value = 1 Then
         If Combo2.ListIndex >= 1 Then
            ConectarBD
            ConbdSapp.Open
            Sqlmeta = "Select * from clientes where cl_fnac is not null and estado =" & 1
            With Regmeta
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlmeta, ConbdSapp, , , adCmdText
            End With
            If Regmeta.RecordCount > 0 Then
               Regmeta.MoveFirst
               Do While Not Regmeta.EOF
                  Sqlconsmeta = "Select * from convenio where cnv_codigo ='" & Regmeta("cl_codconv") & "'"
                  With Regbusca
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockOptimistic
                      .Open Sqlconsmeta, ConbdSapp, , , adCmdText
                  End With
                  If Regbusca.RecordCount > 0 Then
                     If IsNull(Regbusca("cnv_grupo")) = False Then
                        If Regbusca("cnv_grupo") = Combo2.Text Then
                           Data1.Recordset.AddNew
                           Data1.Recordset("cl_codigo") = Regmeta("cl_codigo")
                           Data1.Recordset("cl_apellid") = Regmeta("cl_apellid")
                           Data1.Recordset("cl_cedula") = Regmeta("cl_cedula")
                           Data1.Recordset("cl_codced") = Regmeta("cl_codced")
                           Data1.Recordset("cl_codconv") = Regmeta("cl_codconv")
                           Data1.Recordset("cl_nomconv") = Regmeta("cl_nomconv")
                           Data1.Recordset("cl_fnac") = Regmeta("cl_fnac")
                           Data1.Recordset("cl_fecing") = Regmeta("cl_fecing")
                           Data1.Recordset("cl_telefon") = Regmeta("cl_telefon")
                           Data1.Recordset("cl_dpto") = Regmeta("cl_dpto")
                           Data1.Recordset("cl_zona") = Regmeta("cl_zona")
                           Data1.Recordset.Update
                        End If
                     End If
                  End If
                  Regbusca.Close
                  Regmeta.MoveNext
               Loop
               Regmeta.Close
               Data1.Refresh
               If Data1.Recordset.RecordCount > 0 Then
                  Data1.Recordset.MoveFirst
                  Do While Not Data1.Recordset.EOF
                     CalculaEdad (Data1.Recordset("cl_fnac"))
                     If Val(laba.Caption) = 0 Then
                        If Val(labm.Caption) = 0 Then
                           If Val(labd.Caption) < 10 Then
                              Sqlmeta = "Select * from linmmdd where cod_prod =" & 190001 & " and cod_cli =" & Data1.Recordset("cl_codigo")
                              With Regmeta
                                  .CursorLocation = adUseClient
                                  .CursorType = adOpenKeyset
                                  .LockType = adLockOptimistic
                                  .Open Sqlmeta, ConbdSapp, , , adCmdText
                              End With
                              If Regmeta.RecordCount > 0 Then
                                 Data1.Recordset.Delete
                              Else
                                 Data1.Recordset.Edit
                                 Data1.Recordset("cl_dircobr") = "META 1 -CAPTACION RECIEN NACIDO"
                                 Data1.Recordset.Update
                              End If
                              Regmeta.Close
                           Else
                              Sqlmeta = "Select * from linmmdd where cod_prod =" & 190003 & " and cod_cli =" & Data1.Recordset("cl_codigo")
                              With Regmeta
                                  .CursorLocation = adUseClient
                                  .CursorType = adOpenKeyset
                                  .LockType = adLockOptimistic
                                  .Open Sqlmeta, ConbdSapp, , , adCmdText
                              End With
                              If Regmeta.RecordCount > 0 Then
                                 Data1.Recordset.Delete
                              Else
                                 Data1.Recordset.Edit
                                 Data1.Recordset("cl_dircobr") = "META 1 -CONTROL NIÑO 1ER.AÑO"
                                 Data1.Recordset.Update
                              End If
                              Regmeta.Close
                           End If
                        Else
                           Sqlmeta = "Select * from linmmdd where cod_prod =" & 190003 & " and cod_cli =" & Data1.Recordset("cl_codigo")
                           With Regmeta
                               .CursorLocation = adUseClient
                               .CursorType = adOpenKeyset
                               .LockType = adLockOptimistic
                               .Open Sqlmeta, ConbdSapp, , , adCmdText
                           End With
                           If Regmeta.RecordCount > 0 Then
                              Data1.Recordset.Delete
                           Else
                              Data1.Recordset.Edit
                              Data1.Recordset("cl_dircobr") = "META 1 -CONTROL NIÑO 1ER.AÑO"
                              Data1.Recordset.Update
                           End If
                           Regmeta.Close
                        End If
                     Else
                        If Val(laba.Caption) = 1 Then
                           Sqlmeta = "Select * from linmmdd where cod_prod =" & 190003 & " and cod_cli =" & Data1.Recordset("cl_codigo")
                           With Regmeta
                               .CursorLocation = adUseClient
                               .CursorType = adOpenKeyset
                               .LockType = adLockOptimistic
                               .Open Sqlmeta, ConbdSapp, , , adCmdText
                           End With
                           If Regmeta.RecordCount > 0 Then
                              Data1.Recordset.Delete
                           Else
                              Data1.Recordset.Edit
                              Data1.Recordset("cl_dircobr") = "META 1 -CONTROL NIÑO 1ER.AÑO"
                              Data1.Recordset.Update
                           End If
                           Regmeta.Close
                        Else
                           If Val(laba.Caption) = 2 Then
                              Sqlmeta = "Select * from linmmdd where cod_prod =" & 190004 & " and cod_cli =" & Data1.Recordset("cl_codigo")
                              With Regmeta
                                  .CursorLocation = adUseClient
                                  .CursorType = adOpenKeyset
                                  .LockType = adLockOptimistic
                                  .Open Sqlmeta, ConbdSapp, , , adCmdText
                              End With
                              If Regmeta.RecordCount > 0 Then
                                 Data1.Recordset.Delete
                              Else
                                 Data1.Recordset.Edit
                                 Data1.Recordset("cl_dircobr") = "META 1 -CONTROL NIÑO 2DO.AÑO"
                                 Data1.Recordset.Update
                              End If
                              Regmeta.Close
                           Else
                              If Val(laba.Caption) = 3 Then
                                 Sqlmeta = "Select * from linmmdd where cod_prod =" & 190005 & " and cod_cli =" & Data1.Recordset("cl_codigo")
                                 With Regmeta
                                     .CursorLocation = adUseClient
                                     .CursorType = adOpenKeyset
                                     .LockType = adLockOptimistic
                                     .Open Sqlmeta, ConbdSapp, , , adCmdText
                                 End With
                                 If Regmeta.RecordCount > 0 Then
                                    Data1.Recordset.Delete
                                 Else
                                    Data1.Recordset.Edit
                                    Data1.Recordset("cl_dircobr") = "META 1 -CONTROL NIÑO 3ER.AÑO"
                                    Data1.Recordset.Update
                                 End If
                                 Regmeta.Close
                              Else
                                 If Val(laba.Caption) = 4 Then
                                    Sqlmeta = "Select * from linmmdd where cod_prod =" & 190030 & " and cod_cli =" & Data1.Recordset("cl_codigo")
                                    With Regmeta
                                        .CursorLocation = adUseClient
                                        .CursorType = adOpenKeyset
                                        .LockType = adLockOptimistic
                                        .Open Sqlmeta, ConbdSapp, , , adCmdText
                                    End With
                                    If Regmeta.RecordCount > 0 Then
                                       Data1.Recordset.Delete
                                    Else
                                       Data1.Recordset.Edit
                                       Data1.Recordset("cl_dircobr") = "META 1 -CONTROL NIÑO 4TO.AÑO"
                                       Data1.Recordset.Update
                                    End If
                                    Regmeta.Close
                                 Else
                                    If Val(laba.Caption) = 5 Then
                                       Sqlmeta = "Select * from linmmdd where cod_prod =" & 190031 & " and cod_cli =" & Data1.Recordset("cl_codigo")
                                       With Regmeta
                                           .CursorLocation = adUseClient
                                           .CursorType = adOpenKeyset
                                           .LockType = adLockOptimistic
                                           .Open Sqlmeta, ConbdSapp, , , adCmdText
                                       End With
                                       If Regmeta.RecordCount > 0 Then
                                          Data1.Recordset.Delete
                                       Else
                                          Data1.Recordset.Edit
                                          Data1.Recordset("cl_dircobr") = "META 1 -CONTROL NIÑO 5TO.AÑO"
                                          Data1.Recordset.Update
                                       End If
                                       Regmeta.Close
                                    Else
                                       Data1.Recordset.Delete
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                     Data1.Recordset.MoveNext
                  Loop
                  Data1.RecordSource = "Select * from infcli order by cl_codigo"
                  Data1.Refresh
                  frm_infmetasnew.MousePointer = 0
                  MsgBox "Terminado"
                  cr1.ReportFileName = App.Path & "\infmetasnew.rpt"
                  cr1.ReportTitle = "Informe de socios ACTIVOS con METAS PENDIENTES"
                  cr1.Action = 1
               End If
            Else
               frm_infmetasnew.MousePointer = 0
               MsgBox "No hay registros"
            End If
            ConbdSapp.Close
         Else
            frm_infmetasnew.MousePointer = 0
            MsgBox "Debe seleccionar una mutualista", vbInformation
         End If
      Else
        ConectarBDM
        ConbdSappM.Open
        If Combo1.Text = "TODAS" Then
           Sqlcli = "Select * from t_meta1 where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by cedula,fecha"
        Else
           Sqlcli = "Select * from t_meta1 where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and meta2desc ='" & Combo1.Text & "' order by cedula,fecha"
        End If
        With Regcli
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open Sqlcli, ConbdSappM, , , adCmdText
        End With
        If Regcli.RecordCount > 0 Then
           Set Xobjexel = New Excel.Application
           Set Xlibexel = Xobjexel.Workbooks.Add
           Set Xarchexel = Xlibexel.Worksheets.Add
           Xarchexel.Name = Trim(Combo1.Text)
           Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo2.Text) & ".xls")
           Xarchtex = "C:\planillas\" & Trim(Combo2.Text) & ".xls"
           Regcli.MoveFirst
           Do While Not Regcli.EOF
              Data1.Recordset.AddNew
              Data1.Recordset("cl_fecing") = Format(Regcli("fecha"), "dd/mm/yyyy")
              Data1.Recordset("cl_codigo") = Val(Regcli("matric"))
              Data1.Recordset("cl_cedula") = Val(Regcli("cedula"))
              Data1.Recordset("cl_codced") = Val(Regcli("codced"))
              Data1.Recordset("cl_apellid") = Regcli("nombre")
              Data1.Recordset("cl_nombre") = Regcli("meta2desc")
              Data1.Recordset("cl_nrocobr") = Val(Regcli("base"))
              Data1.Recordset("cl_fnac") = Format(Regcli("fecctrl"), "dd/mm/yyyy")
              Data1.Recordset("cl_direcci") = Regcli("edadtex2")
              If IsNull(Regcli("peso")) = False Then
                 Data1.Recordset("cl_fax") = Mid(Regcli("peso"), 1, 10)
              End If
              
              If IsNull(Regcli("lactan")) = False Then
                 Data1.Recordset("cl_localid") = Mid(Regcli("lactan"), 1, 35)
              End If
              If IsNull(Regcli("vacuna")) = False Then
                 Data1.Recordset("cl_nomvend") = Mid(Regcli("vacuna"), 1, 35)
              End If
              If IsNull(Regcli("ecocad")) = False Then
                 Data1.Recordset("cl_zona") = Mid(Regcli("ecocad"), 1, 25)
              End If
              If IsNull(Regcli("fecprox")) = False Then
                 Data1.Recordset("cl_nom_sup") = Mid(Regcli("fecprox"), 1, 25)
              End If
              If IsNull(Regcli("medico")) = False Then
                 Data1.Recordset("cl_medflia") = Mid(Regcli("medico"), 1, 25)
              End If
              If IsNull(Regcli("obs")) = False Then
                 Data1.Recordset("cl_entre") = Mid(Regcli("obs"), 1, 80)
              End If
              Data1.Recordset("cl_codconv") = Regcli("cnvcod")
              Data1.Recordset.Update
              Regcli.MoveNext
           Loop
           Data1.Refresh
           If Combo2.Text = "TODAS" Then
           Else
              If Data1.Recordset.RecordCount > 0 Then
                 ConectarBD
                 ConbdSapp.Open
                 Data1.Recordset.MoveFirst
                 Do While Not Data1.Recordset.EOF
                    Sqlmeta = "Select * from convenio where cnv_codigo ='" & Data1.Recordset("cl_codconv") & "'"
                    With Regmeta
                        .CursorLocation = adUseClient
                        .CursorType = adOpenKeyset
                        .LockType = adLockOptimistic
                        .Open Sqlmeta, ConbdSapp, , , adCmdText
                    End With
                    If Regmeta.RecordCount > 0 Then
                       If IsNull(Regmeta("cnv_grupo")) = False Then
                          If Regmeta("cnv_grupo") = Combo2.Text Then
                          Else
                             Data1.Recordset.Delete
                          End If
                       Else
                          Data1.Recordset.Delete
                       End If
                    Else
                       Data1.Recordset.Delete
                    End If
                    Data1.Recordset.MoveNext
                 Loop
                 Data1.Refresh
                 Regmeta.Close
                 ConbdSapp.Close
              End If
           End If
           If Data1.Recordset.RecordCount > 0 Then
              Data1.Recordset.MoveFirst
              Xarchexel.Cells(Xlin, XCol) = "CENTRO DE COMPUTOS DE SAPP"
              Xlin = Xlin + 1
              XCol = XCol + 1
              Xarchexel.Range("A1", "C3").Font.Size = 16
              Xarchexel.Cells(Xlin, XCol) = "PLANILLA DE " & Combo2.Text & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
              Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
              XCol = 1
              Xlin = Xlin + 2
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
              Xarchexel.Range("A" & Trim(Str(Xlin)), "O" & Trim(Str(Xlin))).Interior.color = RGB(215, 120, 120)
              Xarchexel.Range("A" & Trim(Str(Xlin))).ColumnWidth = 5
              Xarchexel.Cells(Xlin, XCol) = "BASE" 'a
              XCol = XCol + 1
              Xarchexel.Range("B" & Trim(Str(Xlin))).ColumnWidth = 35
              Xarchexel.Cells(Xlin, XCol) = "NOMBRE" 'b
              XCol = XCol + 1
              Xarchexel.Range("C" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "CEDULA" 'c
              XCol = XCol + 1
              Xarchexel.Range("D" & Trim(Str(Xlin))).ColumnWidth = 4
              Xarchexel.Cells(Xlin, XCol) = "NRO.CTROL." 'd
              XCol = XCol + 1
              Xarchexel.Cells(Xlin, XCol) = "FECHA CTROL." 'd
              XCol = XCol + 1
              Xarchexel.Range("F" & Trim(Str(Xlin))).ColumnWidth = 25
              Xarchexel.Cells(Xlin, XCol) = "EDAD" 'f
              XCol = XCol + 1
              Xarchexel.Cells(Xlin, XCol) = "PESO" 'g
              XCol = XCol + 1
              Xarchexel.Range("H" & Trim(Str(Xlin))).ColumnWidth = 20
              Xarchexel.Cells(Xlin, XCol) = "ALIMENTACION"
              XCol = XCol + 1
              Xarchexel.Range("I" & Trim(Str(Xlin))).ColumnWidth = 20
              Xarchexel.Cells(Xlin, XCol) = "VACUNAS"
              XCol = XCol + 1
              Xarchexel.Range("J" & Trim(Str(Xlin))).ColumnWidth = 20
              Xarchexel.Cells(Xlin, XCol) = "ECOGRAFIA CADERA"
              XCol = XCol + 1
              Xarchexel.Range("K" & Trim(Str(Xlin))).ColumnWidth = 20
              Xarchexel.Cells(Xlin, XCol) = "PROX.CONTROL"
              XCol = XCol + 1
              Xarchexel.Range("L" & Trim(Str(Xlin))).ColumnWidth = 20
              Xarchexel.Cells(Xlin, XCol) = "MEDICO"
              XCol = XCol + 1
              Xarchexel.Range("M" & Trim(Str(Xlin))).ColumnWidth = 30
              Xarchexel.Cells(Xlin, XCol) = "OBSERVACIONES"
              Xced = Data1.Recordset("cl_cedula")
              Xcanttot = 1
              Xlin = Xlin + 1
              Do While Not Data1.Recordset.EOF
                 If Val(Xced) = Val(Data1.Recordset("cl_cedula")) Then
                    If Xcantced > 1 Then
                       XCol = 4
                       Xarchexel.Cells(Xlin, XCol) = Xcanttot
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Format(Data1.Recordset("cl_fnac"), "dd/mm/yyyy")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_direcci")
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_fax")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_fax")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_localid")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_localid")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_nomvend")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nomvend")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_zona")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_zona")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_nom_sup")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nom_sup")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_medflia")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_medflia")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_entre")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_entre")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       Xcantced = 4
                       Xlin = Xlin + 1
                       Data1.Recordset.MoveNext
                       Xcanttot = Xcanttot + 1
                    Else
                       XCol = 1
                       Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nrocobr")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_apellid")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Trim(Str(Data1.Recordset("cl_cedula"))) & "-" & Trim(Str(Data1.Recordset("cl_codced")))
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Xcanttot
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Format(Data1.Recordset("cl_fnac"), "dd/mm/yyyy")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_direcci")
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_fax")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_fax")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_localid")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_localid")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_nomvend")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nomvend")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_zona")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_zona")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_nom_sup")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nom_sup")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_medflia")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_medflia")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       XCol = XCol + 1
                       If IsNull(Data1.Recordset("cl_entre")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_entre")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = ""
                       End If
                       Xcantced = 4
                       Xlin = Xlin + 1
                       Data1.Recordset.MoveNext
                       Xcanttot = Xcanttot + 1
                    End If
                 Else
                    Xcantced = 0
                    Xced = Data1.Recordset("cl_cedula")
                    Xcanttot = 0
                 End If
              Loop
              Xlibexel.Save
              Xlibexel.Close
              Xobjexel.Quit
              Xlabrir.Workbooks.Open Xarchtex, , False
              Xlabrir.Visible = True
              Xlabrir.WindowState = xlMaximized
           Else
              Xlibexel.Close
              Xobjexel.Quit
           End If
        End If
        frm_infmetasnew.MousePointer = 0
        MsgBox "Terminado"
        Regcli.Close
        ConbdSappM.Close
      End If
   Else
        Command3_Click
   End If
End If
frm_infmetasnew.MousePointer = 0

Command1.Enabled = True
Command2.Enabled = True
      

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
If Combo1.Text = "CTROL.45 A 64AÑOS" Or _
   Combo1.Text = "CTROL.65 A 74AÑOS" Or _
   Combo1.Text = "CTROL. >75 AÑOS" Then
   If Check1.value = 1 Then
      If Combo2.ListIndex >= 1 Then
         ConectarBD
         ConbdSapp.Open
         Sqlmeta = "Select * from clientes where cl_fnac is not null and estado =" & 1
         With Regmeta
              .CursorLocation = adUseClient
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open Sqlmeta, ConbdSapp, , , adCmdText
         End With
         If Regmeta.RecordCount > 0 Then
            Regmeta.MoveFirst
            Do While Not Regmeta.EOF
               Sqlconsmeta = "Select * from convenio where cnv_codigo ='" & Regmeta("cl_codconv") & "'"
               With Regbusca
                   .CursorLocation = adUseClient
                   .CursorType = adOpenKeyset
                   .LockType = adLockOptimistic
                   .Open Sqlconsmeta, ConbdSapp, , , adCmdText
               End With
               If Regbusca.RecordCount > 0 Then
                  If IsNull(Regbusca("cnv_grupo")) = False Then
                     If Regbusca("cnv_grupo") = Combo2.Text Then
                        Data1.Recordset.AddNew
                        Data1.Recordset("cl_codigo") = Regmeta("cl_codigo")
                        Data1.Recordset("cl_apellid") = Regmeta("cl_apellid")
                        Data1.Recordset("cl_cedula") = Regmeta("cl_cedula")
                        Data1.Recordset("cl_codced") = Regmeta("cl_codced")
                        Data1.Recordset("cl_codconv") = Regmeta("cl_codconv")
                        Data1.Recordset("cl_nomconv") = Regmeta("cl_nomconv")
                        Data1.Recordset("cl_fnac") = Regmeta("cl_fnac")
                        Data1.Recordset("cl_fecing") = Regmeta("cl_fecing")
                        Data1.Recordset("cl_telefon") = Regmeta("cl_telefon")
                        Data1.Recordset("cl_dpto") = Regmeta("cl_dpto")
                        Data1.Recordset("cl_zona") = Regmeta("cl_zona")
                        Data1.Recordset.Update
                     End If
                  End If
               End If
               Regbusca.Close
               Regmeta.MoveNext
            Loop
            Regmeta.Close
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.MoveFirst
               Do While Not Data1.Recordset.EOF
                  CalculaEdad (Data1.Recordset("cl_fnac"))
                  If Val(laba.Caption) >= 45 And Val(laba.Caption) <= 64 Then
                     Sqlmeta = "Select * from linmmdd where cod_prod in (190014,190028) and cod_cli =" & Data1.Recordset("cl_codigo")
                     With Regmeta
                         .CursorLocation = adUseClient
                         .CursorType = adOpenKeyset
                         .LockType = adLockOptimistic
                         .Open Sqlmeta, ConbdSapp, , , adCmdText
                     End With
                     If Regmeta.RecordCount > 0 Then
                        Data1.Recordset.Delete
                     Else
                        Data1.Recordset.Edit
                        Data1.Recordset("cl_dircobr") = "META 2 -CONTROL 45 A 64 AÑOS"
                        Data1.Recordset.Update
                     End If
                     Regmeta.Close
                  Else
                     If Val(laba.Caption) >= 65 And Val(laba.Caption) <= 74 Then
                        Sqlmeta = "Select * from linmmdd where cod_prod in (190018,190019,190023) and cod_cli =" & Data1.Recordset("cl_codigo")
                        With Regmeta
                            .CursorLocation = adUseClient
                            .CursorType = adOpenKeyset
                            .LockType = adLockOptimistic
                            .Open Sqlmeta, ConbdSapp, , , adCmdText
                        End With
                        If Regmeta.RecordCount > 0 Then
                           Data1.Recordset.Delete
                        Else
                           Data1.Recordset.Edit
                           Data1.Recordset("cl_dircobr") = "META 3 -CONTROL ANUAL 65 A 74 AÑOS"
                           Data1.Recordset.Update
                        End If
                        Regmeta.Close
                     Else
                        If Val(laba.Caption) >= 75 Then
                           Sqlmeta = "Select * from linmmdd where cod_prod in (190020,190021,190022) and cod_cli =" & Data1.Recordset("cl_codigo")
                           With Regmeta
                               .CursorLocation = adUseClient
                               .CursorType = adOpenKeyset
                               .LockType = adLockOptimistic
                               .Open Sqlmeta, ConbdSapp, , , adCmdText
                           End With
                           If Regmeta.RecordCount > 0 Then
                              Data1.Recordset.Delete
                           Else
                              Data1.Recordset.Edit
                              Data1.Recordset("cl_dircobr") = "META 3 -CONTROL >75 AÑOS"
                              Data1.Recordset.Update
                           End If
                           Regmeta.Close
                        Else
                           Data1.Recordset.Delete
                        End If
                     End If
                  End If
                  Data1.Recordset.MoveNext
               Loop
               frm_infmetasnew.MousePointer = 0
               MsgBox "Terminado"
               Data1.RecordSource = "Select * from infcli order by cl_codigo"
               Data1.Refresh
               cr1.ReportFileName = App.Path & "\infmetasnew.rpt"
               cr1.ReportTitle = "Informe de socios ACTIVOS con METAS PENDIENTES"
               cr1.Action = 1
            End If
         Else
            frm_infmetasnew.MousePointer = 0
            MsgBox "No existen registros"
         End If
         ConbdSapp.Close
      Else
         frm_infmetasnew.MousePointer = 0
         MsgBox "Debe seleccionar una mutualista", vbInformation
      End If
   Else
     ConectarBDM
     ConbdSappM.Open
     If Combo1.Text = "TODAS" Then
        Sqlcli = "Select * from t_meta1 where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by cedula,fecha"
     Else
        Sqlcli = "Select * from t_meta1 where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and meta2desc ='" & Combo1.Text & "' order by cedula,fecha"
     End If
     With Regcli
         .CursorLocation = adUseClient
         .CursorType = adOpenKeyset
         .LockType = adLockOptimistic
         .Open Sqlcli, ConbdSappM, , , adCmdText
     End With
     If Regcli.RecordCount > 0 Then
        Set Xobjexel = New Excel.Application
        Set Xlibexel = Xobjexel.Workbooks.Add
        Set Xarchexel = Xlibexel.Worksheets.Add
        Xarchexel.Name = Trim(Combo1.Text)
        Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo2.Text) & ".xls")
        Xarchtex = "C:\planillas\" & Trim(Combo2.Text) & ".xls"
        Regcli.MoveFirst
        Do While Not Regcli.EOF
           Data1.Recordset.AddNew
           Data1.Recordset("cl_fecing") = Format(Regcli("fecha"), "dd/mm/yyyy")
           Data1.Recordset("cl_codigo") = Val(Regcli("matric"))
           Data1.Recordset("cl_cedula") = Val(Regcli("cedula"))
           Data1.Recordset("cl_codced") = Val(Regcli("codced"))
           Data1.Recordset("cl_apellid") = Regcli("nombre")
           Data1.Recordset("cl_nombre") = Regcli("meta2desc")
           Data1.Recordset("cl_nrocobr") = Val(Regcli("base"))
           Data1.Recordset("cl_fnac") = Format(Regcli("fecctrl"), "dd/mm/yyyy")
           Data1.Recordset("cl_direcci") = Regcli("edadtex2")
           If IsNull(Regcli("peso")) = False Then
              Data1.Recordset("cl_fax") = Mid(Regcli("peso"), 1, 10)
           End If
           If IsNull(Regcli("lactan")) = False Then
              Data1.Recordset("cl_localid") = Mid(Regcli("lactan"), 1, 35)
           End If
           If IsNull(Regcli("vacuna")) = False Then
              Data1.Recordset("cl_nomvend") = Mid(Regcli("vacuna"), 1, 35)
           End If
           If IsNull(Regcli("ecocad")) = False Then
              Data1.Recordset("cl_zona") = Mid(Regcli("ecocad"), 1, 25)
           End If
           If IsNull(Regcli("fecprox")) = False Then
              Data1.Recordset("cl_nom_sup") = Mid(Regcli("fecprox"), 1, 25)
           End If
           If IsNull(Regcli("medico")) = False Then
              Data1.Recordset("cl_medflia") = Mid(Regcli("medico"), 1, 25)
           End If
           If IsNull(Regcli("obs")) = False Then
              Data1.Recordset("cl_entre") = Mid(Regcli("obs"), 1, 80)
           End If
           If IsNull(Regcli("lochc")) = False Then
              Data1.Recordset("cl_nomcobr") = Mid(Regcli("lochc"), 1, 25)
           End If
           If IsNull(Regcli("fecnac")) = False Then
              Data1.Recordset("cl_fultvta") = Format(Regcli("fecnac"), "dd/mm/yyyy")
           End If
           If IsNull(Regcli("cedmed")) = False Then
              Data1.Recordset("cl_telefon") = Mid(Regcli("cedmed"), 1, 20)
           End If
           If IsNull(Regcli("apemed")) = False Then
              Data1.Recordset("cl_nomconv") = Mid(Regcli("apemed"), 1, 30)
           End If
           If IsNull(Regcli("nommed")) = False Then
              Data1.Recordset("cl_socmnom") = Mid(Regcli("nommed"), 1, 20)
           End If
           
           Data1.Recordset("cl_codconv") = Regcli("cnvcod")
           Data1.Recordset.Update
           Regcli.MoveNext
        Loop
        Data1.Refresh
        If Combo2.Text = "TODAS" Then
        Else
           If Data1.Recordset.RecordCount > 0 Then
              ConectarBD
              ConbdSapp.Open
              Data1.Recordset.MoveFirst
              Do While Not Data1.Recordset.EOF
                 Sqlmeta = "Select * from convenio where cnv_codigo ='" & Data1.Recordset("cl_codconv") & "'"
                 With Regmeta
                     .CursorLocation = adUseClient
                     .CursorType = adOpenKeyset
                     .LockType = adLockOptimistic
                     .Open Sqlmeta, ConbdSapp, , , adCmdText
                 End With
                 If Regmeta.RecordCount > 0 Then
                    If IsNull(Regmeta("cnv_grupo")) = False Then
                       If Regmeta("cnv_grupo") = Combo2.Text Then
                       Else
                          Data1.Recordset.Delete
                       End If
                    Else
                       Data1.Recordset.Delete
                    End If
                 Else
                    Data1.Recordset.Delete
                 End If
                 Data1.Recordset.MoveNext
              Loop
              Data1.Refresh
              Regmeta.Close
              ConbdSapp.Close
           End If
        End If
        If Data1.Recordset.RecordCount > 0 Then
           Data1.Recordset.MoveFirst
           Xarchexel.Cells(Xlin, XCol) = "CENTRO DE COMPUTOS DE SAPP"
           Xlin = Xlin + 1
           XCol = XCol + 1
           Xarchexel.Range("A1", "C3").Font.Size = 16
           Xarchexel.Cells(Xlin, XCol) = "PLANILLA DE " & Combo2.Text & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
           Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
           XCol = 1
           Xlin = Xlin + 2
           Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
           Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
           Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
           Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
           Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
           Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
           Xarchexel.Range("A" & Trim(Str(Xlin)), "O" & Trim(Str(Xlin))).Interior.color = RGB(215, 120, 120)
           Xarchexel.Range("A" & Trim(Str(Xlin))).ColumnWidth = 10
           Xarchexel.Cells(Xlin, XCol) = "CI Afiliado" 'a
           XCol = XCol + 1
           Xarchexel.Range("B" & Trim(Str(Xlin))).ColumnWidth = 35
           Xarchexel.Cells(Xlin, XCol) = "Nro.HC" 'b
           XCol = XCol + 1
           Xarchexel.Range("C" & Trim(Str(Xlin))).ColumnWidth = 10
           Xarchexel.Cells(Xlin, XCol) = "Localización HC" 'c
           XCol = XCol + 1
           Xarchexel.Range("D" & Trim(Str(Xlin))).ColumnWidth = 10
           Xarchexel.Cells(Xlin, XCol) = "Fec.Nacimiento" 'd
           XCol = XCol + 1
           Xarchexel.Cells(Xlin, XCol) = "Ced.Médico" 'd
           XCol = XCol + 1
           Xarchexel.Range("F" & Trim(Str(Xlin))).ColumnWidth = 25
           Xarchexel.Cells(Xlin, XCol) = "Apellido del médico" 'f
           XCol = XCol + 1
           Xarchexel.Cells(Xlin, XCol) = "Nombre del médico" 'g
           XCol = XCol + 1
           Xarchexel.Range("H" & Trim(Str(Xlin))).ColumnWidth = 20
           Xarchexel.Cells(Xlin, XCol) = "Especialidad"
           XCol = XCol + 1
           Xarchexel.Range("I" & Trim(Str(Xlin))).ColumnWidth = 20
           Xarchexel.Cells(Xlin, XCol) = "Fec.Consulta 1"
           XCol = XCol + 1
           Xarchexel.Range("J" & Trim(Str(Xlin))).ColumnWidth = 20
           Xarchexel.Cells(Xlin, XCol) = "Fec.Consulta 2"
           Xcanttot = 1
           Xlin = Xlin + 1
           Do While Not Data1.Recordset.EOF
              XCol = 1
              Xarchexel.Cells(Xlin, XCol) = Trim(Str(Data1.Recordset("cl_cedula"))) & "-" & Trim(Str(Data1.Recordset("cl_codced")))
              XCol = XCol + 1
              Xarchexel.Cells(Xlin, XCol) = Trim(Str(Data1.Recordset("cl_cedula"))) & "-" & Trim(Str(Data1.Recordset("cl_codced")))
              XCol = XCol + 1
              If IsNull(Data1.Recordset("cl_nomcobr")) = False Then
                 Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nomcobr")
              Else
                 Xarchexel.Cells(Xlin, XCol) = ""
              End If
              XCol = XCol + 1
              Xarchexel.Cells(Xlin, XCol) = Format(Data1.Recordset("cl_fultvta"), "dd/mm/yyyy")
              XCol = XCol + 1
              If IsNull(Data1.Recordset("cl_telefon")) = False Then
                 Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_telefono")
              Else
                 Xarchexel.Cells(Xlin, XCol) = ""
              End If
              XCol = XCol + 1
              If IsNull(Data1.Recordset("cl_nomconv")) = False Then
                 Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nomconv")
              Else
                 Xarchexel.Cells(Xlin, XCol) = ""
              End If
              XCol = XCol + 1
              If IsNull(Data1.Recordset("cl_socmnom")) = False Then
                 Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_socmnom")
              Else
                 Xarchexel.Cells(Xlin, XCol) = ""
              End If
              XCol = XCol + 1
              Xarchexel.Cells(Xlin, XCol) = "MED.GRAL."
              XCol = XCol + 1
              Xarchexel.Cells(Xlin, XCol) = Format(Data1.Recordset("cl_fnac"), "dd/mm/yyyy")
              Xlin = Xlin + 1
              Data1.Recordset.MoveNext
              Xcanttot = Xcanttot + 1
           Loop
           Xlibexel.Save
           Xlibexel.Close
           Xobjexel.Quit
           Xlabrir.Workbooks.Open Xarchtex, , False
           Xlabrir.Visible = True
           Xlabrir.WindowState = xlMaximized
        Else
           Xlibexel.Close
           Xobjexel.Quit
        End If
     End If
     frm_infmetasnew.MousePointer = 0
     MsgBox "Terminado"
     Regcli.Close
     ConbdSappM.Close
   End If
Else
   If Combo1.Text = "CTROL.EMBARAZADAS" Then
      If Check1.value = 1 Then
         MsgBox "No existe control para ésta opción", vbInformation
      Else
        ConectarBDM
        ConbdSappM.Open
        If Combo1.Text = "TODAS" Then
           Sqlcli = "Select * from t_meta1 where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by cedula,fecha"
        Else
           Sqlcli = "Select * from t_meta1 where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and meta2desc ='" & Combo1.Text & "' order by cedula,fecha"
        End If
        With Regcli
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open Sqlcli, ConbdSappM, , , adCmdText
        End With
        If Regcli.RecordCount > 0 Then
           Set Xobjexel = New Excel.Application
           Set Xlibexel = Xobjexel.Workbooks.Add
           Set Xarchexel = Xlibexel.Worksheets.Add
           Xarchexel.Name = Trim(Combo1.Text)
           Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo2.Text) & ".xls")
           Xarchtex = "C:\planillas\" & Trim(Combo2.Text) & ".xls"
           Regcli.MoveFirst
           Do While Not Regcli.EOF
              Data1.Recordset.AddNew
              Data1.Recordset("cl_fecing") = Format(Regcli("fecha"), "dd/mm/yyyy")
              Data1.Recordset("cl_codigo") = Val(Regcli("matric"))
              Data1.Recordset("cl_cedula") = Val(Regcli("cedula"))
              Data1.Recordset("cl_codced") = Val(Regcli("codced"))
              Data1.Recordset("cl_apellid") = Regcli("nombre")
              Data1.Recordset("cl_nombre") = Regcli("meta2desc")
              Data1.Recordset("cl_nrocobr") = Val(Regcli("base"))
              Data1.Recordset("cl_fnac") = Format(Regcli("fecctrl"), "dd/mm/yyyy")
              Data1.Recordset("cl_direcci") = Regcli("edadtex2")
              If IsNull(Regcli("peso")) = False Then
                 Data1.Recordset("cl_fax") = Mid(Regcli("peso"), 1, 10)
              End If
              If IsNull(Regcli("lactan")) = False Then
                 Data1.Recordset("cl_localid") = Mid(Regcli("lactan"), 1, 35)
              End If
              If IsNull(Regcli("vacuna")) = False Then
                 Data1.Recordset("cl_nomvend") = Mid(Regcli("vacuna"), 1, 35)
              End If
              If IsNull(Regcli("ecocad")) = False Then
                 Data1.Recordset("cl_zona") = Mid(Regcli("ecocad"), 1, 25)
              End If
              If IsNull(Regcli("fecprox")) = False Then
                 Data1.Recordset("cl_nom_sup") = Mid(Regcli("fecprox"), 1, 25)
              End If
              If IsNull(Regcli("medico")) = False Then
                 Data1.Recordset("cl_medflia") = Mid(Regcli("medico"), 1, 25)
              End If
              If IsNull(Regcli("obs")) = False Then
                 Data1.Recordset("cl_entre") = Mid(Regcli("obs"), 1, 80)
              End If
              If IsNull(Regcli("lochc")) = False Then
                 Data1.Recordset("cl_nomcobr") = Mid(Regcli("lochc"), 1, 25)
              End If
              If IsNull(Regcli("fecnac")) = False Then
                 Data1.Recordset("cl_fultvta") = Format(Regcli("fecnac"), "dd/mm/yyyy")
              End If
              If IsNull(Regcli("cedmed")) = False Then
                 Data1.Recordset("cl_telefon") = Mid(Regcli("cedmed"), 1, 20)
              End If
              If IsNull(Regcli("apemed")) = False Then
                 Data1.Recordset("cl_nomconv") = Mid(Regcli("apemed"), 1, 30)
              End If
              If IsNull(Regcli("nommed")) = False Then
                 Data1.Recordset("cl_socmnom") = Mid(Regcli("nommed"), 1, 20)
              End If
              Data1.Recordset("cl_codconv") = Regcli("cnvcod")
              Data1.Recordset("cl_forpago") = Regcli("nroconse")
              Data1.Recordset("cl_tipocli") = Regcli("cbopap")
              Data1.Recordset("cl_fultmov") = Regcli("fecvdrl")
              Data1.Recordset("cl_atrasop") = Regcli("cbomamo")
              Data1.Recordset("cl_atrasoa") = Regcli("prot")
              If IsNull(Regcli("obsemb")) = False Then
                 Data1.Recordset("cl_entre") = Mid(Regcli("obsemb"), 1, 80)
              End If
              If IsNull(Regcli("fecanti")) = False Then
                 Data1.Recordset("cl_faviso1") = Format(Regcli("fecanti"), "dd/mm/yyyy")
              End If
              If IsNull(Regcli("semgono")) = False Then
                 Data1.Recordset("cl_descpag") = Mid(Regcli("semgono"), 1, 25)
              End If
              If IsNull(Regcli("odont2")) = False Then
                 Data1.Recordset("cl_cantpag") = Regcli("odont2")
              End If
              
              Data1.Recordset.Update
              Regcli.MoveNext
           Loop
           Data1.Refresh
           If Combo2.Text = "TODAS" Then
              ConectarBD
              ConbdSapp.Open
              If Data1.Recordset.RecordCount > 0 Then
                 Data1.Recordset.MoveFirst
                 Do While Not Data1.Recordset.EOF
                    Sqlmeta = "Select * from convenio where cnv_codigo ='" & Data1.Recordset("cl_codconv") & "'"
                    With Regmeta
                        .CursorLocation = adUseClient
                        .CursorType = adOpenKeyset
                        .LockType = adLockOptimistic
                        .Open Sqlmeta, ConbdSapp, , , adCmdText
                    End With
                    If Regmeta.RecordCount > 0 Then
                       If IsNull(Regmeta("cnv_grupo")) = False Then
                          If Regmeta("cnv_grupo") = Combo2.Text Then
                             Data1.Recordset.Edit
                             Data1.Recordset("cl_nrosocm") = Regmeta("cnv_grupo")
                             Data1.Recordset.Update
                          End If
                       End If
                    End If
                    Data1.Recordset.MoveNext
                 Loop
              End If
           Else
              If Data1.Recordset.RecordCount > 0 Then
                 ConectarBD
                 ConbdSapp.Open
                 Data1.Recordset.MoveFirst
                 Do While Not Data1.Recordset.EOF
                    Sqlmeta = "Select * from convenio where cnv_codigo ='" & Data1.Recordset("cl_codconv") & "'"
                    With Regmeta
                        .CursorLocation = adUseClient
                        .CursorType = adOpenKeyset
                        .LockType = adLockOptimistic
                        .Open Sqlmeta, ConbdSapp, , , adCmdText
                    End With
                    If Regmeta.RecordCount > 0 Then
                       If IsNull(Regmeta("cnv_grupo")) = False Then
                          If Regmeta("cnv_grupo") = Combo2.Text Then
                             Data1.Recordset.Edit
                             Data1.Recordset("cl_nrosocm") = Regmeta("cnv_grupo")
                             Data1.Recordset.Update
                          Else
                             Data1.Recordset.Delete
                          End If
                       Else
                          Data1.Recordset.Delete
                       End If
                    Else
                       Data1.Recordset.Delete
                    End If
                    Data1.Recordset.MoveNext
                 Loop
                 Data1.Refresh
                 Regmeta.Close
                 ConbdSapp.Close
              End If
           End If
           If Data1.Recordset.RecordCount > 0 Then
              Data1.Recordset.MoveFirst
              Xarchexel.Cells(Xlin, XCol) = "CENTRO DE COMPUTOS DE SAPP"
              Xlin = Xlin + 1
              XCol = XCol + 1
              Xarchexel.Range("A1", "C3").Font.Size = 16
              Xarchexel.Cells(Xlin, XCol) = "PLANILLA DE " & Combo2.Text & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
              Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
              XCol = 1
              Xlin = Xlin + 2
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
              Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
              Xarchexel.Range("A" & Trim(Str(Xlin)), "O" & Trim(Str(Xlin))).Interior.color = RGB(215, 120, 120)
              Xarchexel.Range("A" & Trim(Str(Xlin))).ColumnWidth = 35
              Xarchexel.Cells(Xlin, XCol) = "Apellidos y Nombres" 'a
              XCol = XCol + 1
              Xarchexel.Range("B" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "Cédula" 'b
              XCol = XCol + 1
              Xarchexel.Range("C" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "Edad" 'c
              XCol = XCol + 1
              Xarchexel.Range("D" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "Mutualista" 'd
              XCol = XCol + 1
              Xarchexel.Cells(Xlin, XCol) = "Nro.Cons." 'd
              XCol = XCol + 1
              Xarchexel.Range("F" & Trim(Str(Xlin))).ColumnWidth = 25
              Xarchexel.Cells(Xlin, XCol) = "Semanas de Amenorrea" 'f
              XCol = XCol + 1
              Xarchexel.Cells(Xlin, XCol) = "Deriv.Odont." 'g
              XCol = XCol + 1
              Xarchexel.Range("H" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "Fec.HIV"
              XCol = XCol + 1
              Xarchexel.Range("I" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "Fec.VDRL"
              XCol = XCol + 1
              Xarchexel.Range("J" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "Consent.Inform."
              XCol = XCol + 1
              Xarchexel.Range("K" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "PAP"
              XCol = XCol + 1
              Xarchexel.Range("L" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "Mamografía"
              XCol = XCol + 1
              Xarchexel.Range("M" & Trim(Str(Xlin))).ColumnWidth = 10
              Xarchexel.Cells(Xlin, XCol) = "Anticoncep."
              Xcanttot = 1
              Xlin = Xlin + 1
              Do While Not Data1.Recordset.EOF
                 XCol = 1
                 Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_apellid")
                 XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = Trim(Str(Data1.Recordset("cl_cedula"))) & "-" & Trim(Str(Data1.Recordset("cl_codced")))
                 XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_direcci")
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_nrosocm")) = False Then
                    Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nrosocm")
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_forpago")) = False Then
                    Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_forpago")
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_descpag")) = False Then
                    Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_descpag")
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_cantpag")) = False Then
                    If Data1.Recordset("cl_cantpag") = 0 Then
                       Xarchexel.Cells(Xlin, XCol) = "SI"
                    Else
                       Xarchexel.Cells(Xlin, XCol) = "NO"
                    End If
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_fultmov")) = False Then
                    Xarchexel.Cells(Xlin, XCol) = "HIV:" & Format(Data1.Recordset("cl_fultmov"), "dd/mm/yyyy")
                    XCol = XCol + 1
                    Xarchexel.Cells(Xlin, XCol) = "VDRL:" & Format(Data1.Recordset("cl_fultmov"), "dd/mm/yyyy")
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                    XCol = XCol + 1
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_atrasoa")) = False Then
                    If Data1.Recordset("cl_atrasoa") = 0 Then
                       Xarchexel.Cells(Xlin, XCol) = "SI"
                    Else
                       Xarchexel.Cells(Xlin, XCol) = "NO"
                    End If
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_tipocli")) = False Then
                    If Data1.Recordset("cl_tipocli") = 0 Then
                       Xarchexel.Cells(Xlin, XCol) = "SI"
                    Else
                       Xarchexel.Cells(Xlin, XCol) = "NO"
                    End If
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_atrasop")) = False Then
                    If Data1.Recordset("cl_atrasop") = 0 Then
                       Xarchexel.Cells(Xlin, XCol) = "SI"
                    Else
                       Xarchexel.Cells(Xlin, XCol) = "NO"
                    End If
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_entre")) = False Then
                    Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_entre")
                 Else
                    Xarchexel.Cells(Xlin, XCol) = ""
                 End If
                 Xlin = Xlin + 1
                 Data1.Recordset.MoveNext
                 Xcanttot = Xcanttot + 1
              Loop
              Xlibexel.Save
              Xlibexel.Close
              Xobjexel.Quit
              Xlabrir.Workbooks.Open Xarchtex, , False
              Xlabrir.Visible = True
              Xlabrir.WindowState = xlMaximized
           Else
              Xlibexel.Close
              Xobjexel.Quit
           End If
        End If
        frm_infmetasnew.MousePointer = 0
        MsgBox "Terminado"
        Regcli.Close
        ConbdSappM.Close
      End If
   Else
      If Combo1.Text = "CTROL.12 A 19AÑOS" Then
         If Check1.value = 1 Then
            If Combo2.ListIndex >= 1 Then
               ConectarBD
               ConbdSapp.Open
               Sqlmeta = "Select * from clientes where cl_fnac is not null and estado =" & 1
               With Regmeta
                    .CursorLocation = adUseClient
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    .Open Sqlmeta, ConbdSapp, , , adCmdText
               End With
               If Regmeta.RecordCount > 0 Then
                  Regmeta.MoveFirst
                  Do While Not Regmeta.EOF
                     Sqlconsmeta = "Select * from convenio where cnv_codigo ='" & Regmeta("cl_codconv") & "'"
                     With Regbusca
                         .CursorLocation = adUseClient
                         .CursorType = adOpenKeyset
                         .LockType = adLockOptimistic
                         .Open Sqlconsmeta, ConbdSapp, , , adCmdText
                     End With
                     If Regbusca.RecordCount > 0 Then
                        If IsNull(Regbusca("cnv_grupo")) = False Then
                           If Regbusca("cnv_grupo") = Combo2.Text Then
                              Data1.Recordset.AddNew
                              Data1.Recordset("cl_codigo") = Regmeta("cl_codigo")
                              Data1.Recordset("cl_apellid") = Regmeta("cl_apellid")
                              Data1.Recordset("cl_cedula") = Regmeta("cl_cedula")
                              Data1.Recordset("cl_codced") = Regmeta("cl_codced")
                              Data1.Recordset("cl_codconv") = Regmeta("cl_codconv")
                              Data1.Recordset("cl_nomconv") = Regmeta("cl_nomconv")
                              Data1.Recordset("cl_fnac") = Regmeta("cl_fnac")
                              Data1.Recordset("cl_fecing") = Regmeta("cl_fecing")
                              Data1.Recordset("cl_telefon") = Regmeta("cl_telefon")
                              Data1.Recordset("cl_dpto") = Regmeta("cl_dpto")
                              Data1.Recordset("cl_zona") = Regmeta("cl_zona")
                              Data1.Recordset.Update
                           End If
                        End If
                     End If
                     Regbusca.Close
                     Regmeta.MoveNext
                  Loop
                  Regmeta.Close
                  Data1.Refresh
                  If Data1.Recordset.RecordCount > 0 Then
                     Data1.Recordset.MoveFirst
                     Do While Not Data1.Recordset.EOF
                        CalculaEdad (Data1.Recordset("cl_fnac"))
                        If Val(laba.Caption) >= 12 And Val(laba.Caption) <= 19 Then
                           Sqlmeta = "Select * from linmmdd where cod_prod in (190011,190012) and cod_cli =" & Data1.Recordset("cl_codigo")
                           With Regmeta
                               .CursorLocation = adUseClient
                               .CursorType = adOpenKeyset
                               .LockType = adLockOptimistic
                               .Open Sqlmeta, ConbdSapp, , , adCmdText
                           End With
                           If Regmeta.RecordCount > 0 Then
                              Data1.Recordset.Delete
                           Else
                              Data1.Recordset.Edit
                              Data1.Recordset("cl_dircobr") = "META 2 -CONTROL 12 a 19 AÑOS/HOJA SIA"
                              Data1.Recordset.Update
                           End If
                           Regmeta.Close
                        Else
                           Data1.Recordset.Delete
                        End If
                     Loop
                     frm_infmetasnew.MousePointer = 0
                     MsgBox "Terminado"
                     Data1.RecordSource = "Select * from infcli order by cl_codigo"
                     Data1.Refresh
                     cr1.ReportFileName = App.Path & "\infmetasnew.rpt"
                     cr1.ReportTitle = "Informe de socios ACTIVOS con METAS PENDIENTES"
                     cr1.Action = 1
                  End If
               End If
               ConbdSapp.Close
            Else
               frm_infmetasnew.MousePointer = 0
               MsgBox "Debe seleccionar una mutualista", vbInformation
            End If
         Else
            ConectarBDM
            ConbdSappM.Open
            If Combo1.Text = "TODAS" Then
               Sqlcli = "Select * from t_meta1 where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by cedula,fecha"
            Else
               Sqlcli = "Select * from t_meta1 where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and meta2desc ='" & Combo1.Text & "' order by cedula,fecha"
            End If
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSappM, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               Set Xobjexel = New Excel.Application
               Set Xlibexel = Xobjexel.Workbooks.Add
               Set Xarchexel = Xlibexel.Worksheets.Add
               Xarchexel.Name = Trim(Combo1.Text)
               Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo2.Text) & ".xls")
               Xarchtex = "C:\planillas\" & Trim(Combo2.Text) & ".xls"
               Regcli.MoveFirst
               Do While Not Regcli.EOF
                  Data1.Recordset.AddNew
                  Data1.Recordset("cl_fecing") = Format(Regcli("fecha"), "dd/mm/yyyy")
                  Data1.Recordset("cl_codigo") = Val(Regcli("matric"))
                  Data1.Recordset("cl_cedula") = Val(Regcli("cedula"))
                  Data1.Recordset("cl_codced") = Val(Regcli("codced"))
                  Data1.Recordset("cl_apellid") = Regcli("nombre")
                  Data1.Recordset("cl_nombre") = Regcli("meta2desc")
                  Data1.Recordset("cl_nrocobr") = Val(Regcli("base"))
                  Data1.Recordset("cl_fnac") = Format(Regcli("fecctrl"), "dd/mm/yyyy")
                  Data1.Recordset("cl_direcci") = Regcli("edadtex2")
                  If IsNull(Regcli("fecprox")) = False Then
                     Data1.Recordset("cl_nom_sup") = Mid(Regcli("fecprox"), 1, 25)
                  End If
                  If IsNull(Regcli("medico")) = False Then
                     Data1.Recordset("cl_medflia") = Mid(Regcli("medico"), 1, 25)
                  End If
                  If IsNull(Regcli("sia")) = False Then
                     Data1.Recordset("cl_nrovende") = Regcli("sia")
                  End If
                  If IsNull(Regcli("fecnac")) = False Then
                     Data1.Recordset("cl_fultmov") = Format(Regcli("fecnac"), "dd/mm/yyyy")
                  End If
                  Data1.Recordset.Update
                  Regcli.MoveNext
               Loop
               Data1.Refresh
               If Combo2.Text = "TODAS" Then
               Else
                  If Data1.Recordset.RecordCount > 0 Then
                     ConectarBD
                     ConbdSapp.Open
                     Data1.Recordset.MoveFirst
                     Do While Not Data1.Recordset.EOF
                        Sqlmeta = "Select * from convenio where cnv_codigo ='" & Data1.Recordset("cl_codconv") & "'"
                        With Regmeta
                            .CursorLocation = adUseClient
                            .CursorType = adOpenKeyset
                            .LockType = adLockOptimistic
                            .Open Sqlmeta, ConbdSapp, , , adCmdText
                        End With
                        If Regmeta.RecordCount > 0 Then
                           If IsNull(Regmeta("cnv_grupo")) = False Then
                              If Regmeta("cnv_grupo") = Combo2.Text Then
                              Else
                                 Data1.Recordset.Delete
                              End If
                           Else
                              Data1.Recordset.Delete
                           End If
                        Else
                           Data1.Recordset.Delete
                        End If
                        Data1.Recordset.MoveNext
                     Loop
                     Data1.Refresh
                     Regmeta.Close
                     ConbdSapp.Close
                  End If
               End If
               If Data1.Recordset.RecordCount > 0 Then
                  Data1.Recordset.MoveFirst
                  Xarchexel.Cells(Xlin, XCol) = "CENTRO DE COMPUTOS DE SAPP"
                  Xlin = Xlin + 1
                  XCol = XCol + 1
                  Xarchexel.Range("A1", "C3").Font.Size = 16
                  Xarchexel.Cells(Xlin, XCol) = "PLANILLA DE " & Combo2.Text & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                  Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
                  XCol = 1
                  Xlin = Xlin + 2
                  Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                  Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                  Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                  Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                  Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                  Xarchexel.Range("A4", "O" & Trim(Str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                  Xarchexel.Range("A" & Trim(Str(Xlin)), "O" & Trim(Str(Xlin))).Interior.color = RGB(215, 120, 120)
                  Xarchexel.Range("A" & Trim(Str(Xlin))).ColumnWidth = 15
                  Xarchexel.Cells(Xlin, XCol) = "CEDULA" 'a
                  XCol = XCol + 1
                  Xarchexel.Range("B" & Trim(Str(Xlin))).ColumnWidth = 35
                  Xarchexel.Cells(Xlin, XCol) = "NOMBRE" 'b
                  XCol = XCol + 1
                  Xarchexel.Range("C" & Trim(Str(Xlin))).ColumnWidth = 10
                  Xarchexel.Cells(Xlin, XCol) = "FEC.NAC." 'c
                  XCol = XCol + 1
                  Xarchexel.Range("D" & Trim(Str(Xlin))).ColumnWidth = 25
                  Xarchexel.Cells(Xlin, XCol) = "MEDICO REF." 'd
                  XCol = XCol + 1
                  Xarchexel.Cells(Xlin, XCol) = "HOJA SIA" 'd
                  XCol = XCol + 1
                  Xarchexel.Range("F" & Trim(Str(Xlin))).ColumnWidth = 25
                  Xarchexel.Cells(Xlin, XCol) = "PROX.CONTROL" 'f
                  XCol = XCol + 1
                  Xlin = Xlin + 1
                  Do While Not Data1.Recordset.EOF
                     XCol = 1
                     Xarchexel.Cells(Xlin, XCol) = Trim(Str(Data1.Recordset("cl_cedula"))) & "-" & Trim(Str(Data1.Recordset("cl_codced")))
                     XCol = XCol + 1
                     Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_apellid")
                     XCol = XCol + 1
                     If IsNull(Data1.Recordset("cl_fultmov")) = False Then
                        Xarchexel.Cells(Xlin, XCol) = Format(Data1.Recordset("cl_fultmov"), "dd/mm/yyyy")
                     Else
                        Xarchexel.Cells(Xlin, XCol) = ""
                     End If
                     XCol = XCol + 1
                     If IsNull(Data1.Recordset("cl_medflia")) = False Then
                        Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_medflia")
                     Else
                        Xarchexel.Cells(Xlin, XCol) = ""
                     End If
                     XCol = XCol + 1
                     If IsNull(Data1.Recordset("cl_nrovende")) = False Then
                        If Data1.Recordset("cl_nrovende") = 0 Then
                           Xarchexel.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(Data1.Recordset("cl_nom_sup")) = False Then
                        Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("cl_nom_sup")
                     Else
                        Xarchexel.Cells(Xlin, XCol) = ""
                     End If
                     Xlin = Xlin + 1
                     Data1.Recordset.MoveNext
                  Loop
                  frm_infmetasnew.MousePointer = 0
                  MsgBox "Terminado"
                  Xlibexel.Save
                  Xlibexel.Close
                  Xobjexel.Quit
                  Xlabrir.Workbooks.Open Xarchtex, , False
                  Xlabrir.Visible = True
                  Xlabrir.WindowState = xlMaximized
               Else
                  frm_infmetasnew.MousePointer = 0
                  MsgBox "No hay registros"
               End If
               Xlibexel.Close
               Xobjexel.Quit
            Else
               frm_infmetasnew.MousePointer = 0
               MsgBox "No hay registros"
            End If
            ConbdSappM.Close
         End If
      End If
   End If
End If

End Sub

Private Sub Form_Load()

ConectarBDM
ConbdSappM.Open

Sqlcli = "Select * from t_descmeta2"
With Regcli
     .CursorLocation = adUseClient
     .CursorType = adOpenKeyset
     .LockType = adLockOptimistic
     .Open Sqlcli, ConbdSappM, , , adCmdText
End With

If Regcli.RecordCount > 0 Then
   Do While Not Regcli.EOF
      Combo1.AddItem Regcli("descrip")
      Regcli.MoveNext
   Loop
   Combo1.AddItem "TODAS"
End If
Regcli.Close
ConbdSappM.Close

Data1.DatabaseName = App.Path & "\informes.mdb"
Data1.RecordSource = "infcli"
Data1.Refresh


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
Public Function ConectarBD()
ConbdSapp.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PORT=3306;DATABASE=mmsyssapp;USER=root;PASSWORD=sapp1987;OPTION=3;"

End Function

Public Function ConectarBDM()
ConbdSappM.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=sappbd;USER=root;PASSWORD=sapp1987;OPTION=3;"

End Function

Private Sub MaskEdBox1_Change()

End Sub


Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub CalculaEdad(ByVal FNaci As Date)
Dim FAct As String
Dim Anios As String
Dim Meses As String
Dim Dias As String
Dim newday As String
Dim newmonth As String
Dim newyear As String

FAct = Format(Now, "dd/MM/yyyy")
FNaci = Format(FNaci, "dd/MM/yyyy")

'Calcula los años
Anios = DateDiff("yyyy", CDate(Format(FNaci, "dd/MM/yyyy")), CDate(FAct))
'Si el mes actual es menor que el mes de la fecha de nacimiento entonces
If Month(CDate(FAct)) < Month(CDate(FNaci)) Then
 'Restele uno a los años
 Anios = Anios - 1
 newmonth = Month(CDate(FAct)) + 12
 Else
 'Deje el mes actual tal y como estan
 newmonth = Month(CDate(FAct))
 End If

 'Si el mes actual es igual al mes de la fecha de nacimiento entonces
If Month(CDate(FAct)) = Month(CDate(FNaci)) Then
 'Si el día de la fecha actual es menor al día de la fecha de nacimiento
 If Day(CDate(FAct)) < Day(CDate(FNaci)) Then
 'Restele uno a los años
 Anios = Anios - 1
 End If
End If

If Day(CDate(FAct)) < Day(CDate(FNaci)) Then

   If Month(FNaci) = 1 Or Month(FNaci) = 3 Or Month(FNaci) = 5 Or _
      Month(FNaci) = 7 Or Month(FNaci) = 8 Or Month(FNaci) = 10 Or _
      Month(FNaci) = 12 Then
      newday = Day(CDate(FAct)) + 31
   Else
      If Month(FNaci) = 2 Then
         newday = Day(CDate(FAct)) + 28
      Else
         newday = Day(CDate(FAct)) + 30
      End If
   End If
   newmonth = newmonth - 1
Else
   newday = Day(CDate(FAct))
End If

If Month(CDate(FNaci)) = Month(Date) Then
   
   Meses = 0
Else
   Meses = newmonth - Month(CDate(FNaci))
End If

If Meses < 0 And Anios = 0 Then
   Meses = Meses + 12
End If

Dias = newday - Day(CDate(FNaci))

If FNaci <= FAct Then

'Me.TextBox3.Text = Anios & " Años, " & Meses & " Meses, " & Dias & " Dias."
'''   labedad.Caption = Anios
   If Month(Date) = Month(FNaci) Then
      If Day(Date) > Day(FNaci) Then
         Meses = Meses
      Else
         If Day(Date) = Day(FNaci) Then
            Meses = 0
         Else
            Meses = 11
         End If
      End If
   End If
   laba.Caption = Anios
   labm.Caption = Meses
   labd.Caption = Dias
Else
'   MsgBox "Fecha Inválida"
   laba.Caption = 0
   labm.Caption = 0
   labd.Caption = 0

End If

End Sub


