VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infhcemet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes para Metas desde HCE"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6390
   Icon            =   "frm_infhcemet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6390
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   2400
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      Picture         =   "frm_infhcemet.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Procesar informe"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos para el informe"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.Data adodc1 
         Caption         =   "adodc1"
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
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   2640
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infhcemet.frx":0B14
         Left            =   2040
         List            =   "frm_infhcemet.frx":0B1E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3960
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
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "Opción de informe:"
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
         Top             =   1080
         Width           =   1815
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   1560
      Picture         =   "frm_infhcemet.frx":0B46
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   3495
   End
End
Attribute VB_Name = "frm_infhcemet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xelvalnuevo As Long
Dim Xelvalnuevo2 As Long

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
Data1.RecordSource = "infcli"
Data1.Refresh
frm_infhcemet.MousePointer = 11
Command1.Enabled = False

If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
   If Combo1.ListIndex = 0 Then
      adodc1.RecordSource = "Select * from hc_mcyotro where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
   Else
      If Combo1.ListIndex = 1 Then
         adodc1.RecordSource = "Select * from prestamo where nomc ='" & "MEDICO DE REFERENCIA" & "' and fecing >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecing <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecing DESC"
      Else
         adodc1.RecordSource = "Select * from hc_mcyotro where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
      End If
   End If
   adodc1.Refresh
   If adodc1.Recordset.RecordCount > 0 Then
      If Combo1.ListIndex = 1 Then
'            labmr.Caption = adohcante.Recordset("nomc") & " " & adohcante.Recordset("desccar") & " FECHA:" & adohcante.Recordset("fecing")
'         Adodc1.Recordset.MoveFirst
         Do While Not adodc1.Recordset.EOF
            Adodc2.RecordSource = "Select * from cabezal_hc where cb_mat =" & Val(adodc1.Recordset("nom1"))
            Adodc2.Refresh
            If Adodc2.Recordset.RecordCount > 0 Then
               Data1.Recordset.AddNew
               Data1.Recordset("cl_codconv") = Adodc2.Recordset("cb_codconv")
               Data1.Recordset("cl_direcci") = "MED.REF. " & adodc1.Recordset("desccar")
               Data1.Recordset("cl_codigo") = Adodc2.Recordset("cb_mat")
               If IsNull(Adodc2.Recordset("cb_fnac")) = False Then
                  Data1.Recordset("cl_fnac") = Adodc2.Recordset("cb_fnac")
               End If
               If IsNull(Adodc2.Recordset("cb_ape1")) = False Then
                  If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                     Data1.Recordset("cl_apellid") = Adodc2.Recordset("cb_ape1") & " " & Adodc2.Recordset("cb_nom1")
                  Else
                     Data1.Recordset("cl_apellid") = Adodc2.Recordset("cb_ape1")
                  End If
               Else
                  If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                     Data1.Recordset("cl_apellid") = "NN" & Adodc2.Recordset("cb_nom1")
                  Else
                     Data1.Recordset("cl_apellid") = "NN"
                  End If
               End If
               
               Adodc2.RecordSource = "Select * from cabezal_hcdig where mat =" & Val(adodc1.Recordset("nom1")) & " and id =" & adodc1.Recordset("retleg")
               Adodc2.Refresh
               If Adodc2.Recordset.RecordCount > 0 Then
                  Data1.Recordset("cl_fecing") = Adodc2.Recordset("fecha")
                  Data1.Recordset("cl_nombre") = Mid(Adodc2.Recordset("hc_nommed"), 1, 30)
                  Data1.Recordset("cl_nrocobr") = Adodc2.Recordset("hc_base")
                  Data1.Recordset("cl_dpto") = Adodc2.Recordset("cedtext")
               End If
               Data1.Recordset.Update
            End If
            adodc1.Recordset.MoveNext
         Loop
         frm_infhcemet.MousePointer = 0
         Command1.Enabled = True
         MsgBox "Proceso terminado"
         cr1.ReportFileName = App.Path & "\infhcemet1.rpt"
         cr1.ReportTitle = "Informe de Med.Referencia Fecha: " & md.Text & " hasta:" & mh.Text
         cr1.Action = 1
      Else
         adodc1.Recordset.MoveFirst
         Do While Not adodc1.Recordset.EOF
            If IsNull(adodc1.Recordset("hc_htasi")) = False Then
               If adodc1.Recordset("hc_htasi") = 1 Then
                  If IsNull(adodc1.Recordset("hc_violencia")) = False Then
                     If adodc1.Recordset("hc_violencia") = 1 Then
                        Adodc2.RecordSource = "Select * from cabezal_hc where cb_mat =" & adodc1.Recordset("hc_mat")
                        Adodc2.Refresh
                        If Adodc2.Recordset.RecordCount > 0 Then
                           Data1.Recordset.AddNew
                           Data1.Recordset("cl_codconv") = Adodc2.Recordset("cb_codconv")
                           Data1.Recordset("cl_direcci") = "VD y HTA Positivo"
                           Data1.Recordset("cl_codigo") = Adodc2.Recordset("cb_mat")
                           If IsNull(Adodc2.Recordset("cb_fnac")) = False Then
                              Data1.Recordset("cl_fnac") = Adodc2.Recordset("cb_fnac")
                           End If
                           If IsNull(Adodc2.Recordset("cb_ape1")) = False Then
                              If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                                 Data1.Recordset("cl_apellid") = Mid(Adodc2.Recordset("cb_ape1"), 1, 15) & " " & Mid(Adodc2.Recordset("cb_nom1"), 1, 14)
                              Else
                                 Data1.Recordset("cl_apellid") = Adodc2.Recordset("cb_ape1")
                              End If
                           Else
                              If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                                 Data1.Recordset("cl_apellid") = "NN" & Adodc2.Recordset("cb_nom1")
                              Else
                                 Data1.Recordset("cl_apellid") = "NN"
                              End If
                           End If
                           Adodc2.RecordSource = "Select * from cabezal_hcdig where mat =" & adodc1.Recordset("hc_mat") & " and id =" & adodc1.Recordset("hc_nro")
                           Adodc2.Refresh
                           If Adodc2.Recordset.RecordCount > 0 Then
                              Data1.Recordset("cl_fecing") = Adodc2.Recordset("fecha")
                              Data1.Recordset("cl_nombre") = Mid(Adodc2.Recordset("hc_nommed"), 1, 30)
                              Data1.Recordset("cl_nrocobr") = Adodc2.Recordset("hc_base")
                              Data1.Recordset("cl_dpto") = Adodc2.Recordset("cedtext")
                           End If
                           Data1.Recordset.Update
                        End If
                     Else
                        Adodc2.RecordSource = "Select * from cabezal_hc where cb_mat =" & adodc1.Recordset("hc_mat")
                        Adodc2.Refresh
                        If Adodc2.Recordset.RecordCount > 0 Then
                           Data1.Recordset.AddNew
                           Data1.Recordset("cl_codconv") = Adodc2.Recordset("cb_codconv")
                           Data1.Recordset("cl_direcci") = "HTA Positivo"
                           Data1.Recordset("cl_codigo") = Adodc2.Recordset("cb_mat")
                           If IsNull(Adodc2.Recordset("cb_fnac")) = False Then
                              Data1.Recordset("cl_fnac") = Adodc2.Recordset("cb_fnac")
                           End If
                           If IsNull(Adodc2.Recordset("cb_ape1")) = False Then
                              If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                                 Data1.Recordset("cl_apellid") = Mid(Adodc2.Recordset("cb_ape1"), 1, 15) & " " & Mid(Adodc2.Recordset("cb_nom1"), 1, 14)
                              Else
                                 Data1.Recordset("cl_apellid") = Adodc2.Recordset("cb_ape1")
                              End If
                           Else
                              If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                                 Data1.Recordset("cl_apellid") = "NN" & Adodc2.Recordset("cb_nom1")
                              Else
                                 Data1.Recordset("cl_apellid") = "NN"
                              End If
                           End If
                           Adodc2.RecordSource = "Select * from cabezal_hcdig where mat =" & adodc1.Recordset("hc_mat") & " and id =" & adodc1.Recordset("hc_nro")
                           Adodc2.Refresh
                           If Adodc2.Recordset.RecordCount > 0 Then
                              Data1.Recordset("cl_fecing") = Adodc2.Recordset("fecha")
                              Data1.Recordset("cl_nombre") = Mid(Adodc2.Recordset("hc_nommed"), 1, 30)
                              Data1.Recordset("cl_nrocobr") = Adodc2.Recordset("hc_base")
                              Data1.Recordset("cl_dpto") = Adodc2.Recordset("cedtext")
                           End If
                           Data1.Recordset.Update
                        End If
                     
                     End If
                  Else
                      Adodc2.RecordSource = "Select * from cabezal_hc where cb_mat =" & adodc1.Recordset("hc_mat")
                      Adodc2.Refresh
                      If Adodc2.Recordset.RecordCount > 0 Then
                         Data1.Recordset.AddNew
                         Data1.Recordset("cl_codconv") = Adodc2.Recordset("cb_codconv")
                         Data1.Recordset("cl_direcci") = "HTA Positivo"
                         Data1.Recordset("cl_codigo") = Adodc2.Recordset("cb_mat")
                         If IsNull(Adodc2.Recordset("cb_fnac")) = False Then
                            Data1.Recordset("cl_fnac") = Adodc2.Recordset("cb_fnac")
                         End If
                         If IsNull(Adodc2.Recordset("cb_ape1")) = False Then
                            If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                               Data1.Recordset("cl_apellid") = Mid(Adodc2.Recordset("cb_ape1"), 1, 15) & " " & Mid(Adodc2.Recordset("cb_nom1"), 1, 14)
                            Else
                               Data1.Recordset("cl_apellid") = Adodc2.Recordset("cb_ape1")
                            End If
                         Else
                            If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                               Data1.Recordset("cl_apellid") = "NN" & Adodc2.Recordset("cb_nom1")
                            Else
                               Data1.Recordset("cl_apellid") = "NN"
                            End If
                         End If
                         Adodc2.RecordSource = "Select * from cabezal_hcdig where mat =" & adodc1.Recordset("hc_mat") & " and id =" & adodc1.Recordset("hc_nro")
                         Adodc2.Refresh
                         If Adodc2.Recordset.RecordCount > 0 Then
                            Data1.Recordset("cl_fecing") = Adodc2.Recordset("fecha")
                            Data1.Recordset("cl_nombre") = Mid(Adodc2.Recordset("hc_nommed"), 1, 30)
                            Data1.Recordset("cl_nrocobr") = Adodc2.Recordset("hc_base")
                            Data1.Recordset("cl_dpto") = Adodc2.Recordset("cedtext")
                         End If
                         Data1.Recordset.Update
                      End If
                  End If
               Else
                  If IsNull(adodc1.Recordset("hc_violencia")) = False Then
                     If adodc1.Recordset("hc_violencia") = 1 Then
                        Adodc2.RecordSource = "Select * from cabezal_hc where cb_mat =" & adodc1.Recordset("hc_mat")
                        Adodc2.Refresh
                        If Adodc2.Recordset.RecordCount > 0 Then
                           Data1.Recordset.AddNew
                           Data1.Recordset("cl_codconv") = Adodc2.Recordset("cb_codconv")
                           Data1.Recordset("cl_direcci") = "VD Positivo"
                           Data1.Recordset("cl_codigo") = Adodc2.Recordset("cb_mat")
                           If IsNull(Adodc2.Recordset("cb_fnac")) = False Then
                              Data1.Recordset("cl_fnac") = Adodc2.Recordset("cb_fnac")
                           End If
                           If IsNull(Adodc2.Recordset("cb_ape1")) = False Then
                              If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                                 Data1.Recordset("cl_apellid") = Mid(Adodc2.Recordset("cb_ape1"), 1, 15) & " " & Mid(Adodc2.Recordset("cb_nom1"), 1, 14)
                              Else
                                 Data1.Recordset("cl_apellid") = Adodc2.Recordset("cb_ape1")
                              End If
                           Else
                              If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                                 Data1.Recordset("cl_apellid") = "NN" & Adodc2.Recordset("cb_nom1")
                              Else
                                 Data1.Recordset("cl_apellid") = "NN"
                              End If
                           End If
                           Adodc2.RecordSource = "Select * from cabezal_hcdig where mat =" & adodc1.Recordset("hc_mat") & " and id =" & adodc1.Recordset("hc_nro")
                           Adodc2.Refresh
                           If Adodc2.Recordset.RecordCount > 0 Then
                              Data1.Recordset("cl_fecing") = Adodc2.Recordset("fecha")
                              Data1.Recordset("cl_nombre") = Mid(Adodc2.Recordset("hc_nommed"), 1, 30)
                              Data1.Recordset("cl_nrocobr") = Adodc2.Recordset("hc_base")
                              Data1.Recordset("cl_dpto") = Adodc2.Recordset("cedtext")
                           End If
                           Data1.Recordset.Update
                        End If
                     End If
                  End If
               End If
            Else
                  If IsNull(adodc1.Recordset("hc_violencia")) = False Then
                     If adodc1.Recordset("hc_violencia") = 1 Then
                        Adodc2.RecordSource = "Select * from cabezal_hc where cb_mat =" & adodc1.Recordset("hc_mat")
                        Adodc2.Refresh
                        If Adodc2.Recordset.RecordCount > 0 Then
                           Data1.Recordset.AddNew
                           Data1.Recordset("cl_codconv") = Adodc2.Recordset("cb_codconv")
                           Data1.Recordset("cl_direcci") = "VD Positivo"
                           Data1.Recordset("cl_codigo") = Adodc2.Recordset("cb_mat")
                           If IsNull(Adodc2.Recordset("cb_fnac")) = False Then
                              Data1.Recordset("cl_fnac") = Adodc2.Recordset("cb_fnac")
                           End If
                           If IsNull(Adodc2.Recordset("cb_ape1")) = False Then
                              If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                                 Data1.Recordset("cl_apellid") = Mid(Adodc2.Recordset("cb_ape1"), 1, 15) & " " & Mid(Adodc2.Recordset("cb_nom1"), 1, 14)
                              Else
                                 Data1.Recordset("cl_apellid") = Adodc2.Recordset("cb_ape1")
                              End If
                           Else
                              If IsNull(Adodc2.Recordset("cb_nom1")) = False Then
                                 Data1.Recordset("cl_apellid") = "NN" & Adodc2.Recordset("cb_nom1")
                              Else
                                 Data1.Recordset("cl_apellid") = "NN"
                              End If
                           End If
                           Adodc2.RecordSource = "Select * from cabezal_hcdig where mat =" & adodc1.Recordset("hc_mat") & " and id =" & adodc1.Recordset("hc_nro")
                           Adodc2.Refresh
                           If Adodc2.Recordset.RecordCount > 0 Then
                              Data1.Recordset("cl_fecing") = Adodc2.Recordset("fecha")
                              Data1.Recordset("cl_nombre") = Mid(Adodc2.Recordset("hc_nommed"), 1, 30)
                              Data1.Recordset("cl_nrocobr") = Adodc2.Recordset("hc_base")
                              Data1.Recordset("cl_dpto") = Adodc2.Recordset("cedtext")
                           End If
                           Data1.Recordset.Update
                        End If
                     End If
                  End If
            
            End If
            adodc1.Recordset.MoveNext
         Loop
         frm_infhcemet.MousePointer = 0
         Command1.Enabled = True
         MsgBox "Proceso terminado"
         cr1.ReportFileName = App.Path & "\infhcemet1.rpt"
         cr1.ReportTitle = "Informe de VD y HTA Fecha: " & md.Text & " hasta:" & mh.Text
         cr1.Action = 1
      
      End If
   Else
      frm_infhcemet.MousePointer = 0
      Command1.Enabled = True
   End If

End If
frm_infhcemet.MousePointer = 0
Command1.Enabled = True


End Sub

Private Sub Form_Load()
adodc1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Adodc2.ConnectionString = "dsn=" & Xconexrmt

Data1.DatabaseName = App.Path & "\informes.mdb"


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
