VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infresesp 
   BackColor       =   &H0000C000&
   Caption         =   "Informes de reservas PEDIATRIA"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infresesp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   1800
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_esp 
      Caption         =   "data_esp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox t_b 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Text            =   "99"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Data data_lis 
      Caption         =   "data_lis"
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
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Resumen"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2760
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Detalle"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_infresesp.frx":058A
      Left            =   1680
      List            =   "frm_infresesp.frx":059A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   3375
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   375
      Left            =   3480
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
   Begin MSMask.MaskEdBox md 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
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
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(99=Todas)"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Base:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Informe de:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "FECHAS:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frm_infresesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
data_inf.RecordSource = "infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If

If mh.Text <> "__/__/____" And md.Text <> "__/__/____" Then
   frm_infresesp.MousePointer = 11
   If t_b.Text <> 99 Then
      data_lis.RecordSource = "Select * from mant_sol where cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# And cl_grupo =" & t_b.Text & " order by cl_fnac,cl_grupo"
      data_lis.Refresh
   Else
      data_lis.RecordSource = "Select * from mant_sol where cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by cl_fnac,cl_grupo"
      data_lis.Refresh
   End If
   If data_lis.Recordset.RecordCount > 0 Then
      data_lis.Recordset.MoveFirst
      Do While Not data_lis.Recordset.EOF
         If IsNull(data_lis.Recordset("cl_nom_sup")) = False Then
            If Trim(data_lis.Recordset("cl_nom_sup")) <> "" Then
                data_inf.Recordset.AddNew
                data_inf.Recordset("cl_codigo") = data_lis.Recordset("cl_nrovend") 'nro
                data_inf.Recordset("cl_fnac") = data_lis.Recordset("cl_fnac")
                data_inf.Recordset("cl_grupo") = data_lis.Recordset("cl_grupo") 'base
                data_inf.Recordset("cl_fax") = data_lis.Recordset("cl_fax") ' codigo
                data_inf.Recordset("cl_atrasoa") = data_lis.Recordset("cl_atrasoa") 'matricula
                data_inf.Recordset("cl_descpag") = data_lis.Recordset("cl_descpag") ' cod convenio
                data_esp.RecordSource = "Select * from espec where codigo ='" & Trim(data_lis.Recordset("cl_fax")) & "'"
                data_esp.Refresh
                If data_esp.Recordset.RecordCount > 0 Then
                   data_inf.Recordset("cl_apellid") = Mid(data_esp.Recordset("desc"), 1, 60)
                Else
                   data_inf.Recordset("cl_apellid") = "NO ENCONTRADO"
                End If
                data_inf.Recordset("cl_nom_sup") = Mid(data_lis.Recordset("cl_nom_sup"), 1, 25)
                data_inf.Recordset("cl_telefon") = Mid(data_lis.Recordset("cl_desc2"), 1, 20)
                data_inf.Recordset("cl_ruc") = data_lis.Recordset("cl_ruc")
                data_inf.Recordset("cl_codconv") = data_lis.Recordset("cl_codconv")
                data_inf.Recordset("cl_Cedula") = data_lis.Recordset("cl_zona")
                data_inf.Recordset("cl_codced") = data_lis.Recordset("cl_nomcobr")
                If IsNull(data_lis.Recordset("cl_atrasop")) = False Then
                   If data_lis.Recordset("cl_atrasop") = 0 Then
                      data_inf.Recordset("cl_nomcobr") = "META"
                   Else
                      If data_lis.Recordset("cl_atrasop") = 1 Then
                         data_inf.Recordset("cl_nomcobr") = "CONTROL (NO META)"
                      Else
                         If data_lis.Recordset("cl_atrasop") = 2 Then
                            data_inf.Recordset("cl_nomcobr") = "CONSULTA"
                         Else
                            If data_lis.Recordset("cl_atrasop") = 3 Then
                               data_inf.Recordset("cl_nomcobr") = "RN (RECIEN NACIDO)"
                            Else
                               data_inf.Recordset("cl_nomcobr") = "SIN REGISTRAR"
                            End If
                         End If
                      End If
                   End If
                Else
                   data_inf.Recordset("cl_nomcobr") = "SIN REGISTRAR"
                End If
                data_inf.Recordset("cl_numero") = data_lis.Recordset("cl_numero")
                If IsNull(data_lis.Recordset("cl_numero")) = False Then
                   If data_lis.Recordset("cl_numero") = 2 Then
                      data_inf.Recordset("cl_nombre") = "FALTA S/AVISO"
                   End If
                End If
                data_inf.Recordset("cl_fultmov") = data_lis.Recordset("cl_fultmov")
                Dim Xlaeda As Long
                Dim Xquee As String
                If IsNull(data_lis.Recordset("cl_fultmov")) = False Then
                   Xlaeda = data_lis.Recordset("cl_fnac") - data_lis.Recordset("cl_fultmov")
                   If Xlaeda > 365 Then
                      Xlaeda = Xlaeda / 365
                      Xquee = "A"
                   Else
                      If Xlaeda > 30 Then
                         Xlaeda = Xlaeda / 30
                         Xquee = "M"
                      Else
                         Xquee = "D"
                      End If
                   End If
                End If
                data_inf.Recordset("cl_cantdia") = Xlaeda
                data_inf.Recordset("cl_dpto") = Xquee
                
                data_inf.Recordset.Update
            End If
        End If
        data_lis.Recordset.MoveNext
      Loop
      If Combo1.ListIndex = 1 Then
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            Do While Not data_inf.Recordset.EOF
               If IsNull(data_inf.Recordset("cl_numero")) = True Then
                  data_inf.Recordset.Delete
               Else
                  If data_inf.Recordset("cl_numero") <> 2 Then
                     data_inf.Recordset.Delete
                  End If
               End If
               data_inf.Recordset.MoveNext
            Loop
            data_inf.Refresh
         End If
      End If
      frm_infresesp.MousePointer = 0
      MsgBox "Proceso terminado"
      cr1.ReportFileName = App.Path & "\infresespd.rpt"
      If Combo1.ListIndex = 1 Then
         cr1.ReportTitle = "Informe de pacientes que FALTARON a la CONSULTA desde: " & md.Text & " hasta: " & mh.Text
      Else
         cr1.ReportTitle = "Informe de pacientes anotados para PEDIATRIA desde: " & md.Text & " hasta: " & mh.Text
      End If
      cr1.Action = 1
   Else
      MsgBox "No hay registros"
   End If

End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_lis.DatabaseName = App.Path & "\sapp.mdb"
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_esp.DatabaseName = App.Path & "\sapp.mdb"

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
