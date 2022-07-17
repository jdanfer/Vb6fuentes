VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_fechasesp 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de fechas especialistas"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6285
   Icon            =   "frm_fechasesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6285
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox mfech 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
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
   Begin Crystal.CrystalReport cr1 
      Left            =   3000
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_fec 
      Caption         =   "data_fec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lista"
      Top             =   1560
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
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "inflista"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones de impresión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.Data data_esp 
         Caption         =   "data_esp"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Data data_medic 
         Caption         =   "data_medic"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox t_cmed 
         Height          =   285
         Left            =   4920
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cbomed 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   5055
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H008080FF&
         Caption         =   "Listar por médico seleccionado"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Solo datos de pediatría"
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
         Left            =   360
         TabIndex        =   10
         Top             =   3360
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Incluir Estudios de Análisis Clínicos"
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
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.CommandButton b_sale 
         BackColor       =   &H00C0C0FF&
         Caption         =   "SALIR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         MouseIcon       =   "frm_fechasesp.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "frm_fechasesp.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton b_acep 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ACEPTAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         MouseIcon       =   "frm_fechasesp.frx":0B8E
         MousePointer    =   99  'Custom
         Picture         =   "frm_fechasesp.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H008080FF&
         Caption         =   "Listar el especialista seleccionado por fecha"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   5055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H008080FF&
         Caption         =   "Listar todos los especialistas por rango de fecha"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frm_fechasesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
Dim Xedn As Long
If Option1.value = True Then
   If md.Text <> "__/__/____" Then
      If mh.Text <> "__/__/____" Then
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            Do While Not data_inf.Recordset.EOF
               data_inf.Recordset.Delete
               data_inf.Recordset.MoveNext
            Loop
         End If
         If Check2.value = 1 Then
            data_fec.RecordSource = "Select * from fechasesp where desc >='" & "PEDIATRIA " & "' and desc <='" & "PG" & "' and fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
            data_fec.Refresh
         Else
            data_fec.RecordSource = "Select * from fechasesp where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
            data_fec.Refresh
         End If
         If data_fec.Recordset.RecordCount > 0 Then
            Do While Not data_fec.Recordset.EOF
               data_inf.Recordset.AddNew
               data_inf.Recordset("base") = data_fec.Recordset("base")
               data_inf.Recordset("cod") = data_fec.Recordset("cod")
               data_inf.Recordset("desc") = data_fec.Recordset("desc")
               data_inf.Recordset("fecha") = data_fec.Recordset("fecha")
               data_inf.Recordset("descfec") = data_fec.Recordset("descfec")
''               data_inf.Recordset("horacom") = data_fec.Recordset("horacom")
               
               data_inf.Recordset.Update
               data_fec.Recordset.MoveNext
            Loop
            data_inf.RecordSource = "Select * from inflista order by base"
            data_inf.Refresh
            cr1.ReportFileName = App.Path & "\inffechas.rpt"
            cr1.Action = 1
         Else
            MsgBox "No hay registros para imprimir", vbInformation, "Mensaje"
         End If
      End If
   End If
End If
If Option2.value = True Then
   If mfec.Text <> "__/__/____" Then
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            data_inf.Recordset.Delete
            data_inf.Recordset.MoveNext
         Loop
      End If
      If mfech.Text = "__/__/____" Then
         If Check2.value = 1 Then
            data_fec.RecordSource = "Select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# And cl_fax ='" & frm_espec.txt_cod.Text & "'"
            data_fec.Refresh
         Else
            data_fec.RecordSource = "Select * from lista where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# And cod ='" & frm_espec.txt_cod.Text & "'"
            data_fec.Refresh
         End If
      Else
         If Check2.value = 1 Then
            data_fec.RecordSource = "Select * from mant_sol where cl_fnac >=#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mfech.Text, "yyyy/mm/dd") & "# And cl_fax ='" & frm_espec.txt_cod.Text & "'"
            data_fec.Refresh
         Else
            data_fec.RecordSource = "Select * from lista where fecha >=#" & Format(mfec.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mfech.Text, "yyyy/mm/dd") & "# And cod ='" & frm_espec.txt_cod.Text & "'"
            data_fec.Refresh
         End If
      End If
      If data_fec.Recordset.RecordCount > 0 Then
         Do While Not data_fec.Recordset.EOF
            If Check2.value = 1 Then
                data_inf.Recordset.AddNew
                data_inf.Recordset("nro") = data_fec.Recordset("cl_nrovend")
                If IsNull(data_fec.Recordset("cl_fnac")) = False Then
                   data_inf.Recordset("fecha") = data_fec.Recordset("cl_fnac")
                End If
                data_inf.Recordset("base") = data_fec.Recordset("cl_grupo")
                data_inf.Recordset("cod") = data_fec.Recordset("cl_fax")
                data_inf.Recordset("matric") = data_fec.Recordset("cl_atrasoa")
                If IsNull(data_fec.Recordset("cl_descpag")) = False Then
                   data_inf.Recordset("conve") = Mid(data_fec.Recordset("cl_descpag"), 1, 6)
                End If
                data_inf.Recordset("desc") = frm_espec.txt_desc.Text
                data_inf.Recordset("nompac") = data_fec.Recordset("cl_nom_sup")
                If IsNull(data_fec.Recordset("cl_fultmov")) = False Then
                   Xedn = Date - data_fec.Recordset("cl_fultmov")
                   Xedn = Xedn / 365
                   data_inf.Recordset("obs") = Format(data_fec.Recordset("cl_fultmov"), "dd/mm/yyyy") & " EDAD: " & Trim(Str(Xedn))
                End If
                data_inf.Recordset("tel") = data_fec.Recordset("cl_desc2")
                data_inf.Recordset("horacom") = data_fec.Recordset("cl_ruc")
                data_inf.Recordset("hc") = data_fec.Recordset("cl_codconv")
                If IsNull(data_fec.Recordset("cl_atrasop")) = False Then
                   If data_fec.Recordset("cl_atrasop") = 0 Then
                      data_inf.Recordset("descfec") = "META"
                   Else
                      If data_fec.Recordset("cl_atrasop") = 1 Then
                         data_inf.Recordset("descfec") = "CONTROL (NO META)"
                      Else
                         If data_fec.Recordset("cl_atrasop") = 2 Then
                            data_inf.Recordset("descfec") = "CONSULTA"
                         Else
                            If data_fec.Recordset("cl_atrasop") = 3 Then
                               data_inf.Recordset("descfec") = "RN (RECIEN NACIDO)"
                            Else
                               data_inf.Recordset("descfec") = "SIN REGISTRAR"
                            End If
                         End If
                      End If
                   End If
                Else
                   data_inf.Recordset("descfec") = "SIN REGISTRAR"
                End If
                If IsNull(data_fec.Recordset("cl_nomcobr")) = False Then
                   data_inf.Recordset("desc") = Trim(Str(data_fec.Recordset("cl_zona"))) & "-" & Trim(Str(data_fec.Recordset("cl_nomcobr")))
                Else
                   data_inf.Recordset("desc") = "0"
                End If
                If IsNull(data_fec.Recordset("cl_val1")) = False Then
                   data_inf.Recordset("edad") = Trim(Str(data_fec.Recordset("cl_val1"))) & " AÑOS " & Trim(Str(data_fec.Recordset("cl_val2"))) & " MESES " & Trim(Str(data_fec.Recordset("cl_val3"))) & " DIAS"
                End If
                data_inf.Recordset.Update
            Else
                data_inf.Recordset.AddNew
                data_inf.Recordset("nro") = data_fec.Recordset("nro")
                data_inf.Recordset("fecha") = data_fec.Recordset("fecha")
                data_inf.Recordset("base") = data_fec.Recordset("base")
                data_inf.Recordset("cod") = data_fec.Recordset("cod")
                data_inf.Recordset("matric") = data_fec.Recordset("matric")
                data_inf.Recordset("conve") = data_fec.Recordset("conve")
                data_inf.Recordset("desc") = data_fec.Recordset("desc")
                data_inf.Recordset("nompac") = data_fec.Recordset("nompac")
                data_inf.Recordset("descfec") = data_fec.Recordset("descfec")
                data_inf.Recordset("tel") = data_fec.Recordset("tel")
                data_inf.Recordset("horacom") = data_fec.Recordset("horacom")
                data_inf.Recordset("hc") = data_fec.Recordset("hc")
                data_inf.Recordset.Update
            End If
            data_fec.Recordset.MoveNext
          Loop
          data_inf.RecordSource = "Select * from inflista order by nro"
          data_inf.Refresh
          If Check1.value = 1 Then
             cr1.ReportFileName = App.Path & "\inflista2.rpt"
          Else
             If Check2.value = 1 Then
                cr1.ReportFileName = App.Path & "\inflista2p.rpt"
             Else
                If mfech.Text = "__/__/____" Then
                   cr1.ReportFileName = App.Path & "\inflista.rpt"
                Else
                   cr1.ReportFileName = App.Path & "\inflista4.rpt"
                End If
             End If
          End If
          cr1.Action = 1
      Else
          MsgBox "No hay registros para imprimir", vbInformation, "Mensaje"
      End If
   End If
End If
If Option3.value = True Then
   If cbomed.Text <> "" Then
      If md.Text <> "__/__/____" Then
         If mh.Text <> "__/__/____" Then
            If data_inf.Recordset.RecordCount > 0 Then
               data_inf.Recordset.MoveFirst
               Do While Not data_inf.Recordset.EOF
                  data_inf.Recordset.Delete
                  data_inf.Recordset.MoveNext
               Loop
            End If
            data_medic.RecordSource = "Select * from medicos where med_nombre ='" & cbomed.Text & "'"
            data_medic.Refresh
            If data_medic.Recordset.RecordCount > 0 Then
               data_fec.RecordSource = "Select * from fechasesp where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
               data_fec.Refresh
               If data_fec.Recordset.RecordCount > 0 Then
                  Do While Not data_fec.Recordset.EOF
                     data_esp.RecordSource = "Select * from lineas where hora ='" & data_fec.Recordset("cod") & "'"
                     data_esp.Refresh
                     If data_esp.Recordset.RecordCount > 0 Then
                        If data_esp.Recordset("cod_prod") = data_medic.Recordset("med_cod") Then
                           data_inf.Recordset.AddNew
                           data_inf.Recordset("base") = data_fec.Recordset("base")
                           data_inf.Recordset("cod") = data_fec.Recordset("cod")
                           data_inf.Recordset("desc") = data_fec.Recordset("desc")
                           data_inf.Recordset("fecha") = data_fec.Recordset("fecha")
                           data_inf.Recordset("descfec") = data_fec.Recordset("descfec")
                           data_inf.Recordset.Update
                        End If
                     End If
                     data_fec.Recordset.MoveNext
                  Loop
                  MsgBox "Proceso terminado"
                  data_inf.RecordSource = "Select * from inflista order by base"
                  data_inf.Refresh
                  cr1.ReportFileName = App.Path & "\inffechas.rpt"
                  cr1.Action = 1
               Else
                  MsgBox "No hay registros con la fecha seleccionada"
               End If
            Else
               MsgBox "No se encuentra el médico seleccionado"
            End If
         End If
      Else
         MsgBox "Ingrese un rango de fechas a imprimir"
      End If
   Else
      MsgBox "Seleccione un médico"
   End If
End If



End Sub

Private Sub b_sale_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_fec.DatabaseName = App.Path & "\sapp.mdb"
data_medic.DatabaseName = App.Path & "\sapp.mdb"
data_medic.RecordSource = "Select * from medicos order by med_nombre"
data_medic.Refresh
data_esp.DatabaseName = App.Path & "\sapp.mdb"

If data_medic.Recordset.RecordCount > 0 Then
   data_medic.Recordset.MoveFirst
   Do While Not data_medic.Recordset.EOF
      cbomed.AddItem data_medic.Recordset("med_nombre")
      data_medic.Recordset.MoveNext
   Loop
End If

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub

Private Sub Option1_Click()
md.Visible = True
mh.Visible = True
mfec.Visible = False
md.SetFocus

End Sub

Private Sub Option2_Click()
'md.Visible = False
'mh.Visible = False
mfec.Visible = True
mfec.SetFocus

End Sub
