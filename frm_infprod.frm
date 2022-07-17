VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infprod 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de produccion"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5115
   Icon            =   "frm_infprod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_deu 
      Caption         =   "data_deu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_zona 
      Caption         =   "data_zona"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3060
   End
   Begin Crystal.CrystalReport crpro 
      Left            =   4680
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton b_can 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Picture         =   "frm_infprod.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton b_acep 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "frm_infprod.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Procesar"
      Top             =   6120
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones de informe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   375
         Left            =   360
         Top             =   720
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
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
         Caption         =   "data_cli"
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
      Begin MSAdodcLib.Adodc data_em 
         Height          =   375
         Left            =   1560
         Top             =   1920
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
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
         Caption         =   "data_em"
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
      Begin VB.TextBox t_zon 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   3240
         Width           =   735
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Detalle de socios por zona con UMP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   5520
         Width           =   3615
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Detalle de todos los convenios"
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
         Left            =   480
         TabIndex        =   14
         Top             =   4080
         Width           =   3375
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Seleccionar cobrador"
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
         TabIndex        =   13
         Top             =   5160
         Width           =   3615
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C00000&
         Caption         =   "Ordenar por Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cobranza de Bases"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4560
         Width           =   3375
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Sin San Jacinto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Sin Bases"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Producción por Convenio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3720
         Width           =   3375
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Producción por Zona"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2880
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Producción por Cobrador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txt_a 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txt_m 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Zona: (Opcional)"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "MES/AÑO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1440
      Picture         =   "frm_infprod.frx":0F56
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2655
   End
End
Attribute VB_Name = "frm_infprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub b_acep_Click()
Dim Nombre, Nomcob As String
Dim Xcob As Integer
Dim Totrec, Totpesos As Long
Dim Xtot, XPor As Double
Dim Xfecnuevas As String

'On Error GoTo Noesta
frm_infprod.MousePointer = 11
b_acep.Enabled = False
b_can.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infemis"
data_inf.RecordSource = "infemis"
data_inf.Refresh
If Option1.Value = True Then
   If txt_m.Text <> "" Then
      If txt_a.Text <> "" Then
         If txt_m.Text > 9 Then
            Nombre = "emi" & Trim(txt_m.Text) & Mid(Trim(txt_a.Text), 3, 2)
         Else
            Nombre = "emi" & "0" & Trim(txt_m.Text) & Mid(Trim(txt_a.Text), 3, 2)
         End If
         data_em.RecordSource = "Select * from " & Nombre & " order by nro_cobr"
         data_em.Refresh
         If data_em.Recordset.RecordCount > 0 Then
            data_em.Recordset.MoveLast
            Xtot = data_em.Recordset.RecordCount
            data_em.Recordset.MoveFirst
            Xcob = data_em.Recordset("nro_cobr")
            Nomcob = data_em.Recordset("nom_cobr")
            Do While Not data_em.Recordset.EOF
               If data_em.Recordset("nro_cobr") = Xcob Then
                  Totrec = Totrec + 1
                  Totpesos = Totpesos + data_em.Recordset("total")
                  Xcob = data_em.Recordset("nro_cobr")
                  Nomcob = data_em.Recordset("nom_cobr")
                  data_em.Recordset.MoveNext
               Else
                  XPor = Totrec / Xtot
                  XPor = XPor * 100
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("nro_cobr") = Xcob
                  data_inf.Recordset("nom_cobr") = Nomcob
                  data_inf.Recordset("mes") = txt_m.Text
                  data_inf.Recordset("ano") = txt_a.Text
                  data_inf.Recordset("importe") = Format(XPor, "Standard")
                  data_inf.Recordset("documento") = Totrec
                  data_inf.Recordset("total") = Totpesos
                  data_inf.Recordset.Update
                  Totrec = 0
                  Totpesos = 0
                  XPor = 0
                  Xcob = data_em.Recordset("nro_cobr")
                  Nomcob = data_em.Recordset("nom_cobr")
               End If
            Loop
            XPor = Totrec / Xtot
            XPor = XPor * 100
            data_inf.Recordset.AddNew
            data_inf.Recordset("nro_cobr") = Xcob
            data_inf.Recordset("nom_cobr") = Nomcob
            data_inf.Recordset("mes") = txt_m.Text
            data_inf.Recordset("ano") = txt_a.Text
            data_inf.Recordset("importe") = Format(XPor, "Standard")
            data_inf.Recordset("documento") = Totrec
            data_inf.Recordset("total") = Totpesos
            data_inf.Recordset.Update
            Totrec = 0
            Totpesos = 0
            XPor = 0
            data_inf.RecordSource = "Select * from infemis order by nro_cobr"
            data_inf.Refresh
            crpro.ReportTitle = "INFORME ORDENADO POR COBRADOR"
            crpro.ReportFileName = App.path & "\infprod.rpt"
            crpro.Action = 1
         End If
      End If
   End If
End If
If Option2.Value = True Then
   If Check6.Value = 1 Then
      If txt_m.Text <> "" Then
         If txt_a.Text <> "" Then
            If txt_m.Text > 9 Then
               Nombre = "emi" & Trim(txt_m.Text) & Mid(Trim(txt_a.Text), 3, 2)
            Else
               Nombre = "emi" & "0" & Trim(txt_m.Text) & Mid(Trim(txt_a.Text), 3, 2)
            End If
            If t_zon.Text = "" Then
               data_em.RecordSource = "Select * from " & Nombre & " order by grupo"
               data_em.Refresh
            Else
               data_em.RecordSource = "Select * from " & Nombre & " where grupo =" & t_zon.Text & " order by grupo"
               data_em.Refresh
            End If
            If data_em.Recordset.RecordCount > 0 Then
               data_em.Recordset.MoveFirst
               Do While Not data_em.Recordset.EOF
                  data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_em.Recordset("cliente")
                  data_cli.Refresh
                  If data_cli.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cliente") = data_em.Recordset("cliente")
                     data_inf.Recordset("cod_cnv") = data_em.Recordset("cod_cnv")
                     data_inf.Recordset("nom_cnv") = data_em.Recordset("nom_cnv")
                     data_inf.Recordset("apellidos") = data_em.Recordset("apellidos")
                     data_inf.Recordset("grupo") = data_em.Recordset("grupo")
                     data_inf.Recordset("nom_cobr") = data_cli.Recordset("cl_socmnom")
                     data_inf.Recordset("zona") = data_em.Recordset("zona")
                     data_inf.Recordset("fecha_ing") = data_em.Recordset("fecha_ing")
                     data_inf.Recordset("mes") = 0
                     data_inf.Recordset("ano") = 0
                     data_inf.Recordset.Update
                  End If
                  data_em.Recordset.MoveNext
               Loop
               MsgBox "Proceso terminado"
            End If
         End If
      End If
   Else
       Dim Xcolre As String
       If txt_m.Text <> "" Then
          If txt_a.Text <> "" Then
             If txt_m.Text > 9 Then
                Nombre = "emi" & Trim(txt_m.Text) & Mid(Trim(txt_a.Text), 3, 2)
                Xfecnuevas = "02/" & txt_m.Text & "/" & txt_a.Text
             Else
                Nombre = "emi" & "0" & Trim(txt_m.Text) & Mid(Trim(txt_a.Text), 3, 2)
                Xfecnuevas = "02/0" & txt_m.Text & "/" & txt_a.Text
             End If
             If Check2.Value = 1 Then
                If Check3.Value = 1 Then
                   data_em.RecordSource = "Select * from " & Nombre & " where nro_cobr <>" & 11 & " And nro_cobr <>" & 6 & " and nro_cobr <>" & 5 & " and fecha <'" & Format(Xfecnuevas, "yyyy-mm-dd") & "' order by color_rec, grupo"
                   data_em.Refresh
                Else
                   data_em.RecordSource = "Select * from " & Nombre & " where nro_cobr <>" & 11 & " And nro_cobr <>" & 6 & " and nro_cobr <>" & 5 & " and fecha <'" & Format(Xfecnuevas, "yyyy-mm-dd") & "' order by grupo"
                   data_em.Refresh
                End If
             Else
                If Check3.Value = 1 Then
                   data_em.RecordSource = "Select * from " & Nombre & " where fecha <'" & Format(Xfecnuevas, "yyyy-mm-dd") & "' order by color_rec, grupo"
                   data_em.Refresh
                Else
                   If Check4.Value = 1 Then
                      Dim Xquecobtex As String
                      Xquecobtex = InputBox("Ingrese número de cobrador:", "Cobrador")
                      If Xquecobtex = "" Then
                         Xquecobtex = "0"
                      End If
                      data_em.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & Val(Xquecobtex) & " order by grupo"
                      data_em.Refresh
                   Else
                      data_em.RecordSource = "Select * from " & Nombre & " where fecha <'" & Format(Xfecnuevas, "yyyy-mm-dd") & "' order by grupo"
                      data_em.Refresh
                   End If
                End If
             End If
             
             If data_em.Recordset.RecordCount > 0 Then
                data_em.Recordset.MoveFirst
                Do While Not data_em.Recordset.EOF
                   If IsNull(data_em.Recordset("grupo")) = True Then
'                      data_em.Recordset.Edit
                      data_em.Recordset("grupo") = 0
                      data_em.Recordset("zona") = "S/Zona"
                      data_em.Recordset.Update
                   End If
                   data_em.Recordset.MoveNext
                Loop
    '            data_em.Recordset.MoveLast
                Xtot = data_em.Recordset.RecordCount
                Xtot = 0
                Xtot = data_em.Recordset.RecordCount
                data_em.Recordset.MoveFirst
                Xcob = data_em.Recordset("grupo")
                Nomcob = data_em.Recordset("zona")
                Xcolre = data_em.Recordset("color_rec")
                Do While Not data_em.Recordset.EOF
                   If Check1.Value = 1 Then
                      If data_em.Recordset("nro_cobr") = 1 Or data_em.Recordset("nro_cobr") = 5 Or _
                         data_em.Recordset("nro_cobr") = 6 Or data_em.Recordset("nro_cobr") = 8 Or _
                         data_em.Recordset("nro_cobr") = 10 Or data_em.Recordset("nro_cobr") = 11 Or _
                         data_em.Recordset("nro_cobr") = 113 Or data_em.Recordset("nro_cobr") = 201 Or _
                         data_em.Recordset("nro_cobr") = 512 Or data_em.Recordset("nro_cobr") = 602 Or data_em.Recordset("nro_cobr") = 679 Or _
                         data_em.Recordset("nro_cobr") = 603 Or data_em.Recordset("nro_cobr") = 604 Or _
                         data_em.Recordset("nro_cobr") = 606 Or data_em.Recordset("nro_cobr") = 615 Or _
                         data_em.Recordset("nro_cobr") = 616 Or data_em.Recordset("nro_cobr") = 635 Or _
                         data_em.Recordset("nro_cobr") = 636 Or data_em.Recordset("nro_cobr") = 653 Or data_em.Recordset("nro_cobr") = 112 Or _
                         data_em.Recordset("nro_cobr") = 672 Or data_em.Recordset("nro_cobr") = 685 Then
                         data_em.Recordset.MoveNext
                      Else
                        If data_em.Recordset("grupo") = Xcob Then
                           Totrec = Totrec + 1
                           Totpesos = Totpesos + data_em.Recordset("total")
                           Xcob = data_em.Recordset("grupo")
                           If IsNull(data_em.Recordset("zona")) = False Then
                              Nomcob = data_em.Recordset("zona")
                           Else
                              Nomcob = "Sin zona"
                           End If
                           Xcolre = data_em.Recordset("color_rec")
                           data_em.Recordset.MoveNext
                        Else
                           XPor = Totrec / Xtot
                           XPor = XPor * 100
                           data_inf.Recordset.AddNew
                           data_inf.Recordset("nro_cobr") = Xcob
                           data_inf.Recordset("nom_cobr") = Nomcob
                           data_inf.Recordset("mes") = txt_m.Text
                           data_inf.Recordset("ano") = txt_a.Text
                           data_inf.Recordset("importe") = Format(XPor, "Standard")
                           data_inf.Recordset("documento") = Totrec
                           data_inf.Recordset("total") = Totpesos
                           data_inf.Recordset("color_rec") = Xcolre
                           data_inf.Recordset.Update
                           Totrec = 0
                           Totpesos = 0
                           XPor = 0
                           Xcob = data_em.Recordset("grupo")
                           Nomcob = data_em.Recordset("zona")
                           Xcolre = data_em.Recordset("color_rec")
                        End If
                      End If
                   Else
                      If data_em.Recordset("grupo") = Xcob Then
                         Totrec = Totrec + 1
                         Totpesos = Totpesos + data_em.Recordset("total")
                         Xcob = data_em.Recordset("grupo")
                         If IsNull(data_em.Recordset("zona")) = False Then
                            Nomcob = data_em.Recordset("zona")
                         Else
                            Nomcob = "Sin zona"
                         End If
                         If IsNull(data_em.Recordset("color_rec")) = False Then
                            Xcolre = data_em.Recordset("color_rec")
                         Else
                            MsgBox "ATENCION!! Recibo sin COLOR VERIFIQUE!! MAT:" & data_em.Recordset("cliente")
                         End If
                         data_em.Recordset.MoveNext
                      Else
                         XPor = Totrec / Xtot
                         XPor = XPor * 100
                         data_inf.Recordset.AddNew
                         data_inf.Recordset("nro_cobr") = Xcob
                         data_inf.Recordset("nom_cobr") = Nomcob
                         data_inf.Recordset("mes") = txt_m.Text
                         data_inf.Recordset("ano") = txt_a.Text
                         data_inf.Recordset("importe") = Format(XPor, "Standard")
                         data_inf.Recordset("documento") = Totrec
                         data_inf.Recordset("total") = Totpesos
                         data_inf.Recordset("color_rec") = Xcolre
                         data_inf.Recordset.Update
                         Totrec = 0
                         Totpesos = 0
                         XPor = 0
                         If IsNull(data_em.Recordset("grupo")) = False Then
                            Xcob = data_em.Recordset("grupo")
                            Nomcob = data_em.Recordset("zona")
                         Else
                            Xcob = 0
                            Nomcob = "S/Zona"
                         End If
                         Xcolre = data_em.Recordset("color_rec")
                      End If
                   End If
                Loop
                XPor = Totrec / Xtot
                XPor = XPor * 100
                data_inf.Recordset.AddNew
                data_inf.Recordset("nro_cobr") = Xcob
                data_inf.Recordset("nom_cobr") = Nomcob
                data_inf.Recordset("mes") = txt_m.Text
                data_inf.Recordset("ano") = txt_a.Text
                data_inf.Recordset("importe") = Format(XPor, "Standard")
                data_inf.Recordset("documento") = Totrec
                data_inf.Recordset("total") = Totpesos
                data_inf.Recordset("color_rec") = Xcolre
                data_inf.Recordset.Update
                Totrec = 0
                Totpesos = 0
                XPor = 0
                data_inf.RecordSource = "Select * from infemis order by grupo"
                data_inf.Refresh
                crpro.ReportTitle = "INFORME ORDENADO POR ZONAS"
                crpro.ReportFileName = App.path & "\infprod.rpt"
                crpro.Action = 1
             End If
          End If
       End If
   End If
End If
If Option3.Value = True Then
   Dim XCat As Integer
   XCat = 1
   If txt_m.Text <> "" Then
      If txt_a.Text <> "" Then
         If txt_m.Text > 9 Then
            Nombre = "emi" & Trim(txt_m.Text) & Mid(Trim(txt_a.Text), 3, 2)
         Else
            Nombre = "emi" & "0" & Trim(txt_m.Text) & Mid(Trim(txt_a.Text), 3, 2)
         End If
         data_em.RecordSource = "Select * from " & Nombre & " order by cod_cnv"
         data_em.Refresh
         Dim Xlacat As String
         Dim Xlacatn As String
         Dim Xfecing As Date
         If data_em.Recordset.RecordCount > 0 Then
            data_em.Recordset.MoveLast
            Xtot = data_em.Recordset.RecordCount
            data_em.Recordset.MoveFirst
            If Check5.Value = 1 Then
               Xlacat = data_em.Recordset("cod_cnv")
               Xlacatn = data_em.Recordset("nom_cnv")
                Do While Not data_em.Recordset.EOF
                   If Trim(data_em.Recordset("cod_cnv")) = Trim(Xlacat) Then
                      Totrec = Totrec + 1
                      Totpesos = Totpesos + data_em.Recordset("total")
                      Xlacatn = data_em.Recordset("nom_cnv")
                      Xlacat = data_em.Recordset("cod_cnv")
                      If IsNull(data_em.Recordset("fecha_ing")) = False Then
                          Xfecing = data_em.Recordset("fecha_ing")
                      Else
                          Xfecing = vbEmpty
                      End If
                      data_em.Recordset.MoveNext
                   Else
                      XPor = Totrec / Xtot
                      XPor = XPor * 100
                      data_inf.Recordset.AddNew
                      data_inf.Recordset("cod_cnv") = Xlacat
                      data_inf.Recordset("nom_cnv") = Xlacatn
                      data_inf.Recordset("mes") = txt_m.Text
                      data_inf.Recordset("ano") = txt_a.Text
                      data_inf.Recordset("importe") = Format(XPor, "Standard")
                      data_inf.Recordset("documento") = Totrec
                      data_inf.Recordset("total") = Totpesos
                      data_inf.Recordset("fecha_ing") = Xfecing
                      data_em.Recordset.MovePrevious
                      data_inf.Recordset("tiquet") = data_em.Recordset("importe")
                      data_em.Recordset.MoveNext
                      data_inf.Recordset.Update
                      
                      Totrec = 0
                      Totpesos = 0
                      XPor = 0
'                      XCat = XCat + 1
                        
                      Xlacat = data_em.Recordset("cod_cnv")
                      Xlacatn = data_em.Recordset("nom_cnv")
                   End If
                Loop
                XPor = Totrec / Xtot
                XPor = XPor * 100
                data_inf.Recordset.AddNew
                data_inf.Recordset("cod_cnv") = Xlacat
                data_inf.Recordset("nom_cnv") = Xlacatn
                data_inf.Recordset("mes") = txt_m.Text
                data_inf.Recordset("ano") = txt_a.Text
                data_inf.Recordset("importe") = Format(XPor, "Standard")
                data_inf.Recordset("documento") = Totrec
                data_inf.Recordset("total") = Totpesos
                data_inf.Recordset("fecha_ing") = Xfecing
                data_inf.Recordset.Update
                Totrec = 0
                Totpesos = 0
                XPor = 0
                XCat = 0
                data_inf.RecordSource = "Select * from infemis order by cod_cnv"
                data_inf.Refresh
                crpro.ReportTitle = "INFORME ORDENADO POR CATEGORIA"
                crpro.ReportFileName = App.path & "\infprodcd.rpt"
                crpro.Action = 1
            Else
                Xcob = XCat
                Nomcob = data_em.Recordset("color_rec")
                
                Do While Not data_em.Recordset.EOF
                   If data_em.Recordset("color_rec") = Trim(Nomcob) Then
                      Totrec = Totrec + 1
                      Totpesos = Totpesos + data_em.Recordset("total")
                      Xcob = XCat
                      Nomcob = data_em.Recordset("color_rec")
                      data_em.Recordset.MoveNext
                   Else
                      XPor = Totrec / Xtot
                      XPor = XPor * 100
                      data_inf.Recordset.AddNew
                      data_inf.Recordset("nro_cobr") = XCat
                      If Trim(Nomcob) = "A" Then
                         Nomcob = "PARCIAL"
                      End If
                      If Trim(Nomcob) = "M" Then
                         Nomcob = "EMERGENCIA"
                      End If
                      If Trim(Nomcob) = "R" Then
                         Nomcob = "AMBULATORIO"
                      End If
                      If Trim(Nomcob) = "V" Then
                         Nomcob = "EMERGENCIA CGAL"
                      End If
                      data_inf.Recordset("nom_cobr") = Trim(Nomcob)
                      data_inf.Recordset("mes") = txt_m.Text
                      data_inf.Recordset("ano") = txt_a.Text
                      data_inf.Recordset("importe") = Format(XPor, "Standard")
                      data_inf.Recordset("documento") = Totrec
                      data_inf.Recordset("total") = Totpesos
                      data_inf.Recordset.Update
                      Totrec = 0
                      Totpesos = 0
                      XPor = 0
                      XCat = XCat + 1
                      Xcob = XCat
                      If IsNull(data_em.Recordset("color_rec")) = False Then
                         Nomcob = data_em.Recordset("color_rec")
                      Else
                         Nomcob = "X"
                      End If
                   End If
                Loop
                XPor = Totrec / Xtot
                XPor = XPor * 100
                If Trim(Nomcob) = "A" Then
                   Nomcob = "PARCIAL"
                End If
                If Trim(Nomcob) = "M" Then
                   Nomcob = "EMERGENCIA"
                End If
                If Trim(Nomcob) = "R" Then
                   Nomcob = "AMBULATORIO"
                End If
                If Trim(Nomcob) = "V" Then
                   Nomcob = "EMERGENCIA CGAL"
                End If
                data_inf.Recordset.AddNew
                data_inf.Recordset("nro_cobr") = XCat
                data_inf.Recordset("nom_cobr") = Trim(Nomcob)
                data_inf.Recordset("mes") = txt_m.Text
                data_inf.Recordset("ano") = txt_a.Text
                data_inf.Recordset("importe") = Format(XPor, "Standard")
                data_inf.Recordset("documento") = Totrec
                data_inf.Recordset("total") = Totpesos
                data_inf.Recordset.Update
                Totrec = 0
                Totpesos = 0
                XPor = 0
                XCat = 0
                data_inf.RecordSource = "Select * from infemis order by color_rec"
                data_inf.Refresh
                crpro.ReportTitle = "INFORME ORDENADO POR COLOR RECIBO"
                crpro.ReportFileName = App.path & "\infprod.rpt"
                crpro.Action = 1
            End If
         End If
      End If
   End If
End If
If Option4.Value = True Then
'   frm_infprod.MousePointer = 11

'   Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\infdeuda.mdb")

'   MiBaseact.Execute "Delete * from infdeuda"
   
'   data_lin.DatabaseName = App.Path & "\infdeuda.mdb"
'   data_lin.RecordSource = "infdeuda"
'   data_lin.Refresh
   
'   If data_deu.Recordset.RecordCount > 0 Then
'      data_deu.Recordset.MoveFirst
'      Do While Not data_lin.Recordset.EOF
'         If IsNull(data_lin.Recordset("mes_paga")) = False Then
'            If IsNull(data_lin.Recordset("ano_paga")) = False Then
'               If IsNull(data_lin.Recordset("cod_cli")) = False Then
'                    data_deu.RecordSource "select * from deudas where cliente =" & data_lin.Recordset("cod_cli") & " and mes =" & data_lin.Recordset("mes_paga") & " and ano =" & data_lin.Recordset("ano_paga")
'                    data_deu.Refresh
'                    If data_deu.Recordset.RecordCount > 0 Then
'                       data_deu.Recordset.Edit
'                       data_deu.Recordset("fecha_pago") = data_lin.Recordset("fecha")
'                       data_deu.Recordset.Update
'                    End If
'               End If
'            End If
 '        End If
 '        data_lin.Recordset.MoveNext
'      Loop
'   End If
   frm_infprod.MousePointer = 0
   MsgBox "No habilitado"
   
End If
frm_infprod.MousePointer = 0
b_acep.Enabled = True
b_can.Enabled = True

frm_infprod.MousePointer = 0

'Exit Sub

'Noesta:
'       If Err.Number = 3078 Then
'          MsgBox "No existe emisión, VERIFIQUE!!", vbCritical, "Mensaje"
'          txt_m.SetFocus
'       Else
'          MsgBox "Hay un error, VERIFIQUE!!", vbCritical, "Mensaje"
'          txt_m.SetFocus
'       End If
       
End Sub

Private Sub b_can_Click()
Unload Me

End Sub

Private Sub Form_Load()
txt_m.Text = Month(Date)
txt_a.Text = Year(Date)
'data_em.DatabaseName = App.Path & "\emisiones.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
data_zona.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_zona.RecordSource = "zonas"
data_zona.Refresh
'data_em.DatabaseName = ""
'data_em.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_em.ConnectionString = "dsn=" & Xconexrmt
data_deu.DatabaseName = App.path & "\deubase.mdb"
data_deu.RecordSource = "deudas"
data_deu.Refresh
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
