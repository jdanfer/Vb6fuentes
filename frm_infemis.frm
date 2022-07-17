VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infemis 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Emisión"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5655
   Icon            =   "frm_infemis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_em 
      Height          =   735
      Left            =   3480
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      ConnectStringType=   3
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
   Begin VB.Data data_inf2 
      Caption         =   "data_inf2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Crystal.CrystalReport cremis 
      Left            =   2880
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Data data_em2 
      Caption         =   "data_em2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   5175
      Begin MSAdodcLib.Adodc adocnv 
         Height          =   330
         Left            =   3120
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "adocnv"
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
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Emitir solo facturas con RUT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox t_iva 
         Height          =   375
         Left            =   3960
         TabIndex        =   18
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "EMITIR INFORME DE NUMERACION DE RECIBOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   4695
      End
      Begin MSMask.MaskEdBox mf 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fecha para nuevas entregas"
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
         TabIndex        =   15
         Top             =   2760
         Width           =   3375
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_infemis.frx":0442
         Left            =   1680
         List            =   "frm_infemis.frx":0458
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frm_infemis.frx":04AC
         Left            =   1680
         List            =   "frm_infemis.frx":04BC
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C00000&
         Caption         =   "Cobrador/Color/Valor"
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
         Top             =   360
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "INFORME:"
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
         Left            =   600
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Cobrador:"
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
         Left            =   600
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
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
      Left            =   4800
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frm_infemis.frx":04EC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   5640
      Width           =   615
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
      Picture         =   "frm_infemis.frx":0A76
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Procesar"
      Top             =   5640
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para informe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5175
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "SIN BASES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "DATOS DESDE SIMULACION"
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
         TabIndex        =   7
         Top             =   1440
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "DATOS DESDE EMISION"
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
         TabIndex        =   3
         Top             =   960
         Value           =   -1  'True
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
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "MES/AÑO:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1080
      Picture         =   "frm_infemis.frx":1000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2415
   End
End
Attribute VB_Name = "frm_infemis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
Dim Nombre, Quecob As String

On Error GoTo Noesta
'CrystalReport1.WindowShowRefreshBtn = True
t_iva.Text = ""
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infemis"
data_inf.Refresh
If mf.Text = "__/__/____" Then
   mf.Text = Date
End If

frm_infemis.MousePointer = 11
If txt_m.Text <> "" And mf.Text <> "__/__/____" Then
   If txt_a.Text <> "" Then
      If Option1.Value = True Then
         Nombre = "emi"
         If txt_m.Text > 9 Then
            Nombre = Nombre + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
         Else
            Nombre = Nombre + "0" + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
         End If
'''         data_em.Connect = "odbc;dsn=" & Xconexrmt & ";"
'         data_em.RecordSource = Nombre
'         data_em.Refresh
         If Combo1.ListIndex = 0 Then
            If Check1.Value = 1 Then
               data_em.RecordSource = "Select * from " & Nombre & " where nro_cobr <>" & 616 & " And nro_cobr <>" & 603 & " And nro_cobr <>" & 208 & _
               " And nro_cobr<>" & 636 & " And nro_cobr<>" & 615 & " And nro_cobr<>" & 635 & " And nro_cobr<>" & 201 & " And nro_cobr <>" & 209 & _
               " and nro_cobr<>" & 602 & " And nro_cobr<>" & 653 & " And nro_cobr<>" & 672 & " And nro_cobr<>" & 679 & " And nro_cobr<>" & 112 & _
               " and nro_cobr<>" & 8 & " And nro_cobr<>" & 685 & " And nro_cobr<>" & 512 & " And nro_cobr <>" & 606 & " And nro_cobr <>" & 6 & " And nro_cobr <>" & 11 & " And nro_cobr <>" & 5 & _
               " and nro_cobr<>" & 113 & " And nro_cobr<>" & 1 & " and nro_cobr<>" & 10 & " order by nro_cobr,color_rec,documento,importe"
               data_em.Refresh
            Else
               If Check2.Value = 1 Then
                  data_em.RecordSource = "Select * from " & Nombre & " order by nro_cobr,color_rec,documento,importe"
                  data_em.Refresh
               Else
                  If Combo2.ListIndex = 4 Then
                     data_em.RecordSource = "select " & Nombre & ".cod_cnv," & Nombre & ".nom_cnv," & Nombre & ".cliente," & Nombre & ".apellidos," & Nombre & ".documento," & Nombre & ".fecha_ing," & Nombre & ".tel_cli," & Nombre & ".deudaap," & _
                     Nombre & ".fecha," & Nombre & ".importe," & Nombre & ".nro_cobr," & Nombre & ".nom_cobr," & Nombre & ".grupo," & Nombre & ".zona," & Nombre & ".descimp," & Nombre & ".mes," & Nombre & ".ano," & _
                     Nombre & ".color_rec," & Nombre & ".tiquet," & Nombre & ".servi," & Nombre & ".deudas," & Nombre & ".iva," & Nombre & ".total," & Nombre & ".ruc," & _
                     "convenio.cnv_codigo,convenio.cnv_cant_r from " & Nombre & _
                     " inner join convenio on " & Nombre & ".cod_cnv=convenio.cnv_codigo where convenio.cnv_cant_r not in (1) and " & Nombre & ".nro_cobr not in (5,11,6) and " & Nombre & ".total >" & 0
                     data_em.Refresh
                  Else
                     data_em.RecordSource = "Select * from " & Nombre & " order by nro_cobr,color_rec,documento,importe"
                     data_em.Refresh
                  End If
               End If
            End If
         Else
            If Combo1.ListIndex = 2 Then
               data_em.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & 5 & " or nro_cobr =" & 11 & " or nro_cobr =" & 6 & " order by nro_cobr,color_rec,documento,importe"
               data_em.Refresh
            Else
               Quecob = InputBox("Ingrese número de cobrador:", "Solicitud de datos")
               If Quecob <> "" Then
                  data_em.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & Quecob & " order by color_rec,documento,importe"
                  data_em.Refresh
               Else
                  data_em.RecordSource = "Select * from " & Nombre & " order by nro_cobr,color_rec,documento,importe"
                  data_em.Refresh
               End If
            End If
         End If
         If Combo1.ListIndex = 3 Then
               data_em.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & 616 & " or nro_cobr =" & 603 & _
               " or nro_cobr =" & 636 & " or nro_cobr =" & 615 & " or nro_cobr =" & 635 & " or nro_cobr =" & 201 & " or nro_cobr =" & 208 & _
               " or nro_cobr =" & 602 & " or nro_cobr =" & 653 & " or nro_cobr =" & 672 & " or nro_cobr =" & 679 & " or nro_cobr =" & 209 & _
               " or nro_cobr =" & 8 & " or nro_cobr =" & 685 & " or nro_cobr =" & 512 & " or nro_cobr =" & 606 & " or nro_cobr =" & 112 & _
               " or nro_cobr =" & 113 & " or nro_cobr =" & 1 & " or nro_cobr =" & 10 & " order by nro_cobr,color_rec,documento,importe"
               data_em.Refresh
         End If
         If Check3.Value = 1 Then
            data_em.RecordSource = "Select * from " & Nombre & " where ruc is not null order by nro_cobr,color_rec,documento,importe"
            data_em.Refresh
         End If
         If data_em.Recordset.RecordCount > 0 Then
            data_em.Recordset.MoveFirst
            Do While Not data_em.Recordset.EOF
               If Check2.Value = 1 Then
                  If data_em.Recordset("fecha") >= CDate(mf.Text) Then
                    data_inf.Recordset.AddNew
                    data_inf.Recordset("cod_cnv") = data_em.Recordset("cod_cnv")
                    data_inf.Recordset("nom_cnv") = data_em.Recordset("nom_cnv")
                    data_inf.Recordset("cliente") = data_em.Recordset("cliente")
                    data_inf.Recordset("apellidos") = data_em.Recordset("apellidos")
                    data_inf.Recordset("documento") = data_em.Recordset("documento")
                    data_inf.Recordset("fecha") = data_em.Recordset("fecha")
                    data_inf.Recordset("importe") = data_em.Recordset("importe")
                    data_inf.Recordset("nro_cobr") = data_em.Recordset("nro_cobr")
                    data_inf.Recordset("nom_cobr") = data_em.Recordset("nom_cobr")
                    data_inf.Recordset("grupo") = data_em.Recordset("grupo")
                    data_inf.Recordset("zona") = data_em.Recordset("zona")
                    data_inf.Recordset("mes") = data_em.Recordset("mes")
                    data_inf.Recordset("ano") = data_em.Recordset("ano")
                    data_inf.Recordset("fecha_ing") = data_em.Recordset("fecha_ing")
                    data_inf.Recordset("origen") = data_em.Recordset("tel_cli")
                    data_inf.Recordset("color_rec") = data_em.Recordset("color_rec")
                    data_inf.Recordset("tiquet") = data_em.Recordset("tiquet")
                    If Format(mf.Text, "yyyy/mm/dd") >= Format("01/11/2020", "yyyy/mm/dd") Then
                       data_inf.Recordset("servi") = data_em.Recordset("descimp")
                    End If
                    data_inf.Recordset("deudas") = data_em.Recordset("deudas")
'                    t_iva.Text = Format(data_em.Recordset("iva"), "#.##")
                    data_inf.Recordset("iva") = data_em.Recordset("iva")
                    data_inf.Recordset("total") = data_em.Recordset("total")
                    data_inf.Recordset("ruc") = data_em.Recordset("ruc")
                    data_inf.Recordset("ap") = data_em.Recordset("deudaap")
                    data_inf.Recordset.Update
                  End If
               Else
                  If data_em.Recordset("fecha") <= CDate(mf.Text) Then
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cod_cnv") = data_em.Recordset("cod_cnv")
                        data_inf.Recordset("nom_cnv") = data_em.Recordset("nom_cnv")
                        data_inf.Recordset("cliente") = data_em.Recordset("cliente")
                        data_inf.Recordset("apellidos") = data_em.Recordset("apellidos")
                        data_inf.Recordset("documento") = data_em.Recordset("documento")
                        data_inf.Recordset("fecha") = data_em.Recordset("fecha")
                        data_inf.Recordset("importe") = data_em.Recordset("importe")
                        data_inf.Recordset("nro_cobr") = data_em.Recordset("nro_cobr")
                        data_inf.Recordset("nom_cobr") = data_em.Recordset("nom_cobr")
                        data_inf.Recordset("grupo") = data_em.Recordset("grupo")
                        data_inf.Recordset("zona") = data_em.Recordset("zona")
                        data_inf.Recordset("mes") = data_em.Recordset("mes")
                        data_inf.Recordset("fecha_ing") = data_em.Recordset("fecha_ing")
                        data_inf.Recordset("origen") = data_em.Recordset("tel_cli")
                        data_inf.Recordset("ano") = data_em.Recordset("ano")
                        data_inf.Recordset("color_rec") = data_em.Recordset("color_rec")
                        data_inf.Recordset("tiquet") = data_em.Recordset("tiquet")
                        If Format(mf.Text, "yyyy/mm/dd") >= Format("01/11/2020", "yyyy/mm/dd") Then
                           data_inf.Recordset("servi") = data_em.Recordset("descimp")
                        End If
                        data_inf.Recordset("deudas") = data_em.Recordset("deudas")
                        data_inf.Recordset("iva") = data_em.Recordset("iva")
                        data_inf.Recordset("total") = data_em.Recordset("total")
                        data_inf.Recordset("ruc") = data_em.Recordset("ruc")
                        data_inf.Recordset("ap") = data_em.Recordset("deudaap")
                        data_inf.Recordset.Update
                  End If
               End If
               data_em.Recordset.MoveNext
            Loop
            data_inf.RecordSource = "Select * from infemis order by nro_cobr,color_rec,documento,importe"
            data_inf.Refresh
            If Combo2.ListIndex = 0 Then
               cremis.ReportTitle = "EMISION CORRESPONDIENTE A...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
               If Check3.Value = 1 Then
                  cremis.ReportFileName = App.path & "\infemidruc.rpt"
               Else
                  cremis.ReportFileName = App.path & "\infemin.rpt"
               End If
               cremis.DiscardSavedData = True
               cremis.Action = 1
            Else
               If Combo2.ListIndex = 1 Then
                  cremis.ReportTitle = "EMISION CORRESPONDIENTE A...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
                  If Check3.Value = 1 Then
                     cremis.ReportFileName = App.path & "\infemidruc.rpt"
                  Else
                     If Quecob <> "" Then
                        cremis.ReportFileName = App.path & "\infemidcob.rpt"
                     Else
                        cremis.ReportFileName = App.path & "\infemid.rpt"
                     End If
                  End If
                  cremis.DiscardSavedData = True
                  cremis.Action = 1
               Else
                  If Combo2.ListIndex = 3 Then
                     cremis.ReportTitle = "EMISION CORRESPONDIENTE A...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
                     cremis.ReportFileName = App.path & "\infeminsc.rpt"
                     cremis.DiscardSavedData = True
                     cremis.Action = 1
                  Else
                     If Combo2.ListIndex = 4 Then
                        cremis.ReportTitle = "EMISION CORRESPONDIENTE A...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
                        cremis.ReportFileName = App.path & "\infemid.rpt"
                        cremis.DiscardSavedData = True
                        cremis.Action = 1
                     Else
                        If Combo2.ListIndex = 5 Then
                           cremis.ReportTitle = "EMISION POR CATEGORÍA CORRESPONDIENTE A...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
                           cremis.ReportFileName = App.path & "\infemicat.rpt"
                           cremis.DiscardSavedData = True
                           cremis.Action = 1
                        Else
                           cremis.ReportTitle = "EMISION CORRESPONDIENTE A...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
                           cremis.ReportFileName = App.path & "\infemsj.rpt"
                           cremis.DiscardSavedData = True
                           cremis.Action = 1
                        End If
                     End If
                  End If
               End If
            End If
         Else
            MsgBox "No existen registros", vbInformation, "Mensaje"
         End If
      End If
      
      If Option2.Value = True Then 'SIMULACION
         Xlugar = App.path & "\simulan.mdb"
         data_em.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & Xlugar
         data_em.RecordSource = "select * from emisim where mes_emi =" & txt_m.Text & " and anio_emi=" & txt_a.Text
         data_em.Refresh
         If Combo1.ListIndex = 0 Then
            If Check1.Value = 1 Then
               data_em.RecordSource = "Select * from emisim where mes_emi =" & txt_m.Text & " and anio_emi =" & txt_a.Text & " and nro_cobr <>" & 616 & " And nro_cobr <>" & 603 & _
               " And nro_cobr <>" & 636 & " And nro_cobr<>" & 615 & " And nro_cobr<>" & 635 & " And nro_cobr<>" & 201 & " And nro_cobr<>" & 208 & _
               " and nro_cobr <>" & 602 & " And nro_cobr<>" & 653 & " And nro_cobr<>" & 672 & " And nro_cobr<>" & 679 & " And nro_cobr<>" & 112 & " And nro_cobr<>" & 209 & _
               " and nro_cobr <>" & 8 & " And nro_cobr<>" & 685 & " And nro_cobr<>" & 512 & " And nro_cobr <>" & 606 & _
               " and nro_cobr <>" & 113 & " And nro_cobr<>" & 1 & " and nro_cobr<>" & 10 & " order by nro_cobr,color_rec,documento,importe"
               data_em.Refresh
            Else
                data_em.RecordSource = "Select * from emisim where mes_emi =" & txt_m.Text & " and anio_emi=" & txt_a.Text & " order by nro_cobr,color_rec,importe"
                data_em.Refresh
            End If
         Else
            Quecob = InputBox("Ingrese número de cobrador:", "Solicitud de datos")
            If Quecob <> "" Then
               data_em.RecordSource = "Select * from emisim where mes_emi =" & txt_m.Text & " and anio_emi=" & txt_a.Text & " and nro_cobr =" & Quecob & " order by color_rec,importe"
               data_em.Refresh
            Else
               data_em.RecordSource = "Select * from emisim where mes_emi =" & txt_m.Text & " and anio_emi=" & txt_a.Text & " order by nro_cobr,color_rec,importe"
               data_em.Refresh
            End If
         End If
         If data_em.Recordset.RecordCount > 0 Then
            data_em.Recordset.MoveFirst
            Do While Not data_em.Recordset.EOF
               data_inf.Recordset.AddNew
               data_inf.Recordset("cod_cnv") = data_em.Recordset("cod_cnv")
               data_inf.Recordset("nom_cnv") = data_em.Recordset("nom_cnv")
               data_inf.Recordset("cliente") = data_em.Recordset("cliente")
               data_inf.Recordset("apellidos") = data_em.Recordset("apellidos")
               data_inf.Recordset("documento") = data_em.Recordset("documento")
               data_inf.Recordset("importe") = data_em.Recordset("importe")
               data_inf.Recordset("nro_cobr") = data_em.Recordset("nro_cobr")
               data_inf.Recordset("nom_cobr") = data_em.Recordset("nom_cobr")
               data_inf.Recordset("grupo") = data_em.Recordset("grupo")
               data_inf.Recordset("zona") = data_em.Recordset("zona")
               data_inf.Recordset("mes") = data_em.Recordset("mes")
               data_inf.Recordset("ano") = data_em.Recordset("ano")
               data_inf.Recordset("fecha_ing") = data_em.Recordset("fecha_ing")
               data_inf.Recordset("color_rec") = data_em.Recordset("color_rec")
               data_inf.Recordset("tiquet") = data_em.Recordset("tiquet")
               If Format(mf.Text, "yyyy/mm/dd") >= Format("01/11/2020", "yyyy/mm/dd") Then
                  data_inf.Recordset("servi") = data_em.Recordset("descimp")
               End If
               data_inf.Recordset("deudas") = data_em.Recordset("deudas")
               data_inf.Recordset("iva") = data_em.Recordset("iva")
               data_inf.Recordset("total") = data_em.Recordset("total")
               data_inf.Recordset("ruc") = data_em.Recordset("ruc")
               data_inf.Recordset("ap") = data_em.Recordset("ap")
               data_inf.Recordset.Update
               data_em.Recordset.MoveNext
            Loop
            data_inf.RecordSource = "select * from infemis order by nro_cobr,color_rec,importe"
            data_inf.Refresh
            If Combo2.ListIndex = 0 Then
               cremis.ReportTitle = "SIMULACION CORRESPONDIENTE A...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
               cremis.ReportFileName = App.path & "\infsimulan.rpt"
               cremis.Action = 1
            Else
               cremis.ReportTitle = "SIMULACION CORRESPONDIENTE A...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
               cremis.ReportFileName = App.path & "\infsimula.rpt"
               cremis.Action = 1
            End If
         Else
            MsgBox "No existen registros", vbInformation, "Mensaje"
         End If
      End If
   End If
Else
   MsgBox "Verifique si ingresó fecha de nuevas entregas", vbInformation
End If
'If Option4.value = True Then
'   cremis.ReportTitle = "CONTROL DE NUMERACION DE RECIBOS DE EMISION...:" & Trim(txt_m.Text) & "/" & Trim(txt_a.Text)
'   cremis.ReportFileName = App.Path & "\infnumemi.rpt"
'   cremis.Action = 1
'End If
frm_infemis.MousePointer = 0

'Unload Me

Exit Sub

Noesta:
       If Err.Number = 3078 Then
          frm_infemis.MousePointer = 0
          MsgBox "No existe emisión, VERIFIQUE!!", vbCritical, "Mensaje"
          txt_m.SetFocus
       Else
          frm_infemis.MousePointer = 0
          MsgBox "Hay un error, VERIFIQUE si ingresó fecha de nuevas entregas!!", vbCritical, "Mensaje"
          txt_m.SetFocus
       End If

End Sub

Private Sub b_can_Click()
Unload Me

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Dim Xlugar As String
Xlugar = App.path & "\informes.mdb"
'If frm_menu.data_parse.Recordset("base") = 10 Then
'   data_em.DatabaseName = App.Path & "\emisiones.mdb"
'Else
'data_em.DatabaseName = ""
'data_em.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_em.RecordSource = "Select * from emi0217 where nro_cobr =" & 605
'data_em.Refresh

'data_inf.DatabaseName = App.Path & "\informes.mdb"
'data_inf.RecordSource = "infemis"
'data_inf.Refresh
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & Xlugar
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infemis"
data_inf.Refresh
data_em.ConnectionString = "dsn=" & Xconexrmt
txt_m.Text = Month(Date)
txt_a.Text = Year(Date)
adocnv.ConnectionString = "dsn=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option3.SetFocus
End If

End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub txt_a_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option1.SetFocus
End If

End Sub

Private Sub txt_m_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_a.SetFocus
End If

End Sub
