VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infsolhc 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de solicitudes de copia de HC"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6615
   Icon            =   "frm_infsolhc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport CR1 
      Left            =   5880
      Top             =   2880
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
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Data data_hc 
      Caption         =   "data_hc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Picture         =   "frm_infsolhc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "frm_infsolhc.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para el informe."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6255
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   480
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "data_lin"
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Detalle"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   1680
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Resumen"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox t_base 
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
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Base (99=TODAS)"
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Rango de fechas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   2520
      Picture         =   "frm_infsolhc.frx":0F56
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "frm_infsolhc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If t_base.Text = "" Then
   t_base.Text = 99
End If
frm_infsolhc.MousePointer = 11
Command1.Enabled = False
Command2.Enabled = False
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "infvtas"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If

If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      If t_base.Text = 99 Then
         data_lin.RecordSource = "Select * from linmmdd where cod_prod =" & 991 & " And fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by fecha"
         data_lin.Refresh
         data_hc.Connect = "odbc;dsn=" & Xconexrmt & ";"
         data_hc.RecordSource = "Select * from provdeu where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "#"
         data_hc.Refresh
      Else
         data_lin.RecordSource = "Select * from linmmdd where cod_prod =" & 991 & " And fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & t_base.Text & " order by fecha"
         data_lin.Refresh
         data_hc.Connect = "odbc;dsn=" & Xconexrmt & ";"
         data_hc.RecordSource = "Select * from provdeu where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "#"
         data_hc.Refresh
      End If
      If data_lin.Recordset.RecordCount > 0 Then
         data_lin.Recordset.MoveFirst
         Do While Not data_lin.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
            data_inf.Recordset("factura") = data_lin.Recordset("factura")
            data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
            data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
            data_inf.Recordset("ced_socio") = data_lin.Recordset("ced_socio")
            data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
            data_hc.Recordset.FindFirst "documento =" & data_lin.Recordset("factura")
'            data_hc.RecordSource = "Select * from provdeu where documento =" & data_lin.Recordset("factura")
'            data_hc.Refresh
            If Not data_hc.Recordset.NoMatch Then
               data_inf.Recordset("nom_superv") = Mid(data_hc.Recordset("nom_cnv"), 1, 25)
               data_inf.Recordset("reg_cab") = data_hc.Recordset("cliente")
               If IsNull(data_hc.Recordset("origen")) = False Then
                  data_inf.Recordset("nom_flia") = Mid(data_hc.Recordset("origen"), 1, 40)
               End If
               If IsNull(data_hc.Recordset("mes")) = False Then
                  If data_hc.Recordset("mes") = 0 Then
                     data_inf.Recordset("tipo") = "TODOS"
                  End If
                  If data_hc.Recordset("mes") = 1 Then
                     data_inf.Recordset("tipo") = "POLICLINICA"
                  End If
                  If data_hc.Recordset("mes") = 2 Then
                     data_inf.Recordset("tipo") = "ESPECIALISTAS"
                  End If
                  If data_hc.Recordset("mes") = 3 Then
                     data_inf.Recordset("tipo") = "TRASLADOS"
                  End If
                  If data_hc.Recordset("mes") = 4 Then
                     data_inf.Recordset("tipo") = "DOMICILIO"
                  End If
               Else
                  data_inf.Recordset("tipo") = "TODOS"
               End If
               If IsNull(data_hc.Recordset("moneda")) = False Then
                  If data_hc.Recordset("moneda") = 0 Then
                     data_inf.Recordset("ruc") = "SECC.POLICIAL"
                  End If
                  If data_hc.Recordset("moneda") = 1 Then
                     data_inf.Recordset("ruc") = "JUZGADO"
                  End If
                  If data_hc.Recordset("moneda") = 2 Then
                     data_inf.Recordset("ruc") = "USUARIO"
                  End If
                  If data_hc.Recordset("moneda") = 3 Then
                     data_inf.Recordset("ruc") = "MADRE"
                  End If
                  If data_hc.Recordset("moneda") = 4 Then
                     data_inf.Recordset("ruc") = "PADRE"
                  End If
                  If data_hc.Recordset("moneda") = 5 Then
                     data_inf.Recordset("ruc") = "HIJO"
                  End If
                  If data_hc.Recordset("moneda") = 6 Then
                     data_inf.Recordset("ruc") = "TUTOR"
                  End If
                  If data_hc.Recordset("moneda") = 7 Then
                     data_inf.Recordset("ruc") = "ABOGADO"
                  End If
                  If data_hc.Recordset("moneda") = 8 Then
                     data_inf.Recordset("ruc") = "OTRO"
                  End If
               
               Else
                  data_inf.Recordset("ruc") = "OTRO"
               End If
               data_inf.Recordset("zona") = data_hc.Recordset("nom_cobr")
               If IsNull(data_hc.Recordset("nombre")) = False Then
                  data_inf.Recordset("nom_medic") = Mid(data_hc.Recordset("nombre"), 1, 50)
               End If
               data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
               data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
               data_inf.Recordset("operador") = data_lin.Recordset("operador")
               data_inf.Recordset("base") = data_lin.Recordset("base")
               data_inf.Recordset.Update
               data_lin.Recordset.MoveNext
            Else
               data_lin.Recordset.MoveNext
               data_inf.Recordset("nom_prod") = "SIN DATOS"
               data_inf.Recordset.Update
            End If
         Loop
         MsgBox "Proceso terminado"
         data_inf.RecordSource = "Select * from infvtas"
         data_inf.Refresh
         If Option1.value = True Then
            cr1.ReportFileName = App.Path & "\infsolhcd.rpt"
            cr1.Action = 1
         End If
         If Option2.value = True Then
            cr1.ReportFileName = App.Path & "\infsolhcf.rpt"
            cr1.Action = 1
         End If
      Else
         MsgBox "No hay registros", vbInformation, "Mensaje"
      End If
   Else
      MsgBox "Ingrese fecha"
   End If
Else
   MsgBox "Ingrese Fecha"
End If
frm_infsolhc.MousePointer = 0
Command1.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

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
   t_base.SetFocus
End If

End Sub

Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
