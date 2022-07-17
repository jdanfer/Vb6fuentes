VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_impemi 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir recibos de emisión"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "frm_impemi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Crystal.CrystalReport cremi 
      Left            =   5160
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_informe 
      Caption         =   "data_informe"
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
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
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
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton b_canc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      Picture         =   "frm_impemi.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton b_acep 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   480
      Picture         =   "frm_impemi.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para imprimir recibos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Sólo Pago anual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   495
         Left            =   3240
         TabIndex        =   19
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox t_radhasta 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox t_raddesde 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   3120
         Width           =   735
      End
      Begin VB.Data infor_det 
         Caption         =   "infor_det"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Data data_cabemi 
         Caption         =   "data_cabemi"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "San Jacinto"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox txt_cob 
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
         Height          =   435
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt_hasta 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txt_desde 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   2280
         Width           =   1815
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
         ItemData        =   "frm_impemi.frx":0CC6
         Left            =   2040
         List            =   "frm_impemi.frx":0CD9
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   1815
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
         Left            =   2640
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
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800000&
         Caption         =   "Rango de Zonas:"
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
         TabIndex        =   16
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Número RECIBO:"
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
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Color:"
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
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
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
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
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
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   0
      Picture         =   "frm_impemi.frx":0D01
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "frm_impemi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
Dim Xquecolor As String
Dim Nombre, NombreCab, Dbnombre As String
Text1.Text = 5

Dbnombre = "DB"
Nombre = "EMI"
NombreCab = "CAB"
If txt_m.Text > 9 Then
   Nombre = Nombre + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
   Dbnombre = Dbnombre + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
   NombreCab = NombreCab + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
Else
   Nombre = Nombre + "0" + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
   Dbnombre = Dbnombre + "0" + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
   NombreCab = NombreCab + "0" + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
End If
If Nombre = "EMI0716" Or Nombre = "EMI0816" Then
   If frm_menu.data_parse.Recordset("base") = 3 Then
      data_emi.DatabaseName = App.path & "\emisnueva"
      data_cabemi.DatabaseName = App.path & "\emisnueva.mdb"
   Else
      data_emi.DatabaseName = App.path & "\emis7y8.mdb"
      data_cabemi.DatabaseName = App.path & "\emis7y8.mdb"
   End If
Else
   If Nombre = "EMI0616" Or Nombre = "EMI0516" Then
      data_emi.Connect = "ODBC;DSN=sapp;"
      data_cabemi.DatabaseName = App.path & "\emitemp.mdb"
      data_cabemi.RecordSource = "Select * from cabemi"
      data_cabemi.Refresh
      If data_cabemi.Recordset.RecordCount > 0 Then
         data_cabemi.Recordset.MoveFirst
         Do While Not data_cabemi.Recordset.EOF
            data_cabemi.Recordset.Delete
            data_cabemi.Recordset.MoveNext
         Loop
         data_cabemi.Refresh
      End If
   Else
      If Nombre = "EMI0916" Or Nombre = "EMI1016" Then
         data_emi.DatabaseName = App.path & "\emis9y10.mdb"
         data_cabemi.DatabaseName = App.path & "\emis9y10.mdb"
      Else
         If Nombre = "EMI1116" Or Nombre = "EMI1216" Then
            data_emi.DatabaseName = App.path & "\emis1112.mdb"
            data_cabemi.DatabaseName = App.path & "\emis1112.mdb"
         Else
            data_emi.DatabaseName = App.path & "\" & Dbnombre & ".mdb"
            data_cabemi.DatabaseName = App.path & "\" & Dbnombre & ".mdb"
         End If
      End If
   End If
End If

If Combo1.ListIndex = 0 Then
   Xquecolor = "R"
End If
If Combo1.ListIndex = 1 Then
   Xquecolor = "A"
End If
If Combo1.ListIndex = 2 Then
   Xquecolor = "M"
End If
If Combo1.ListIndex = 3 Then
   Xquecolor = "V"
End If
If Combo1.ListIndex = 4 Then
   Xquecolor = "C"
End If
Dim Xcantlineas As Integer

frm_impemi.MousePointer = 11
If txt_m.Text <> "" Then
   If txt_a.Text <> "" Then
      If txt_desde.Text <> "" Then
         If txt_hasta.Text <> "" Then
            If t_raddesde.Text <> "" And t_radhasta.Text <> "" Then
               data_emi.RecordSource = "Select * from " & Nombre & " where documento >=" & txt_desde.Text & " And documento <=" & txt_hasta.Text & " and grupo >=" & Val(t_raddesde.Text) & " and grupo <=" & Val(t_radhasta.Text) & " order by nro_cobr,color_rec,documento,importe"
               data_emi.Refresh
            Else
               data_emi.RecordSource = "Select * from " & Nombre & " where documento >=" & txt_desde.Text & " And documento <=" & txt_hasta.Text & " order by nro_cobr,color_rec,documento,importe"
               data_emi.Refresh
            End If
            data_informe.RecordSource = "EMIS"
            data_informe.Refresh
            If data_informe.Recordset.RecordCount > 0 Then
               data_informe.Recordset.MoveFirst
               Do While Not data_informe.Recordset.EOF
                  data_informe.Recordset.Delete
                  data_informe.Recordset.MoveNext
               Loop
               data_informe.Refresh
            End If
            infor_det.RecordSource = "CABEZAL"
            infor_det.Refresh
            If infor_det.Recordset.RecordCount > 0 Then
               infor_det.Recordset.MoveFirst
               Do While Not infor_det.Recordset.EOF
                  infor_det.Recordset.Delete
                  infor_det.Recordset.MoveNext
               Loop
               infor_det.Refresh
            End If
            If data_emi.Recordset.RecordCount > 0 Then
               data_emi.Recordset.MoveFirst
               Do While Not data_emi.Recordset.EOF
                  If data_emi.Recordset("nro_cobr") = 10 Then
                     data_emi.Recordset.MoveNext
                  Else
                    data_informe.Recordset.AddNew
                    data_informe.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
                    data_informe.Recordset("nom_cnv") = data_emi.Recordset("nom_cnv")
                    data_informe.Recordset("cliente") = data_emi.Recordset("cliente")
                    data_informe.Recordset("apellidos") = data_emi.Recordset("apellidos")
                    data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                    data_informe.Recordset("cedula") = data_emi.Recordset("cedula")
                    data_informe.Recordset("dir_cli") = data_emi.Recordset("dir_cli")
                    data_informe.Recordset("loc_cli") = data_emi.Recordset("loc_cli")
                    data_informe.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
                    data_informe.Recordset("grupo") = data_emi.Recordset("grupo")
                    data_informe.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
                    data_informe.Recordset("documento") = data_emi.Recordset("documento")
                    data_informe.Recordset("importe") = data_emi.Recordset("importe")
                    data_informe.Recordset("promos") = data_emi.Recordset("promo")
                    If data_emi.Recordset("debe_haber") = 101 Then
                       data_informe.Recordset("tipofact") = "e-Ticket"
                    Else
                       data_informe.Recordset("tipofact") = "e-Factura"
                    End If
                    data_informe.Recordset("fpago") = "CREDITO"
                    data_informe.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
                    data_informe.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
                    data_informe.Recordset("mes") = data_emi.Recordset("mes")
                    data_informe.Recordset("ano") = data_emi.Recordset("ano")
                    data_informe.Recordset("color_rec") = data_emi.Recordset("color_rec")
                    data_informe.Recordset("tiquet") = data_emi.Recordset("tiquet")
                    data_informe.Recordset("servi") = data_emi.Recordset("servi")
                    data_informe.Recordset("deudas") = data_emi.Recordset("deudas")
                    data_informe.Recordset("iva") = data_emi.Recordset("iva")
                    data_informe.Recordset("total") = data_emi.Recordset("total")
                    data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                    data_informe.Recordset("zona") = data_emi.Recordset("zona")
                    data_informe.Recordset("qr") = data_emi.Recordset("qr")
                    data_informe.Recordset("fvence") = data_emi.Recordset("fvence")
                    data_informe.Recordset("autoriza") = data_emi.Recordset("autoriza")
                    data_informe.Recordset("rangoCAE") = data_emi.Recordset("rangoCAE")
                    data_informe.Recordset("codseg") = data_emi.Recordset("codseg")
                    If IsNull(data_emi.Recordset("ruc")) = False Then
                       data_informe.Recordset("hora") = Null
                    Else
                       data_informe.Recordset("hora") = "X"
                    End If
                    data_informe.Recordset("fecha") = data_emi.Recordset("fecha")
                    data_informe.Recordset("fecha_cobr") = data_emi.Recordset("fecha_cobr")
                    data_informe.Recordset("nom_superv") = "*" & Trim(str(data_emi.Recordset("cliente"))) & "*"
                    data_informe.Recordset("barranrodoc") = "*" & Trim(str(data_emi.Recordset("documento"))) & "*"
                    data_informe.Recordset("tipocta") = data_emi.Recordset("tipocta")
                    data_informe.Recordset("tipodoc") = data_emi.Recordset("tipodoc")
                    data_informe.Recordset.Update
                    If Nombre = "EMI0516" Or Nombre = "EMI0616" Then
                       data_cabemi.Recordset.AddNew
                       data_cabemi.Recordset("serie") = "A"
                       data_cabemi.Recordset("nro_doc") = data_emi.Recordset("documento")
                       data_cabemi.Recordset("cod_srv") = 881
                       data_cabemi.Recordset("descrip") = "CUOTA MENSUAL " & Trim(str(data_emi.Recordset("mes"))) & "/" & Trim(str(data_emi.Recordset("ano")))
                       data_cabemi.Recordset("imp_srv") = data_emi.Recordset("importe")
                       data_cabemi.Recordset("nro_linea") = 1
                       data_cabemi.Recordset("cantidad") = 1
                       data_cabemi.Recordset("monto") = data_emi.Recordset("total")
                       data_cabemi.Recordset("cliente2") = data_emi.Recordset("cliente")
                       data_cabemi.Recordset.Update
                       data_cabemi.Refresh
                    Else
                       data_cabemi.RecordSource = "Select * from " & NombreCab & " where nro_doc =" & data_emi.Recordset("documento")
                       data_cabemi.Refresh
                    End If
                    If data_cabemi.Recordset.RecordCount > 0 Then
                       data_cabemi.Recordset.MoveFirst
                       Do While Not data_cabemi.Recordset.EOF
                          Xcantlineas = Xcantlineas + 1
                          infor_det.Recordset.AddNew
                          infor_det.Recordset("serie") = data_cabemi.Recordset("serie")
                          infor_det.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
                          infor_det.Recordset("cod_srv") = data_cabemi.Recordset("cod_srv")
                          If data_cabemi.Recordset("cod_srv") = 881 Then
                             infor_det.Recordset("descrip") = data_cabemi.Recordset("descrip") & "  " & Trim(str(data_emi.Recordset("mes"))) & "/" & Trim(str(data_emi.Recordset("ano")))
                          Else
                             infor_det.Recordset("descrip") = data_cabemi.Recordset("descrip")
                          End If
                          infor_det.Recordset("imp_srv") = data_cabemi.Recordset("imp_srv")
                          infor_det.Recordset("nro_linea") = data_cabemi.Recordset("nro_linea")
                          infor_det.Recordset("cantidad") = data_cabemi.Recordset("cantidad")
                          infor_det.Recordset("monto") = data_cabemi.Recordset("monto")
                          infor_det.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
                          infor_det.Recordset.Update
                          data_cabemi.Recordset.MoveNext
                       Loop
                       If Xcantlineas < 3 Then
                          data_cabemi.Recordset.MovePrevious
                          Do While Xcantlineas <= 3
                             Xcantlineas = Xcantlineas + 1
                             infor_det.Recordset.AddNew
                             infor_det.Recordset("serie") = data_cabemi.Recordset("serie")
                             infor_det.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
                             infor_det.Recordset("cod_srv") = Null
                             infor_det.Recordset("descrip") = Null
                             infor_det.Recordset("imp_srv") = Null
                             infor_det.Recordset("nro_linea") = Xcantlineas
                             infor_det.Recordset("cantidad") = Null
                             infor_det.Recordset("monto") = Null
                             infor_det.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
                             infor_det.Recordset.Update
                          Loop
                       End If
                       Xcantlineas = 0
                    End If
                    data_emi.Recordset.MoveNext
                  End If
               Loop
            Else
               frm_impemi.MousePointer = 0
               MsgBox "No existen registros", vbInformation, "Mensaje"
            End If
         End If
      Else
         If txt_cob.Text = "" Then
            If Check1.Value = 1 Then
               data_emi.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & 11 & " or nro_cobr =" & 6 & " or nro_cobr =" & 5 & " order by nro_cobr,color_rec,importe"
               data_emi.Refresh
            Else
'acá
               data_emi.RecordSource = "Select * from " & Nombre & " where color_rec ='" & Trim(Xquecolor) & "' order by nro_cobr,color_rec,documento,importe"
               data_emi.Refresh
            End If
            data_informe.RecordSource = "EMIS"
            data_informe.Refresh
            If data_informe.Recordset.RecordCount > 0 Then
               data_informe.Recordset.MoveFirst
               Do While Not data_informe.Recordset.EOF
                  data_informe.Recordset.Delete
                  data_informe.Recordset.MoveNext
               Loop
               data_informe.Refresh
            End If
            infor_det.RecordSource = "CABEZAL"
            infor_det.Refresh
            If infor_det.Recordset.RecordCount > 0 Then
               infor_det.Recordset.MoveFirst
               Do While Not infor_det.Recordset.EOF
                  infor_det.Recordset.Delete
                  infor_det.Recordset.MoveNext
               Loop
               infor_det.Refresh
            End If
            If data_emi.Recordset.RecordCount > 0 Then
               data_emi.Recordset.MoveFirst
               Do While Not data_emi.Recordset.EOF
                  If data_emi.Recordset("nro_cobr") = 616 Or _
                     data_emi.Recordset("nro_cobr") = 636 Or _
                     data_emi.Recordset("nro_cobr") = 615 Or _
                     data_emi.Recordset("nro_cobr") = 635 Or _
                     data_emi.Recordset("nro_cobr") = 602 Or _
                     data_emi.Recordset("nro_cobr") = 653 Or _
                     data_emi.Recordset("nro_cobr") = 672 Or _
                     data_emi.Recordset("nro_cobr") = 113 Or _
                     data_emi.Recordset("nro_cobr") = 1 Or _
                     data_emi.Recordset("nro_cobr") = 10 Or _
                     data_emi.Recordset("nro_cobr") = 8 Or _
                     data_emi.Recordset("nro_cobr") = 603 Or _
                     data_emi.Recordset("nro_cobr") = 685 Or _
                     data_emi.Recordset("nro_cobr") = 201 Or _
                     data_emi.Recordset("nro_cobr") = 604 Or _
                     data_emi.Recordset("nro_cobr") = 606 Or _
                     data_emi.Recordset("nro_cobr") = 676 Or _
                     data_emi.Recordset("nro_cobr") = 688 Or _
                     data_emi.Recordset("nro_cobr") = 512 Or _
                     data_emi.Recordset("nro_cobr") = 679 Then
                     data_emi.Recordset.MoveNext
                  Else
                    If Check2.Value = 1 Then
                       If data_emi.Recordset("promo") = "Pago anual" Then
                            data_informe.Recordset.AddNew
                            data_informe.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
                            data_informe.Recordset("nom_cnv") = data_emi.Recordset("nom_cnv")
                            data_informe.Recordset("cliente") = data_emi.Recordset("cliente")
                            data_informe.Recordset("apellidos") = data_emi.Recordset("apellidos")
                            data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                            data_informe.Recordset("cedula") = data_emi.Recordset("cedula")
                            data_informe.Recordset("dir_cli") = data_emi.Recordset("dir_cli")
                            data_informe.Recordset("loc_cli") = data_emi.Recordset("loc_cli")
                            data_informe.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
                            data_informe.Recordset("grupo") = data_emi.Recordset("grupo")
                            data_informe.Recordset("promos") = data_emi.Recordset("promo")
                            If data_emi.Recordset("debe_haber") = 101 Then
                               data_informe.Recordset("tipofact") = "e-Ticket"
                            Else
                               data_informe.Recordset("tipofact") = "e-Factura"
                            End If
                            data_informe.Recordset("fpago") = "CREDITO"
                            data_informe.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
                            data_informe.Recordset("documento") = data_emi.Recordset("documento")
                            data_informe.Recordset("importe") = data_emi.Recordset("importe")
                            data_informe.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
                            data_informe.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
                            data_informe.Recordset("mes") = data_emi.Recordset("mes")
                            data_informe.Recordset("ano") = data_emi.Recordset("ano")
                            data_informe.Recordset("color_rec") = data_emi.Recordset("color_rec")
                            data_informe.Recordset("tiquet") = data_emi.Recordset("tiquet")
                            data_informe.Recordset("servi") = data_emi.Recordset("servi")
                            data_informe.Recordset("deudas") = data_emi.Recordset("deudas")
                            data_informe.Recordset("iva") = data_emi.Recordset("iva")
                            data_informe.Recordset("total") = data_emi.Recordset("total")
                            data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                            data_informe.Recordset("zona") = data_emi.Recordset("zona")
                            data_informe.Recordset("qr") = data_emi.Recordset("qr")
                            data_informe.Recordset("fvence") = data_emi.Recordset("fvence")
                            data_informe.Recordset("autoriza") = data_emi.Recordset("autoriza")
                            data_informe.Recordset("rangoCAE") = data_emi.Recordset("rangoCAE")
                            data_informe.Recordset("codseg") = data_emi.Recordset("codseg")
                            If IsNull(data_emi.Recordset("ruc")) = False Then
                               data_informe.Recordset("hora") = Null
                            Else
                               data_informe.Recordset("hora") = "X"
                            End If
                            data_informe.Recordset("fecha") = data_emi.Recordset("fecha")
                            data_informe.Recordset("fecha_cobr") = data_emi.Recordset("fecha_cobr")
                            data_informe.Recordset("nom_superv") = "*" & Trim(str(data_emi.Recordset("cliente"))) & "*"
                            data_informe.Recordset("barranrodoc") = "*" & Trim(str(data_emi.Recordset("documento"))) & "*"
                            data_informe.Recordset("tipocta") = data_emi.Recordset("tipocta")
                            data_informe.Recordset("tipodoc") = data_emi.Recordset("tipodoc")
                            data_informe.Recordset.Update
                            If Nombre = "EMI0516" Or Nombre = "EMI0616" Then
                               data_cabemi.Recordset.AddNew
                               data_cabemi.Recordset("serie") = "A"
                               data_cabemi.Recordset("nro_doc") = data_emi.Recordset("documento")
                               data_cabemi.Recordset("cod_srv") = 881
                               data_cabemi.Recordset("descrip") = "CUOTA MENSUAL " & Trim(str(data_emi.Recordset("mes"))) & "/" & Trim(str(data_emi.Recordset("ano")))
                               data_cabemi.Recordset("imp_srv") = data_emi.Recordset("importe")
                               data_cabemi.Recordset("nro_linea") = 1
                               data_cabemi.Recordset("cantidad") = 1
                               data_cabemi.Recordset("monto") = data_emi.Recordset("total")
                               data_cabemi.Recordset("cliente2") = data_emi.Recordset("cliente")
                               data_cabemi.Recordset.Update
                               data_cabemi.Refresh
                            Else
                               data_cabemi.RecordSource = "Select * from " & NombreCab & " where nro_doc =" & data_emi.Recordset("documento")
                               data_cabemi.Refresh
                            End If
                            Xcantlineas = 0
                            If data_cabemi.Recordset.RecordCount > 0 Then
                               data_cabemi.Recordset.MoveFirst
                               Do While Not data_cabemi.Recordset.EOF
                                  Xcantlineas = Xcantlineas + 1
                                  infor_det.Recordset.AddNew
                                  infor_det.Recordset("serie") = data_cabemi.Recordset("serie")
                                  infor_det.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
                                  infor_det.Recordset("cod_srv") = data_cabemi.Recordset("cod_srv")
                                  If data_cabemi.Recordset("cod_srv") = 881 Then
                                     infor_det.Recordset("descrip") = data_cabemi.Recordset("descrip") & "  " & Trim(str(data_emi.Recordset("mes"))) & "/" & Trim(str(data_emi.Recordset("ano")))
                                  Else
                                     infor_det.Recordset("descrip") = data_cabemi.Recordset("descrip")
                                  End If
                                  infor_det.Recordset("imp_srv") = data_cabemi.Recordset("imp_srv")
                                  infor_det.Recordset("nro_linea") = data_cabemi.Recordset("nro_linea")
                                  infor_det.Recordset("cantidad") = data_cabemi.Recordset("cantidad")
                                  infor_det.Recordset("monto") = data_cabemi.Recordset("monto")
                                  infor_det.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
                                  infor_det.Recordset.Update
                                  data_cabemi.Recordset.MoveNext
                               Loop
                               If Xcantlineas < 3 Then
                                  data_cabemi.Recordset.MovePrevious
                                  Do While Xcantlineas <= 3
                                     Xcantlineas = Xcantlineas + 1
                                     infor_det.Recordset.AddNew
                                     infor_det.Recordset("serie") = data_cabemi.Recordset("serie")
                                     infor_det.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
                                     infor_det.Recordset("cod_srv") = Null
                                     infor_det.Recordset("descrip") = Null
                                     infor_det.Recordset("imp_srv") = Null
                                     infor_det.Recordset("nro_linea") = Xcantlineas
                                     infor_det.Recordset("cantidad") = Null
                                     infor_det.Recordset("monto") = Null
                                     infor_det.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
                                     infor_det.Recordset.Update
                                  Loop
                               End If
                               Xcantlineas = 0
                            End If
                       End If
                    Else
                        data_informe.Recordset.AddNew
                        data_informe.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
                        data_informe.Recordset("nom_cnv") = data_emi.Recordset("nom_cnv")
                        data_informe.Recordset("cliente") = data_emi.Recordset("cliente")
                        data_informe.Recordset("apellidos") = data_emi.Recordset("apellidos")
                        data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                        data_informe.Recordset("cedula") = data_emi.Recordset("cedula")
                        data_informe.Recordset("dir_cli") = data_emi.Recordset("dir_cli")
                        data_informe.Recordset("loc_cli") = data_emi.Recordset("loc_cli")
                        data_informe.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
                        data_informe.Recordset("grupo") = data_emi.Recordset("grupo")
                        data_informe.Recordset("promos") = data_emi.Recordset("promo")
                        If data_emi.Recordset("debe_haber") = 101 Then
                           data_informe.Recordset("tipofact") = "e-Ticket"
                        Else
                           data_informe.Recordset("tipofact") = "e-Factura"
                        End If
                        data_informe.Recordset("fpago") = "CREDITO"
                        data_informe.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
                        data_informe.Recordset("documento") = data_emi.Recordset("documento")
                        data_informe.Recordset("importe") = data_emi.Recordset("importe")
                        data_informe.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
                        data_informe.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
                        data_informe.Recordset("mes") = data_emi.Recordset("mes")
                        data_informe.Recordset("ano") = data_emi.Recordset("ano")
                        data_informe.Recordset("color_rec") = data_emi.Recordset("color_rec")
                        data_informe.Recordset("tiquet") = data_emi.Recordset("tiquet")
                        data_informe.Recordset("servi") = data_emi.Recordset("servi")
                        data_informe.Recordset("deudas") = data_emi.Recordset("deudas")
                        data_informe.Recordset("iva") = data_emi.Recordset("iva")
                        data_informe.Recordset("total") = data_emi.Recordset("total")
                        data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                        data_informe.Recordset("zona") = data_emi.Recordset("zona")
                        data_informe.Recordset("qr") = data_emi.Recordset("qr")
                        data_informe.Recordset("fvence") = data_emi.Recordset("fvence")
                        data_informe.Recordset("autoriza") = data_emi.Recordset("autoriza")
                        data_informe.Recordset("rangoCAE") = data_emi.Recordset("rangoCAE")
                        data_informe.Recordset("codseg") = data_emi.Recordset("codseg")
                        If IsNull(data_emi.Recordset("ruc")) = False Then
                           data_informe.Recordset("hora") = Null
                        Else
                           data_informe.Recordset("hora") = "X"
                        End If
                        data_informe.Recordset("fecha") = data_emi.Recordset("fecha")
                        data_informe.Recordset("fecha_cobr") = data_emi.Recordset("fecha_cobr")
                        data_informe.Recordset("nom_superv") = "*" & Trim(str(data_emi.Recordset("cliente"))) & "*"
                        data_informe.Recordset("barranrodoc") = "*" & Trim(str(data_emi.Recordset("documento"))) & "*"
                        data_informe.Recordset("tipocta") = data_emi.Recordset("tipocta")
                        data_informe.Recordset("tipodoc") = data_emi.Recordset("tipodoc")
                        data_informe.Recordset.Update
                        If Nombre = "EMI0516" Or Nombre = "EMI0616" Then
                           data_cabemi.Recordset.AddNew
                           data_cabemi.Recordset("serie") = "A"
                           data_cabemi.Recordset("nro_doc") = data_emi.Recordset("documento")
                           data_cabemi.Recordset("cod_srv") = 881
                           data_cabemi.Recordset("descrip") = "CUOTA MENSUAL " & Trim(str(data_emi.Recordset("mes"))) & "/" & Trim(str(data_emi.Recordset("ano")))
                           data_cabemi.Recordset("imp_srv") = data_emi.Recordset("importe")
                           data_cabemi.Recordset("nro_linea") = 1
                           data_cabemi.Recordset("cantidad") = 1
                           data_cabemi.Recordset("monto") = data_emi.Recordset("total")
                           data_cabemi.Recordset("cliente2") = data_emi.Recordset("cliente")
                           data_cabemi.Recordset.Update
                           data_cabemi.Refresh
                        Else
                           data_cabemi.RecordSource = "Select * from " & NombreCab & " where nro_doc =" & data_emi.Recordset("documento")
                           data_cabemi.Refresh
                        End If
                        Xcantlineas = 0
                        If data_cabemi.Recordset.RecordCount > 0 Then
                           data_cabemi.Recordset.MoveFirst
                           Do While Not data_cabemi.Recordset.EOF
                              Xcantlineas = Xcantlineas + 1
                              infor_det.Recordset.AddNew
                              infor_det.Recordset("serie") = data_cabemi.Recordset("serie")
                              infor_det.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
                              infor_det.Recordset("cod_srv") = data_cabemi.Recordset("cod_srv")
                              If data_cabemi.Recordset("cod_srv") = 881 Then
                                 infor_det.Recordset("descrip") = data_cabemi.Recordset("descrip") & "  " & Trim(str(data_emi.Recordset("mes"))) & "/" & Trim(str(data_emi.Recordset("ano")))
                              Else
                                 infor_det.Recordset("descrip") = data_cabemi.Recordset("descrip")
                              End If
                              infor_det.Recordset("imp_srv") = data_cabemi.Recordset("imp_srv")
                              infor_det.Recordset("nro_linea") = data_cabemi.Recordset("nro_linea")
                              infor_det.Recordset("cantidad") = data_cabemi.Recordset("cantidad")
                              infor_det.Recordset("monto") = data_cabemi.Recordset("monto")
                              infor_det.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
                              infor_det.Recordset.Update
                              data_cabemi.Recordset.MoveNext
                           Loop
                           If Xcantlineas < 3 Then
                              data_cabemi.Recordset.MovePrevious
                              Do While Xcantlineas <= 3
                                 Xcantlineas = Xcantlineas + 1
                                 infor_det.Recordset.AddNew
                                 infor_det.Recordset("serie") = data_cabemi.Recordset("serie")
                                 infor_det.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
                                 infor_det.Recordset("cod_srv") = Null
                                 infor_det.Recordset("descrip") = Null
                                 infor_det.Recordset("imp_srv") = Null
                                 infor_det.Recordset("nro_linea") = Xcantlineas
                                 infor_det.Recordset("cantidad") = Null
                                 infor_det.Recordset("monto") = Null
                                 infor_det.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
                                 infor_det.Recordset.Update
                              Loop
                           End If
                           Xcantlineas = 0
                        End If
                    End If
                    data_emi.Recordset.MoveNext
                  End If
               Loop
            Else
               frm_impemi.MousePointer = 0
               MsgBox "No existen registros", vbInformation, "Mensaje"
            End If
         Else
            If t_raddesde.Text <> "" And t_radhasta.Text <> "" Then
               data_emi.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & txt_cob.Text & " And color_rec ='" & Trim(Xquecolor) & "' and grupo >=" & Val(t_raddesde.Text) & " and grupo <=" & Val(t_radhasta.Text) & " order by nro_cobr,color_rec,documento,importe"
            Else
               data_emi.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & txt_cob.Text & " And color_rec ='" & Trim(Xquecolor) & "' order by nro_cobr,color_rec,documento,importe"
            End If
            data_emi.Refresh
            data_informe.RecordSource = "EMIS"
            data_informe.Refresh
            If data_informe.Recordset.RecordCount > 0 Then
               data_informe.Recordset.MoveFirst
               Do While Not data_informe.Recordset.EOF
                  data_informe.Recordset.Delete
                  data_informe.Recordset.MoveNext
               Loop
               data_informe.Refresh
            End If
            infor_det.RecordSource = "CABEZAL"
            infor_det.Refresh
            If infor_det.Recordset.RecordCount > 0 Then
               infor_det.Recordset.MoveFirst
               Do While Not infor_det.Recordset.EOF
                  infor_det.Recordset.Delete
                  infor_det.Recordset.MoveNext
               Loop
               infor_det.Refresh
            End If
            If data_emi.Recordset.RecordCount > 0 Then
               data_emi.Recordset.MoveFirst
               Do While Not data_emi.Recordset.EOF
                  If data_emi.Recordset("nro_cobr") = 616 Or _
                     data_emi.Recordset("nro_cobr") = 636 Or _
                     data_emi.Recordset("nro_cobr") = 615 Or _
                     data_emi.Recordset("nro_cobr") = 635 Or _
                     data_emi.Recordset("nro_cobr") = 602 Or _
                     data_emi.Recordset("nro_cobr") = 653 Or _
                     data_emi.Recordset("nro_cobr") = 672 Or _
                     data_emi.Recordset("nro_cobr") = 113 Or _
                     data_emi.Recordset("nro_cobr") = 1 Or _
                     data_emi.Recordset("nro_cobr") = 10 Or _
                     data_emi.Recordset("nro_cobr") = 8 Or _
                     data_emi.Recordset("nro_cobr") = 603 Or _
                     data_emi.Recordset("nro_cobr") = 685 Or _
                     data_emi.Recordset("nro_cobr") = 201 Or _
                     data_emi.Recordset("nro_cobr") = 604 Or _
                     data_emi.Recordset("nro_cobr") = 606 Or _
                     data_emi.Recordset("nro_cobr") = 676 Or _
                     data_emi.Recordset("nro_cobr") = 688 Or _
                     data_emi.Recordset("nro_cobr") = 512 Or _
                     data_emi.Recordset("nro_cobr") = 679 Then
                     data_emi.Recordset.MoveNext
                  Else
                    data_informe.Recordset.AddNew
                    data_informe.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
                    data_informe.Recordset("nom_cnv") = data_emi.Recordset("nom_cnv")
                    data_informe.Recordset("cliente") = data_emi.Recordset("cliente")
                    data_informe.Recordset("apellidos") = data_emi.Recordset("apellidos")
                    data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                    data_informe.Recordset("cedula") = data_emi.Recordset("cedula")
                    data_informe.Recordset("dir_cli") = data_emi.Recordset("dir_cli")
                    data_informe.Recordset("loc_cli") = data_emi.Recordset("loc_cli")
                    data_informe.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
                    data_informe.Recordset("grupo") = data_emi.Recordset("grupo")
                    data_informe.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
                    data_informe.Recordset("documento") = data_emi.Recordset("documento")
                    data_informe.Recordset("importe") = data_emi.Recordset("importe")
                    data_informe.Recordset("promos") = data_emi.Recordset("promo")
                    data_informe.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
                    If data_emi.Recordset("debe_haber") = 101 Then
                       data_informe.Recordset("tipofact") = "e-Ticket"
                    Else
                       data_informe.Recordset("tipofact") = "e-Factura"
                    End If
                    data_informe.Recordset("fpago") = "CREDITO"
                    data_informe.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
                    data_informe.Recordset("mes") = data_emi.Recordset("mes")
                    data_informe.Recordset("ano") = data_emi.Recordset("ano")
                    data_informe.Recordset("color_rec") = data_emi.Recordset("color_rec")
                    data_informe.Recordset("tiquet") = data_emi.Recordset("tiquet")
                    data_informe.Recordset("servi") = data_emi.Recordset("servi")
                    data_informe.Recordset("deudas") = data_emi.Recordset("deudas")
                    data_informe.Recordset("iva") = data_emi.Recordset("iva")
                    data_informe.Recordset("total") = data_emi.Recordset("total")
                    data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                    data_informe.Recordset("zona") = data_emi.Recordset("zona")
                    data_informe.Recordset("qr") = data_emi.Recordset("qr")
                    data_informe.Recordset("fvence") = data_emi.Recordset("fvence")
                    data_informe.Recordset("autoriza") = data_emi.Recordset("autoriza")
                    data_informe.Recordset("rangoCAE") = data_emi.Recordset("rangoCAE")
                    data_informe.Recordset("codseg") = data_emi.Recordset("codseg")
                    If IsNull(data_emi.Recordset("ruc")) = False Then
                       data_informe.Recordset("hora") = Null
                    Else
                       data_informe.Recordset("hora") = "X"
                    End If
                    data_informe.Recordset("fecha") = data_emi.Recordset("fecha")
                    data_informe.Recordset("fecha_cobr") = data_emi.Recordset("fecha_cobr")
                    data_informe.Recordset("nom_superv") = "*" & Trim(str(data_emi.Recordset("cliente"))) & "*"
                    data_informe.Recordset("barranrodoc") = "*" & Trim(str(data_emi.Recordset("documento"))) & "*"
                    data_informe.Recordset("tipocta") = data_emi.Recordset("tipocta")
                    data_informe.Recordset("tipodoc") = data_emi.Recordset("tipodoc")
                    data_informe.Recordset.Update
                    If Nombre = "EMI0516" Or Nombre = "EMI0616" Then
                       data_cabemi.Recordset.AddNew
                       data_cabemi.Recordset("serie") = "A"
                       data_cabemi.Recordset("nro_doc") = data_emi.Recordset("documento")
                       data_cabemi.Recordset("cod_srv") = 881
                       data_cabemi.Recordset("descrip") = "CUOTA MENSUAL " & Trim(str(data_emi.Recordset("mes"))) & "/" & Trim(str(data_emi.Recordset("ano")))
                       data_cabemi.Recordset("imp_srv") = data_emi.Recordset("importe")
                       data_cabemi.Recordset("nro_linea") = 1
                       data_cabemi.Recordset("cantidad") = 1
                       data_cabemi.Recordset("monto") = data_emi.Recordset("total")
                       data_cabemi.Recordset("cliente2") = data_emi.Recordset("cliente")
                       data_cabemi.Recordset.Update
                       data_cabemi.Refresh
                    Else
                       data_cabemi.RecordSource = "Select * from " & NombreCab & " where nro_doc =" & data_emi.Recordset("documento")
                       data_cabemi.Refresh
                    End If
                    Xcantlineas = 0
                    If data_cabemi.Recordset.RecordCount > 0 Then
                       data_cabemi.Recordset.MoveFirst
                       Do While Not data_cabemi.Recordset.EOF
                          Xcantlineas = Xcantlineas + 1
                          infor_det.Recordset.AddNew
                          infor_det.Recordset("serie") = data_cabemi.Recordset("serie")
                          infor_det.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
                          infor_det.Recordset("cod_srv") = data_cabemi.Recordset("cod_srv")
                          If data_cabemi.Recordset("cod_srv") = 881 Then
                             infor_det.Recordset("descrip") = data_cabemi.Recordset("descrip") & "  " & Trim(str(data_emi.Recordset("mes"))) & "/" & Trim(str(data_emi.Recordset("ano")))
                          Else
                             infor_det.Recordset("descrip") = data_cabemi.Recordset("descrip")
                          End If
                          infor_det.Recordset("imp_srv") = data_cabemi.Recordset("imp_srv")
                          infor_det.Recordset("nro_linea") = data_cabemi.Recordset("nro_linea")
                          infor_det.Recordset("cantidad") = data_cabemi.Recordset("cantidad")
                          infor_det.Recordset("monto") = data_cabemi.Recordset("monto")
                          infor_det.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
                          infor_det.Recordset.Update
                          data_cabemi.Recordset.MoveNext
                       Loop
                       If Xcantlineas < 3 Then
                          data_cabemi.Recordset.MovePrevious
                          Do While Xcantlineas <= 3
                             Xcantlineas = Xcantlineas + 1
                             infor_det.Recordset.AddNew
                             infor_det.Recordset("serie") = data_cabemi.Recordset("serie")
                             infor_det.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
                             infor_det.Recordset("cod_srv") = Null
                             infor_det.Recordset("descrip") = Null
                             infor_det.Recordset("imp_srv") = Null
                             infor_det.Recordset("nro_linea") = Xcantlineas
                             infor_det.Recordset("cantidad") = Null
                             infor_det.Recordset("monto") = Null
                             infor_det.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
                             infor_det.Recordset.Update
                          Loop
                       End If
                       Xcantlineas = 0
                    End If
                    data_emi.Recordset.MoveNext
                  End If
               Loop
            Else
               frm_impemi.MousePointer = 0
               MsgBox "No existen registros", vbInformation, "Mensaje"
            End If
         End If
      End If
      infor_det.Refresh
      data_informe.RecordSource = "Select * from emis order by nro_cobr,importe"
      data_informe.Refresh
      If Check2.Value <> 1 Then
        If data_informe.Recordset.RecordCount > 0 Then
           data_informe.Recordset.MoveFirst
           Do While Not data_informe.Recordset.EOF
              If data_informe.Recordset("promos") = "Pago anual" Then
                 infor_det.RecordSource = "select * from CABEZAL where cliente2 =" & data_informe.Recordset("cliente")
                 infor_det.Refresh
                 If infor_det.Recordset.RecordCount > 0 Then
                    infor_det.Recordset.MoveFirst
                    Do While Not infor_det.Recordset.EOF
                       infor_det.Recordset.Delete
                       infor_det.Recordset.MoveNext
                    Loop
                 End If
                 data_informe.Recordset.Delete
              End If
              data_informe.Recordset.MoveNext
           Loop
        End If
      End If
      infor_det.Refresh
      data_informe.RecordSource = "Select * from emis order by nro_cobr,importe"
      data_informe.Refresh
      frm_impemi.MousePointer = 0
    
      If data_informe.Recordset.RecordCount > 0 Then
         If Check2.Value = 1 Then
            cremi.ReportFileName = App.path & "\infrecnewanual.rpt"
            cremi.Action = 1
         Else
            If Xquecolor = "C" Then
               cremi.ReportFileName = App.path & "\faccnvnew.rpt"
               cremi.Action = 1
            Else
               If Check1.Value = 1 Then
                  cremi.ReportFileName = App.path & "\infrecnew.rpt"
                  cremi.Action = 1
               Else
                  cremi.ReportFileName = App.path & "\infrecnew.rpt"
                  cremi.Action = 1
               End If
            End If
         End If
      Else
         frm_impemi.MousePointer = 0
         MsgBox "No hay registros"
      End If
   End If
End If
frm_impemi.MousePointer = 0

End Sub

Private Sub b_canc_Click()
Unload Me

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   txt_cob.Text = ""
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_desde.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim Xquecolor As String
Text1.Text = 5
Nombre = "EMI"
If txt_m.Text > 9 Then
   Nombre = Nombre + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
Else
   Nombre = Nombre + "0" + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
End If

If Combo1.ListIndex = 0 Then
   Xquecolor = "R"
End If
If Combo1.ListIndex = 1 Then
   Xquecolor = "A"
End If
If Combo1.ListIndex = 2 Then
   Xquecolor = "M"
End If
If Combo1.ListIndex = 3 Then
   Xquecolor = "V"
End If

If txt_m.Text <> "" Then
   If txt_a.Text <> "" Then
      If txt_desde.Text <> "" Then
         If txt_hasta.Text <> "" Then
            data_emi.RecordSource = "Select * from " & Nombre & " where documento >=" & txt_desde.Text & " And documento <=" & txt_hasta.Text & " order by nro_cobr,color_rec,documento,importe"
            data_emi.Refresh
            If data_informe.Recordset.RecordCount > 0 Then
               data_informe.Recordset.MoveFirst
               Do While Not data_informe.Recordset.EOF
                  data_informe.Recordset.Delete
                  data_informe.Recordset.MoveNext
               Loop
               data_informe.Refresh
            End If
            If data_emi.Recordset.RecordCount > 0 Then
               data_emi.Recordset.MoveFirst
               Do While Not data_emi.Recordset.EOF
                  If data_emi.Recordset("nro_cobr") = 10 Then
                     data_emi.Recordset.MoveNext
                  Else
                    data_informe.Recordset.AddNew
                    data_informe.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
                    data_informe.Recordset("nom_cnv") = data_emi.Recordset("nom_cnv")
                    data_informe.Recordset("cliente") = data_emi.Recordset("cliente")
                    data_informe.Recordset("apellidos") = data_emi.Recordset("apellidos")
                    data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                    data_informe.Recordset("cedula") = data_emi.Recordset("cedula")
                    data_informe.Recordset("dir_cli") = data_emi.Recordset("dir_cli")
                    data_informe.Recordset("loc_cli") = data_emi.Recordset("loc_cli")
                    data_informe.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
                    data_informe.Recordset("grupo") = data_emi.Recordset("grupo")
                    data_informe.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
                    data_informe.Recordset("documento") = data_emi.Recordset("documento")
                    data_informe.Recordset("importe") = data_emi.Recordset("importe")
                    data_informe.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
                    data_informe.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
                    data_informe.Recordset("mes") = data_emi.Recordset("mes")
                    data_informe.Recordset("ano") = data_emi.Recordset("ano")
                    data_informe.Recordset("color_rec") = data_emi.Recordset("color_rec")
                    data_informe.Recordset("tiquet") = data_emi.Recordset("tiquet")
                    data_informe.Recordset("servi") = data_emi.Recordset("servi")
                    data_informe.Recordset("deudas") = data_emi.Recordset("deudas")
                    data_informe.Recordset("iva") = data_emi.Recordset("iva")
                    data_informe.Recordset("total") = data_emi.Recordset("total")
                    data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                    data_informe.Recordset("zona") = data_emi.Recordset("zona")
                    
                    data_informe.Recordset.Update
                    data_emi.Recordset.MoveNext
                  End If
               Loop
            Else
               MsgBox "No existen registros", vbInformation, "Mensaje"
            End If
         End If
      Else
         If txt_cob.Text = "" Then
            If Check1.Value = 1 Then
               data_emi.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & 11 & " or nro_cobr =" & 6 & " or nro_cobr =" & 5 & " order by nro_cobr,color_rec,importe"
               data_emi.Refresh
            Else
               data_emi.RecordSource = "Select * from " & Nombre & " where color_rec ='" & Trim(Xquecolor) & "' order by nro_cobr,color_rec,documento,importe"
               data_emi.Refresh
            End If
            If data_informe.Recordset.RecordCount > 0 Then
               data_informe.Recordset.MoveFirst
               Do While Not data_informe.Recordset.EOF
                  data_informe.Recordset.Delete
                  data_informe.Recordset.MoveNext
               Loop
               data_informe.Refresh
            End If
            If data_emi.Recordset.RecordCount > 0 Then
               data_emi.Recordset.MoveFirst
               Do While Not data_emi.Recordset.EOF
                  If data_emi.Recordset("nro_cobr") = 616 Or _
                     data_emi.Recordset("nro_cobr") = 636 Or _
                     data_emi.Recordset("nro_cobr") = 615 Or _
                     data_emi.Recordset("nro_cobr") = 635 Or _
                     data_emi.Recordset("nro_cobr") = 602 Or _
                     data_emi.Recordset("nro_cobr") = 653 Or _
                     data_emi.Recordset("nro_cobr") = 672 Or _
                     data_emi.Recordset("nro_cobr") = 113 Or _
                     data_emi.Recordset("nro_cobr") = 1 Or _
                     data_emi.Recordset("nro_cobr") = 10 Or _
                     data_emi.Recordset("nro_cobr") = 8 Or _
                     data_emi.Recordset("nro_cobr") = 603 Or _
                     data_emi.Recordset("nro_cobr") = 685 Or _
                     data_emi.Recordset("nro_cobr") = 201 Or _
                     data_emi.Recordset("nro_cobr") = 604 Or _
                     data_emi.Recordset("nro_cobr") = 606 Or _
                     data_emi.Recordset("nro_cobr") = 676 Or _
                     data_emi.Recordset("nro_cobr") = 688 Or _
                     data_emi.Recordset("nro_cobr") = 512 Or _
                     data_emi.Recordset("nro_cobr") = 679 Then
                     data_emi.Recordset.MoveNext
                  Else
                    data_informe.Recordset.AddNew
                    data_informe.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
                    data_informe.Recordset("nom_cnv") = data_emi.Recordset("nom_cnv")
                    data_informe.Recordset("cliente") = data_emi.Recordset("cliente")
                    data_informe.Recordset("apellidos") = data_emi.Recordset("apellidos")
                    data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                    data_informe.Recordset("cedula") = data_emi.Recordset("cedula")
                    data_informe.Recordset("dir_cli") = data_emi.Recordset("dir_cli")
                    data_informe.Recordset("loc_cli") = data_emi.Recordset("loc_cli")
                    data_informe.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
                    data_informe.Recordset("grupo") = data_emi.Recordset("grupo")
                    data_informe.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
                    data_informe.Recordset("documento") = data_emi.Recordset("documento")
                    data_informe.Recordset("importe") = data_emi.Recordset("importe")
                    data_informe.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
                    data_informe.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
                    data_informe.Recordset("mes") = data_emi.Recordset("mes")
                    data_informe.Recordset("ano") = data_emi.Recordset("ano")
                    data_informe.Recordset("color_rec") = data_emi.Recordset("color_rec")
                    data_informe.Recordset("tiquet") = data_emi.Recordset("tiquet")
                    data_informe.Recordset("servi") = data_emi.Recordset("servi")
                    data_informe.Recordset("deudas") = data_emi.Recordset("deudas")
                    data_informe.Recordset("iva") = data_emi.Recordset("iva")
                    data_informe.Recordset("total") = data_emi.Recordset("total")
                    data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                    data_informe.Recordset("zona") = data_emi.Recordset("zona")
                    data_informe.Recordset.Update
                    data_emi.Recordset.MoveNext
                  End If
               Loop
            Else
               MsgBox "No existen registros", vbInformation, "Mensaje"
            End If
         Else
            data_emi.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & txt_cob.Text & " And color_rec ='" & Trim(Xquecolor) & "' order by nro_cobr,color_rec,documento,importe"
            data_emi.Refresh
            If data_informe.Recordset.RecordCount > 0 Then
               data_informe.Recordset.MoveFirst
               Do While Not data_informe.Recordset.EOF
                  data_informe.Recordset.Delete
                  data_informe.Recordset.MoveNext
               Loop
            End If
            If data_emi.Recordset.RecordCount > 0 Then
               data_emi.Recordset.MoveFirst
               Do While Not data_emi.Recordset.EOF
                  data_informe.Recordset.AddNew
                  data_informe.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
                  data_informe.Recordset("nom_cnv") = data_emi.Recordset("nom_cnv")
                  data_informe.Recordset("cliente") = data_emi.Recordset("cliente")
                  data_informe.Recordset("apellidos") = data_emi.Recordset("apellidos")
                  data_informe.Recordset("ruc") = data_emi.Recordset("ruc")
                  data_informe.Recordset("cedula") = data_emi.Recordset("cedula")
                  data_informe.Recordset("dir_cli") = data_emi.Recordset("dir_cli")
                  data_informe.Recordset("loc_cli") = data_emi.Recordset("loc_cli")
                  data_informe.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
                  data_informe.Recordset("grupo") = data_emi.Recordset("grupo")
                  data_informe.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
                  data_informe.Recordset("documento") = data_emi.Recordset("documento")
                  data_informe.Recordset("importe") = data_emi.Recordset("importe")
                  data_informe.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
                  data_informe.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
                  data_informe.Recordset("mes") = data_emi.Recordset("mes")
                  data_informe.Recordset("ano") = data_emi.Recordset("ano")
                  data_informe.Recordset("color_rec") = data_emi.Recordset("color_rec")
                  data_informe.Recordset("tiquet") = data_emi.Recordset("tiquet")
                  data_informe.Recordset("servi") = data_emi.Recordset("servi")
                  data_informe.Recordset("deudas") = data_emi.Recordset("deudas")
                  data_informe.Recordset("iva") = data_emi.Recordset("iva")
                  data_informe.Recordset("total") = data_emi.Recordset("total")
                  data_informe.Recordset("zona") = data_emi.Recordset("zona")
                  data_informe.Recordset.Update
                  data_emi.Recordset.MoveNext
               Loop
            Else
               MsgBox "No existen registros", vbInformation, "Mensaje"
            End If
         End If
      End If
      data_informe.RecordSource = "Select * from infemirec order by nro_cobr,color_rec,documento,importe"
      data_informe.Refresh
      If data_informe.Recordset.RecordCount > 0 Then
         If Check1.Value = 1 Then
            cremi.ReportFileName = App.path & "\rspsapsj.rpt"
            cremi.Action = 1
         Else
            cremi.ReportFileName = App.path & "\rspsapp.rpt"
            cremi.Action = 1
         End If
      End If
   End If
End If

End Sub

Private Sub Form_Load()
Dim Xmes, Xano As Long
Xmes = Month(Date)
Xano = Year(Date)
data_informe.DatabaseName = App.path & "\infemis.mdb"
infor_det.DatabaseName = App.path & "\infemis.mdb"

data_emi.DatabaseName = App.path & "\emisnueva.mdb"
'Nombre = "EMI"
'If Xmes > 9 Then
'   Nombre = Nombre + Trim(Str(Xmes)) + Mid(Trim(Str(Xano)), 3, 2)
'Else
'   Nombre = Nombre + "0" + Trim(Str(Xmes)) + Mid(Trim(Str(Xano)), 3, 2)
'End If
'data_emi.RecordSource = Nombre
'data_emi.Refresh
data_cabemi.DatabaseName = App.path & "\emisnueva.mdb"

txt_m.Text = Xmes
txt_a.Text = Xano
Combo1.ListIndex = 0

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_raddesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_radhasta.SetFocus
End If

End Sub

Private Sub txt_a_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cob.SetFocus
End If

End Sub

Private Sub txt_cob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub txt_desde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_hasta.SetFocus
End If

End Sub

Private Sub txt_hasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub

Private Sub txt_m_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_a.SetFocus
End If

End Sub
