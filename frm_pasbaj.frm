VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pasbaj 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasaje de BAJAS en arqueo"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frm_pasbaj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   375
      Left            =   840
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSAdodcLib.Adodc data_cob 
      Height          =   375
      Left            =   3240
      Top             =   120
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
      Caption         =   "data_cob"
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
   Begin MSAdodcLib.Adodc data_deu 
      Height          =   375
      Left            =   4200
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "data_deu"
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
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   4440
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSAdodcLib.Adodc data_arqueo 
      Height          =   375
      Left            =   2160
      Top             =   960
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
      Caption         =   "data_arqueo"
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
   Begin VB.CommandButton btn_salir 
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
      Left            =   1320
      Picture         =   "frm_pasbaj.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   840
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Pasaje de recibos BAJAS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6975
      Begin VB.CommandButton btn_fin 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Terminar"
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
         Left            =   120
         Picture         =   "frm_pasbaj.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txt_rec 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
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
         Left            =   1800
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txt_mat 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
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
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label labnom 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label labcantp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Left            =   4800
         TabIndex        =   12
         Top             =   2160
         Width           =   1680
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "Total de importe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label labcantr 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Left            =   4920
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cantidad de Recibos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   4320
         X2              =   4320
         Y1              =   120
         Y2              =   3000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Nro. Recibo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Matrícula:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton btn_acep 
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
      Picture         =   "frm_pasbaj.frx":0E0E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt_nrocobr 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   1
      EndProperty
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
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label labnomcob 
      BackColor       =   &H00C0E0FF&
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
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "COBRADOR:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   3360
      Picture         =   "frm_pasbaj.frx":1398
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frm_pasbaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_acep_Click()
If txt_nrocobr.Text <> "" Then
   data_cob.RecordSource = "select * from cobrador where cb_numero =" & txt_nrocobr.Text
   data_cob.Refresh
   If data_cob.Recordset.RecordCount > 0 Then
      txt_nrocobr.Text = data_cob.Recordset("cb_numero")
      labnomcob.Caption = data_cob.Recordset("cb_nombre")
      btn_acep.Enabled = False
      txt_nrocobr.Enabled = False
      Frame1.Enabled = True
      txt_mat.SetFocus
      data_arqueo.RecordSource = "Select * from arqueo where cob =" & txt_nrocobr.Text & " and codpro =" & 98
      data_arqueo.Refresh
      If data_arqueo.Recordset.RecordCount > 0 Then
         MsgBox "EL ARQUEO PARA ESTE COBRADOR ESTA CERRADO.", vbInformation
         Unload Me
      End If
   Else
      MsgBox "No existe cobrador, Verifique", vbInformation, "Arqueos"
      btn_acep.Enabled = True
      txt_nrocobr.Enabled = True
      Frame1.Enabled = False
      txt_nrocobr.SetFocus
   End If
   data_cob.Recordset.Close
   
Else
   MsgBox "No ingresó cobrador, verifique", vbCritical, "Arqueos"
   btn_acep.Enabled = True
   txt_nrocobr.Enabled = True
   Frame1.Enabled = False
   txt_nrocobr.SetFocus
End If

End Sub

Private Sub btn_fin_Click()
Frame1.Enabled = False
btn_acep.Enabled = True
txt_nrocobr.Enabled = True
txt_nrocobr.Text = ""
txt_nrocobr.SetFocus
labnomcob.Caption = ""
End Sub

Private Sub btn_salir_Click()
frm_pasbaj.Hide

End Sub

Private Sub Form_Load()
'data_arqueo.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_arqueo.ConnectionString = "dsn=" & Xconexrmt
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
'data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.ConnectionString = "dsn=" & Xconexrmt
'data_deu.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deu.ConnectionString = "dsn=" & Xconexrmt
data_lin.ConnectionString = "dsn=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_rec.SetFocus
End If
   
End Sub

Private Sub txt_mat_LostFocus()
If txt_mat.Text <> "" Then
'   data_cli.Recordset.FindFirst "cl_codigo =" & txt_mat.Text
   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & txt_mat.Text
   data_cli.Refresh
'   If Not data_cli.Recordset.NoMatch Then
   If data_cli.Recordset.RecordCount > 0 Then
      labnom.Caption = data_cli.Recordset("cl_apellid")
   Else
      labnom.Caption = ""
   End If
   data_cli.Recordset.Close
   
End If

End Sub

Private Sub txt_nrocobr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_nrocobr.Text <> "" Then
      data_cob.RecordSource = "select * from cobrador where cb_numero =" & txt_nrocobr.Text
      data_cob.Refresh
      If data_cob.Recordset.RecordCount > 0 Then
         txt_nrocobr.Text = data_cob.Recordset("cb_numero")
         labnomcob.Caption = data_cob.Recordset("cb_nombre")
      Else
         MsgBox "No existe cobrador, Verifique", vbInformation, "Arqueos"
         txt_nrocobr.SetFocus
      End If
      data_cob.Recordset.Close
   End If
   btn_acep.SetFocus
End If

End Sub

Private Sub txt_rec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mat.SetFocus
End If

End Sub

Private Sub txt_rec_LostFocus()
If txt_rec.Text <> "" Then
'   data_arqueo.Recordset.FindFirst "Matricula =" & txt_mat.Text & " And Nrorec =" & txt_rec.Text
   data_arqueo.RecordSource = "Select * from arqueo where matricula =" & txt_mat.Text & " and nrorec =" & txt_rec.Text
   data_arqueo.Refresh
'   If Not data_arqueo.Recordset.NoMatch Then
   If data_arqueo.Recordset.RecordCount > 0 Then
      If data_arqueo.Recordset("cob") = txt_nrocobr.Text Then
         If IsNull(data_arqueo.Recordset("arqueo")) = False Then
            If IsNull(data_arqueo.Recordset("codpro")) = False Then
               If data_arqueo.Recordset("codpro") = 98 Then
                  MsgBox "Arqueo CERRADO.", vbCritical
               Else
                    If data_arqueo.Recordset("arqueo") = "B" Then
                    Else
         '              data_arqueo.Recordset.Edit
                       data_arqueo.Recordset("arqueo") = "B"
                       data_arqueo.Recordset("fecha") = Date
                       data_arqueo.Recordset("usuar") = WElusuario
                       data_arqueo.Recordset.Update
                       data_deu.RecordSource = "Select * from deudas where cliente =" & data_arqueo.Recordset("matricula") & " and mes =" & data_arqueo.Recordset("mes") & " and ano =" & data_arqueo.Recordset("ano")
                       data_deu.Refresh
                       If data_deu.Recordset.RecordCount > 0 Then
                          If IsNull(data_deu.Recordset("fecha_pago")) = False Then
        '                     data_deu.Recordset.Edit
                             data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_arqueo.Recordset("matricula") & " and mes_paga =" & data_arqueo.Recordset("mes") & " and ano_paga =" & data_arqueo.Recordset("ano")
                             data_lin.Refresh
                             If data_lin.Recordset.RecordCount > 0 Then
                                MsgBox "La factura figura paga en base, debe pasar el comprobante como DEVOLUCIÓN"
                             Else
                                data_deu.Recordset("fecha_pago") = Null
                                data_deu.Recordset.Update
                             End If
                             data_lin.Recordset.Close
                          Else
                             data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_arqueo.Recordset("matricula") & " and mes_paga =" & data_arqueo.Recordset("mes") & " and ano_paga =" & data_arqueo.Recordset("ano")
                             data_lin.Refresh
                             If data_lin.Recordset.RecordCount > 0 Then
                                MsgBox "La factura figura paga en base, debe pasar el comprobante como DEVOLUCIÓN"
                                data_deu.Recordset("fecha_pago") = data_lin.Recordset("fecha")
                                data_deu.Recordset.Update
                             Else
                             End If
                             data_lin.Recordset.Close
                          End If
                       Else
                          data_deu.Recordset.AddNew
                          data_deu.Recordset("cod_cnv") = data_arqueo.Recordset("cat")
                          data_deu.Recordset("nom_cnv") = Mid(data_arqueo.Recordset("nomcat"), 1, 20)
                          data_deu.Recordset("tipocta") = "A"
                          data_deu.Recordset("cliente") = data_arqueo.Recordset("matricula")
                          data_deu.Recordset("nombre") = data_arqueo.Recordset("nombre")
                          data_deu.Recordset("fecha") = data_arqueo.Recordset("fecha")
                          data_deu.Recordset("tipodoc") = "FAC"
                          data_deu.Recordset("documento") = data_arqueo.Recordset("nrorec")
                          data_deu.Recordset("importe") = data_arqueo.Recordset("importe")
                          data_deu.Recordset("moneda") = 1
                          data_deu.Recordset("origen") = "Pendiente..." & Trim(str(data_arqueo.Recordset("mes"))) & "/" & Trim(str(data_arqueo.Recordset("ano")))
                          data_deu.Recordset("nro_vende") = data_arqueo.Recordset("codpro")
                          data_deu.Recordset("grupo") = data_arqueo.Recordset("codzon")
                          data_deu.Recordset("saldo_cc") = 0
                          data_deu.Recordset("mes") = data_arqueo.Recordset("mes")
                          data_deu.Recordset("ano") = data_arqueo.Recordset("ano")
                          data_deu.Recordset("nro_cobr") = data_arqueo.Recordset("cob")
                          data_deu.Recordset("nom_cobr") = data_arqueo.Recordset("nomcob")
                          data_deu.Recordset("estado_cta") = 2
                          data_deu.Recordset("tiquet") = data_arqueo.Recordset("tiquet")
                          data_deu.Recordset("deudas") = data_arqueo.Recordset("deudas")
                          data_deu.Recordset("total") = data_arqueo.Recordset("total")
                          data_deu.Recordset("servi") = data_arqueo.Recordset("servi")
                          data_deu.Recordset("iva") = data_arqueo.Recordset("iva")
                          data_deu.Recordset.Update
                          data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_arqueo.Recordset("cliente")
                          data_cli.Refresh
                          If data_cli.Recordset.RecordCount > 0 Then
                             If IsNull(data_cli.Recordset("saldo_cc")) = False Then
        '                        data_cli.Recordset.Edit
                                data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") + data_arqueo.Recordset("total")
                                If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
                                   data_cli.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa") + 1
                                End If
                                data_cli.Recordset.Update
                             Else
        '                        data_cli.Recordset.Edit
                                data_cli.Recordset("saldo_cc") = data_arqueo.Recordset("total")
                                If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
                                   data_cli.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa") + 1
                                Else
                                  data_cli.Recordset("cl_atrasoa") = 1
                                End If
                                data_cli.Recordset.Update
                             End If
                          End If
                          data_cli.Recordset.Close
                       End If
                       If labcantr.Caption <> "" Then
                          labcantr.Caption = labcantr.Caption + 1
                       Else
                          labcantr.Caption = 1
                       End If
                       If labcantp.Caption <> "" Then
                          labcantp.Caption = labcantp.Caption + data_arqueo.Recordset("total")
                       Else
                          labcantp.Caption = data_arqueo.Recordset("total")
                       End If
                       data_deu.Recordset.Close
                       
                    End If
                End If
            Else
                   data_arqueo.Recordset("arqueo") = "B"
                   data_arqueo.Recordset("fecha") = Date
                   data_arqueo.Recordset("usuar") = WElusuario
                   data_arqueo.Recordset.Update
                   data_deu.RecordSource = "Select * from deudas where cliente =" & data_arqueo.Recordset("matricula") & " and mes =" & data_arqueo.Recordset("mes") & " and ano =" & data_arqueo.Recordset("ano")
                   data_deu.Refresh
                   If data_deu.Recordset.RecordCount > 0 Then
                      If IsNull(data_deu.Recordset("fecha_pago")) = False Then
    '                     data_deu.Recordset.Edit
                         data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_arqueo.Recordset("matricula") & " and mes_paga =" & data_arqueo.Recordset("mes") & " and ano_paga =" & data_arqueo.Recordset("ano")
                         data_lin.Refresh
                         If data_lin.Recordset.RecordCount > 0 Then
                            MsgBox "La factura figura paga en base, debe pasar el comprobante como DEVOLUCIÓN"
                         Else
                            data_deu.Recordset("fecha_pago") = Null
                            data_deu.Recordset.Update
                         End If
                         data_lin.Recordset.Close
                      Else
                         data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_arqueo.Recordset("matricula") & " and mes_paga =" & data_arqueo.Recordset("mes") & " and ano_paga =" & data_arqueo.Recordset("ano")
                         data_lin.Refresh
                         If data_lin.Recordset.RecordCount > 0 Then
                            MsgBox "La factura figura paga en base, debe pasar el comprobante como DEVOLUCIÓN"
                            data_deu.Recordset("fecha_pago") = data_lin.Recordset("fecha")
                            data_deu.Recordset.Update
                         Else
                         End If
                         data_lin.Recordset.Close
                      End If
                   Else
                      data_deu.Recordset.AddNew
                      data_deu.Recordset("cod_cnv") = data_arqueo.Recordset("cat")
                      data_deu.Recordset("nom_cnv") = Mid(data_arqueo.Recordset("nomcat"), 1, 20)
                      data_deu.Recordset("tipocta") = "A"
                      data_deu.Recordset("cliente") = data_arqueo.Recordset("matricula")
                      data_deu.Recordset("nombre") = data_arqueo.Recordset("nombre")
                      data_deu.Recordset("fecha") = data_arqueo.Recordset("fecha")
                      data_deu.Recordset("tipodoc") = "FAC"
                      data_deu.Recordset("documento") = data_arqueo.Recordset("nrorec")
                      data_deu.Recordset("importe") = data_arqueo.Recordset("importe")
                      data_deu.Recordset("moneda") = 1
                      data_deu.Recordset("origen") = "Pendiente..." & Trim(str(data_arqueo.Recordset("mes"))) & "/" & Trim(str(data_arqueo.Recordset("ano")))
                      data_deu.Recordset("nro_vende") = data_arqueo.Recordset("codpro")
                      data_deu.Recordset("grupo") = data_arqueo.Recordset("codzon")
                      data_deu.Recordset("saldo_cc") = 0
                      data_deu.Recordset("mes") = data_arqueo.Recordset("mes")
                      data_deu.Recordset("ano") = data_arqueo.Recordset("ano")
                      data_deu.Recordset("nro_cobr") = data_arqueo.Recordset("cob")
                      data_deu.Recordset("nom_cobr") = data_arqueo.Recordset("nomcob")
                      data_deu.Recordset("estado_cta") = 2
                      data_deu.Recordset("tiquet") = data_arqueo.Recordset("tiquet")
                      data_deu.Recordset("deudas") = data_arqueo.Recordset("deudas")
                      data_deu.Recordset("total") = data_arqueo.Recordset("total")
                      data_deu.Recordset("servi") = data_arqueo.Recordset("servi")
                      data_deu.Recordset("iva") = data_arqueo.Recordset("iva")
                      data_deu.Recordset.Update
                      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_arqueo.Recordset("cliente")
                      data_cli.Refresh
                      If data_cli.Recordset.RecordCount > 0 Then
                         If IsNull(data_cli.Recordset("saldo_cc")) = False Then
    '                        data_cli.Recordset.Edit
                            data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") + data_arqueo.Recordset("total")
                            If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
                               data_cli.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa") + 1
                            End If
                            data_cli.Recordset.Update
                         Else
    '                        data_cli.Recordset.Edit
                            data_cli.Recordset("saldo_cc") = data_arqueo.Recordset("total")
                            If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
                               data_cli.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa") + 1
                            Else
                              data_cli.Recordset("cl_atrasoa") = 1
                            End If
                            data_cli.Recordset.Update
                         End If
                      End If
                      data_cli.Recordset.Close
                   End If
                   If labcantr.Caption <> "" Then
                      labcantr.Caption = labcantr.Caption + 1
                   Else
                      labcantr.Caption = 1
                   End If
                   If labcantp.Caption <> "" Then
                      labcantp.Caption = labcantp.Caption + data_arqueo.Recordset("total")
                   Else
                      labcantp.Caption = data_arqueo.Recordset("total")
                   End If
                   data_deu.Recordset.Close
            End If
         Else
            MsgBox "Hay un error, no se puede pasar el recibo", vbCritical
         End If
      Else
         MsgBox "Figura otro número de cobrador, VERIFIQUE: " & str(data_arqueo.Recordset("cob")), vbInformation, "Arqueos"
      End If
   Else
      MsgBox "No se encontró recibo, VERIFIQUE COBRADOR Y RECIBO", vbInformation, "Arqueos"
   End If
   data_arqueo.Recordset.Close
   
End If
txt_mat.Text = ""
txt_rec.Text = ""

End Sub
