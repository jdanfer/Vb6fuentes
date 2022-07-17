VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_carfaccnv 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar nuevas entregas al arqueo"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_carfaccnv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cargar nuevas entregas"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.Data data_cab 
      Caption         =   "data_cab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_notascr 
      Caption         =   "data_notascr"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_cob 
      Caption         =   "data_cob"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_fac 
      Caption         =   "data_fac"
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
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_arq 
      Caption         =   "data_arq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton b_proc 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin MSMask.MaskEdBox mfech 
      Height          =   375
      Left            =   4320
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
   Begin MSMask.MaskEdBox mfecd 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Fecha desde es día siguiente a la emisión"
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
      Caption         =   "Rango de Fechas:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   3000
      Picture         =   "frm_carfaccnv.frx":058A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frm_carfaccnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_proc_Click()
Dim Xemi As String
Dim CantFact As Integer
CantFact = 0
If Month(mfecd.Text) > 9 Then
   Xemi = "emi" & Trim(str(Month(CDate(mfecd.Text)))) & Mid(Trim(str(Year(CDate(mfecd.Text)))), 3, 2)
Else
   Xemi = "emi0" & Trim(str(Month(CDate(mfecd.Text)))) & Mid(Trim(str(Year(CDate(mfecd.Text)))), 3, 2)
End If

If mfecd.Text <> "__/__/____" And mfech.Text <> "__/__/____" Then
   If Check1.Value = 1 Then
      data_emi.RecordSource = "Select * from " & Trim(Xemi) & " where fecha >=#" & Format(mfecd.Text, "yyyy/mm/dd") & "#"
      data_emi.Refresh
      b_proc.Enabled = False
      frm_carfaccnv.MousePointer = 11
      If data_emi.Recordset.RecordCount > 0 Then
         data_emi.Recordset.MoveFirst
         Do While Not data_emi.Recordset.EOF
            data_arq.RecordSource = "Select * from arqueo where matricula =" & data_emi.Recordset("cliente") & " and nrorec =" & data_emi.Recordset("documento")
            data_arq.Refresh
            If data_arq.Recordset.RecordCount > 0 Then
            Else
               data_arq.Recordset.AddNew
               data_arq.Recordset("matricula") = data_emi.Recordset("cliente")
               data_arq.Recordset("nombre") = data_emi.Recordset("apellidos")
               data_arq.Recordset("mes") = data_emi.Recordset("mes")
               data_arq.Recordset("ano") = data_emi.Recordset("ano")
               data_arq.Recordset("color") = data_emi.Recordset("color_rec")
               data_arq.Recordset("cat") = data_emi.Recordset("cod_cnv")
               data_arq.Recordset("nomcat") = data_emi.Recordset("nom_cnv")
               data_arq.Recordset("arqueo") = "E"
               data_arq.Recordset("importe") = data_emi.Recordset("importe")
               data_arq.Recordset("fecha") = Date
               data_arq.Recordset("nrorec") = data_emi.Recordset("documento")
               data_arq.Recordset("usuar") = WElusuario
               data_arq.Recordset("moneda") = data_emi.Recordset("moneda")
               data_arq.Recordset("cob") = data_emi.Recordset("nro_cobr")
               data_arq.Recordset("nomcob") = data_emi.Recordset("nom_cobr")
               If IsNull(data_emi.Recordset("grupo")) = False Then
                  data_arq.Recordset("codzon") = data_emi.Recordset("grupo")
               Else
                  data_arq.Recordset("codzon") = 0
               End If
               data_arq.Recordset("codsup") = data_emi.Recordset("nro_superv")
               data_arq.Recordset("codpro") = data_emi.Recordset("nro_vende")
               data_arq.Recordset("tiquet") = data_emi.Recordset("tiquet")
               data_arq.Recordset("total") = data_emi.Recordset("total")
               data_arq.Recordset("varia") = data_emi.Recordset("deudas")
               data_arq.Recordset("iva") = data_emi.Recordset("iva")
               data_arq.Recordset("deudas") = data_emi.Recordset("deudas")
               data_arq.Recordset("servi") = 0
               data_arq.Recordset.Update
            End If
            data_emi.Recordset.MoveNext
         Loop
      End If
      frm_carfaccnv.MousePointer = 0
      MsgBox "Proceso terminado"
   Else
      data_fac.RecordSource = "Select * from linmmdd where fecha >=#" & Format(mfecd.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mfech.Text, "yyyy/mm/dd") & "# and base in(101,102) and grupo is not null and grupo not in (110,111,0) and tot_lin >" & 0 & " order by fecha"
      data_fac.Refresh
      b_proc.Enabled = False
      If data_fac.Recordset.RecordCount > 0 Then
         frm_carfaccnv.MousePointer = 11
         data_fac.Recordset.MoveFirst
         Do While Not data_fac.Recordset.EOF
            data_cab.RecordSource = "Select * from clirespl where cl_numero =" & data_fac.Recordset("factura")
            data_cab.Refresh
            If data_cab.Recordset.RecordCount > 0 Then
               If data_cab.Recordset("cl_tipocli") = 112 Or _
                  data_cab.Recordset("cl_tipocli") = 102 Then
               Else
                  data_arq.Recordset.AddNew
                  data_arq.Recordset("matricula") = data_fac.Recordset("cod_cli")
                  data_arq.Recordset("nombre") = data_fac.Recordset("nom_cli")
                  data_arq.Recordset("mes") = data_fac.Recordset("mes_paga")
                  data_arq.Recordset("ano") = data_fac.Recordset("ano_paga")
                  data_arq.Recordset("color") = "M"
                  data_arq.Recordset("cat") = data_fac.Recordset("convenio")
                  data_arq.Recordset("nomcat") = data_fac.Recordset("convenio")
                  data_arq.Recordset("arqueo") = "E"
                  data_arq.Recordset("importe") = data_fac.Recordset("tot_lin")
                  data_arq.Recordset("fecha") = data_fac.Recordset("fecha")
                  data_arq.Recordset("nrorec") = data_fac.Recordset("factura")
                  data_arq.Recordset("usuar") = WElusuario
                  data_arq.Recordset("moneda") = 1
                  data_arq.Recordset("cob") = data_fac.Recordset("grupo")
                  data_cob.RecordSource = "Select * from cobrador where cb_numero =" & data_fac.Recordset("grupo")
                  data_cob.Refresh
                  If data_cob.Recordset.RecordCount > 0 Then
                     data_arq.Recordset("nomcob") = data_cob.Recordset("cb_nombre")
                  Else
                     data_arq.Recordset("nomcob") = "SC"
                  End If
                  data_arq.Recordset("codzon") = 100
                  data_arq.Recordset("codpro") = 1
                  data_arq.Recordset("codsup") = 1
                  data_arq.Recordset("tiquet") = 0
                  data_arq.Recordset("total") = data_fac.Recordset("tot_lin") + data_fac.Recordset("valor_iva")
                  data_arq.Recordset("varia") = 0
                  data_arq.Recordset("iva") = data_fac.Recordset("valor_iva")
                  data_arq.Recordset("deudas") = 0
                  data_arq.Recordset("servi") = 0
                  data_arq.Recordset.Update
                  CantFact = CantFact + 1
               End If
            End If
            data_fac.Recordset.MoveNext
         Loop
         frm_carfaccnv.MousePointer = 0
         MsgBox "Proceso terminado. Se cargaron " & Trim(str(CantFact))
      End If
   End If
End If

   
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_arq.DatabaseName = App.Path & "\sapp.mdb"
data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_arq.RecordSource = "arqueo"
data_arq.Refresh

data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_cab.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_notascr.DatabaseName = App.path & "\notascrarq.mdb"

data_fac.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
