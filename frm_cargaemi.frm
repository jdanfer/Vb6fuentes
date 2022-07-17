VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_cargaemi 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar emisión a las deudas"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "frm_cargaemi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6195
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cargar verificando si existe"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Data data_deu 
      Caption         =   "data_deu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txt_ano 
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
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txt_mes 
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
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Terminar"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Procesar..."
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
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Mes y año de emisión:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frm_cargaemi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xqueemi As String

If Check1.value = 1 Then
   Command3_Click
Else
    frm_cargaemi.MousePointer = 11
    Command1.Enabled = False
    Command2.Enabled = False
    If txt_mes.Text > 9 Then
       Xqueemi = "EMI" & Trim(txt_mes.Text) & Mid(Trim(txt_ano.Text), 3, 2)
    Else
       Xqueemi = "EMI0" & Trim(txt_mes.Text) & Mid(Trim(txt_ano.Text), 3, 2)
    End If
    
    data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_emi.RecordSource = "Select * from " & Trim(Xqueemi) & " order by cliente"
    data_emi.Refresh
    data_deu.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_deu.RecordSource = "deudas"
    data_deu.Refresh
    
    If data_emi.Recordset.RecordCount > 0 Then
       data_emi.Recordset.MoveLast
       pb1.Max = data_emi.Recordset.RecordCount
       pb1.value = 0
       data_emi.Recordset.MoveFirst
       Do While Not data_emi.Recordset.EOF
            data_deu.Recordset.AddNew
            data_deu.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
            data_deu.Recordset("nom_cnv") = Mid(data_emi.Recordset("nom_cnv"), 1, 20)
            data_deu.Recordset("tipocta") = "CC"
            data_deu.Recordset("cliente") = data_emi.Recordset("cliente")
            data_deu.Recordset("nombre") = data_emi.Recordset("apellidos")
            data_deu.Recordset("fecha") = data_emi.Recordset("fecha")
            data_deu.Recordset("tipodoc") = "FAC"
            If IsNull(data_emi.Recordset("documento")) = False Then
               data_deu.Recordset("documento") = data_emi.Recordset("documento")
            Else
               data_deu.Recordset("documento") = 0
            End If
            If IsNull(data_emi.Recordset("importe")) = False Then
               data_deu.Recordset("importe") = data_emi.Recordset("importe")
            Else
               data_deu.Recordset("importe") = 0
            End If
            data_deu.Recordset("moneda") = 1
            data_deu.Recordset("origen") = "EMISION..." & Trim(Str(data_emi.Recordset("mes"))) & "/" & Trim(Str(data_emi.Recordset("ano")))
            If IsNull(data_emi.Recordset("nro_vende")) = False Then
               data_deu.Recordset("nro_vende") = data_emi.Recordset("nro_vende")
            Else
               data_deu.Recordset("nro_vende") = 0
            End If
            If IsNull(data_emi.Recordset("grupo")) = False Then
               data_deu.Recordset("grupo") = data_emi.Recordset("grupo")
            Else
               data_deu.Recordset("grupo") = 0
            End If
            data_deu.Recordset("saldo_cc") = 0
            data_deu.Recordset("mes") = data_emi.Recordset("mes")
            data_deu.Recordset("ano") = data_emi.Recordset("ano")
            If IsNull(data_emi.Recordset("nro_cobr")) = False Then
               data_deu.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
            Else
               data_deu.Recordset("nro_cobr") = 0
            End If
            If IsNull(data_emi.Recordset("nom_cobr")) = False Then
               data_deu.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
            Else
               data_deu.Recordset("nom_cobr") = ""
            End If
            data_deu.Recordset("estado_cta") = 1
            If IsNull(data_emi.Recordset("tiquet")) = False Then
               data_deu.Recordset("tiquet") = data_emi.Recordset("tiquet")
            Else
               data_deu.Recordset("tiquet") = 0
            End If
            If IsNull(data_emi.Recordset("deudas")) = False Then
               data_deu.Recordset("deudas") = data_emi.Recordset("deudas")
            Else
               data_deu.Recordset("deudas") = 0
            End If
            If IsNull(data_emi.Recordset("total")) = False Then
               data_deu.Recordset("total") = data_emi.Recordset("total")
            Else
               data_deu.Recordset("total") = 0
            End If
            If IsNull(data_emi.Recordset("servi")) = False Then
               data_deu.Recordset("servi") = data_emi.Recordset("servi")
            Else
               data_deu.Recordset("servi") = 0
            End If
            If IsNull(data_emi.Recordset("iva")) = False Then
               data_deu.Recordset("iva") = data_emi.Recordset("iva")
            Else
               data_deu.Recordset("iva") = 0
            End If
            data_deu.Recordset("nro_superv") = 50
            data_deu.Recordset.Update
            data_emi.Recordset.MoveNext
            pb1.value = pb1.value + 1
       Loop
       DoEvents
    End If
End If
frm_cargaemi.MousePointer = 0
Command1.Enabled = True
Command2.Enabled = True
MsgBox "Proceso de carga de emisión finalizado...", vbInformation, "Deudas"
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim Xqueemi2 As String
frm_cargaemi.MousePointer = 11
Command1.Enabled = False
Command2.Enabled = False
If txt_mes.Text > 9 Then
   Xqueemi2 = "EMI" & Trim(txt_mes.Text) & Mid(Trim(txt_ano.Text), 3, 2)
Else
   Xqueemi2 = "EMI0" & Trim(txt_mes.Text) & Mid(Trim(txt_ano.Text), 3, 2)
End If

data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_emi.RecordSource = "Select * from " & Trim(Xqueemi2) & " order by cliente"
data_emi.Refresh
data_deu.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_deu.RecordSource = "deudas"
'data_deu.Refresh

If data_emi.Recordset.RecordCount > 0 Then
   data_emi.Recordset.MoveLast
   pb1.Max = data_emi.Recordset.RecordCount
   pb1.value = 0
   data_emi.Recordset.MoveFirst
   Do While Not data_emi.Recordset.EOF
      If data_emi.Recordset("nro_cobr") = 615 Or data_emi.Recordset("nro_cobr") = 616 Or _
         data_emi.Recordset("nro_cobr") = 636 Or _
         data_emi.Recordset("nro_cobr") = 635 Or _
         data_emi.Recordset("nro_cobr") = 602 Or _
         data_emi.Recordset("nro_cobr") = 653 Or _
         data_emi.Recordset("nro_cobr") = 672 Or _
         data_emi.Recordset("nro_cobr") = 113 Or _
         data_emi.Recordset("nro_cobr") = 685 Or _
         data_emi.Recordset("nro_cobr") = 10 Or _
         data_emi.Recordset("nro_cobr") = 1 Or _
         data_emi.Recordset("nro_cobr") = 679 Or _
         data_emi.Recordset("nro_cobr") = 685 Or _
         data_emi.Recordset("nro_cobr") = 512 Or _
         data_emi.Recordset("nro_cobr") = 201 Or _
         data_emi.Recordset("nro_cobr") = 208 Or _
         data_emi.Recordset("nro_cobr") = 209 Then
      
        data_deu.RecordSource = "Select * from deudas where cliente =" & data_emi.Recordset("cliente") & " and mes =" & data_emi.Recordset("mes") & " and ano =" & data_emi.Recordset("ano")
        data_deu.Refresh
        If data_deu.Recordset.RecordCount > 0 Then
        Else
           data_deu.Recordset.AddNew
           data_deu.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
           data_deu.Recordset("nom_cnv") = Mid(data_emi.Recordset("nom_cnv"), 1, 20)
           data_deu.Recordset("tipocta") = "CC"
           data_deu.Recordset("cliente") = data_emi.Recordset("cliente")
           data_deu.Recordset("nombre") = data_emi.Recordset("apellidos")
           data_deu.Recordset("fecha") = data_emi.Recordset("fecha")
           data_deu.Recordset("tipodoc") = "FAC"
           If IsNull(data_emi.Recordset("documento")) = False Then
              data_deu.Recordset("documento") = data_emi.Recordset("documento")
           Else
              data_deu.Recordset("documento") = 0
           End If
           If IsNull(data_emi.Recordset("importe")) = False Then
              data_deu.Recordset("importe") = data_emi.Recordset("importe")
           Else
              data_deu.Recordset("importe") = 0
           End If
           data_deu.Recordset("moneda") = 1
           data_deu.Recordset("origen") = "EMISION..." & Trim(Str(data_emi.Recordset("mes"))) & "/" & Trim(Str(data_emi.Recordset("ano")))
           If IsNull(data_emi.Recordset("nro_vende")) = False Then
              data_deu.Recordset("nro_vende") = data_emi.Recordset("nro_vende")
           Else
              data_deu.Recordset("nro_vende") = 0
           End If
           If IsNull(data_emi.Recordset("grupo")) = False Then
              data_deu.Recordset("grupo") = data_emi.Recordset("grupo")
           Else
              data_deu.Recordset("grupo") = 0
           End If
           data_deu.Recordset("saldo_cc") = 0
           data_deu.Recordset("mes") = data_emi.Recordset("mes")
           data_deu.Recordset("ano") = data_emi.Recordset("ano")
           If IsNull(data_emi.Recordset("nro_cobr")) = False Then
              data_deu.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
           Else
              data_deu.Recordset("nro_cobr") = 0
           End If
           If IsNull(data_emi.Recordset("nom_cobr")) = False Then
              data_deu.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
           Else
              data_deu.Recordset("nom_cobr") = ""
           End If
           data_deu.Recordset("estado_cta") = 1
           If IsNull(data_emi.Recordset("tiquet")) = False Then
              data_deu.Recordset("tiquet") = data_emi.Recordset("tiquet")
           Else
              data_deu.Recordset("tiquet") = 0
           End If
           If IsNull(data_emi.Recordset("deudas")) = False Then
              data_deu.Recordset("deudas") = data_emi.Recordset("deudas")
           Else
              data_deu.Recordset("deudas") = 0
           End If
           If IsNull(data_emi.Recordset("total")) = False Then
              data_deu.Recordset("total") = data_emi.Recordset("total")
           Else
              data_deu.Recordset("total") = 0
           End If
           If IsNull(data_emi.Recordset("servi")) = False Then
              data_deu.Recordset("servi") = data_emi.Recordset("servi")
           Else
              data_deu.Recordset("servi") = 0
           End If
           If IsNull(data_emi.Recordset("iva")) = False Then
              data_deu.Recordset("iva") = data_emi.Recordset("iva")
           Else
              data_deu.Recordset("iva") = 0
           End If
           data_deu.Recordset("nro_superv") = 50
           data_deu.Recordset.Update
        End If
      End If
      data_emi.Recordset.MoveNext
      pb1.value = pb1.value + 1
   Loop
   DoEvents
End If
'frm_cargaemi.MousePointer = 0
'Command1.Enabled = True
'Command2.Enabled = True
'MsgBox "Proceso de carga de emisión finalizado...", vbInformation, "Deudas"
'Unload Me

End Sub

Private Sub Form_Load()
txt_mes.Text = Month(Date)
txt_ano.Text = Year(Date)
If WElusuario = "JFERNAN" Then
   Check1.Enabled = True
Else
   Check1.Enabled = False
End If

End Sub
