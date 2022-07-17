VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_procdeu 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso de pendientes"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "frm_procdeu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data data_lindeu 
      Caption         =   "data_lindeu"
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
      Top             =   2280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
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
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
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
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_arqueo 
      Caption         =   "data_arqueo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
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
      RecordSource    =   "CLIENTES"
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton btn_cancela 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Picture         =   "frm_procdeu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton btn_proc 
      BackColor       =   &H00C0E0FF&
      Caption         =   "PROCESAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      Picture         =   "frm_procdeu.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txt_ano 
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
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txt_mes 
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
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Este proceso demora aproximadamente 2 horas."
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
      Left            =   480
      TabIndex        =   7
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MES Y AÑO:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frm_procdeu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cancela_Click()
frm_procdeu.Hide

End Sub

Private Sub btn_proc_Click()
Dim Nomemi As String
Dim SiNo As String

'data_cli.DatabaseName = App.Path & "\sapp.mdb"
'data_deu.DatabaseName = App.Path & "\deudas.mdb"
data_arqueo.Connect = "ODBC;DSN=sapp;"
'data_lin.DatabaseName = App.Path & "\sapp.mdb"
'data_arqueo.RecordSource = "Select * from arqueo where arqueo ='" & "P" & "' or arqueo ='" & "B" & "'"
'data_arqueo.Refresh
'data_emi.DatabaseName = ""
data_emi.Connect = "ODBC;DSN=sapp;"
Data1.Connect = "ODBC;DSN=sapp;"
'Data1.RecordSource = "deudas"
'Data1.Refresh
'Data2.Connect = "ODBC;DSN=sapp;"
'xl 1049

frm_procdeu.MousePointer = 11

If txt_mes.Text <> "" Then
   If txt_ano.Text <> "" Then
      If Val(txt_mes.Text) > 9 Then
         Nomemi = "EMI" & Trim(txt_mes.Text) & Trim(Mid(txt_ano.Text, 3, 2))
      Else
         Nomemi = "EMI0" & Trim(txt_mes.Text) & Trim(Mid(txt_ano.Text, 3, 2))
      End If
   End If
End If

If txt_mes.Text <> "" Then
   If txt_ano.Text <> "" Then
      Data1.RecordSource = "Select * from deudas where mes =" & txt_mes.Text & " and ano =" & txt_ano.Text
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.MoveLast
         If Data1.Recordset.RecordCount > 3000 Then
            MsgBox "Emisión ya está cargada"
         Else
            data_emi.RecordSource = "Select * from " & Nomemi
            data_emi.Refresh
            If data_emi.Recordset.RecordCount > 0 Then
               data_emi.Recordset.MoveFirst
               Do While Not data_emi.Recordset.EOF
                  Data1.Recordset.AddNew
                  Data1.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
                  Data1.Recordset("nom_cnv") = Mid(data_emi.Recordset("nom_cnv"), 1, 20)
                  Data1.Recordset("tipocta") = "CC"
                  Data1.Recordset("cliente") = data_emi.Recordset("cliente")
                  Data1.Recordset("nombre") = data_emi.Recordset("apellidos")
                  Data1.Recordset("fecha") = data_emi.Recordset("fecha")
                  Data1.Recordset("tipodoc") = "FAC"
                  If IsNull(data_emi.Recordset("documento")) = False Then
                     Data1.Recordset("documento") = data_emi.Recordset("documento")
                  Else
                     Data1.Recordset("documento") = 0
                  End If
                  If IsNull(data_emi.Recordset("importe")) = False Then
                     Data1.Recordset("importe") = data_emi.Recordset("importe")
                  Else
                     Data1.Recordset("importe") = 0
                  End If
                  Data1.Recordset("moneda") = 1
                  Data1.Recordset("origen") = "EMISION..." & Trim(Str(data_emi.Recordset("mes"))) & "/" & Trim(Str(data_emi.Recordset("ano")))
                  If IsNull(data_emi.Recordset("nro_vende")) = False Then
                     Data1.Recordset("nro_vende") = data_emi.Recordset("nro_vende")
                  Else
                     Data1.Recordset("nro_vende") = 0
                  End If
                  If IsNull(data_emi.Recordset("grupo")) = False Then
                     Data1.Recordset("grupo") = data_emi.Recordset("grupo")
                  Else
                     Data1.Recordset("grupo") = 0
                  End If
                  Data1.Recordset("saldo_cc") = 0
                  Data1.Recordset("mes") = data_emi.Recordset("mes")
                  Data1.Recordset("ano") = data_emi.Recordset("ano")
                  If IsNull(data_emi.Recordset("nro_cobr")) = False Then
                     Data1.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
                  Else
                     Data1.Recordset("nro_cobr") = 0
                  End If
                  If IsNull(data_emi.Recordset("nom_cobr")) = False Then
                     Data1.Recordset("nom_cobr") = data_emi.Recordset("nom_cobr")
                  Else
                     Data1.Recordset("nom_cobr") = ""
                  End If
                  Data1.Recordset("estado_cta") = 1
                  If IsNull(data_emi.Recordset("tiquet")) = False Then
                     Data1.Recordset("tiquet") = data_emi.Recordset("tiquet")
                  Else
                     Data1.Recordset("tiquet") = 0
                  End If
                  If IsNull(data_emi.Recordset("deudas")) = False Then
                     Data1.Recordset("deudas") = data_emi.Recordset("deudas")
                  Else
                     Data1.Recordset("deudas") = 0
                  End If
                  If IsNull(data_emi.Recordset("total")) = False Then
                     Data1.Recordset("total") = data_emi.Recordset("total")
                  Else
                     Data1.Recordset("total") = 0
                  End If
                  If IsNull(data_emi.Recordset("servi")) = False Then
                     Data1.Recordset("servi") = data_emi.Recordset("servi")
                  Else
                     Data1.Recordset("servi") = 0
                  End If
                  If IsNull(data_emi.Recordset("iva")) = False Then
                     Data1.Recordset("iva") = data_emi.Recordset("iva")
                  Else
                     Data1.Recordset("iva") = 0
                  End If
                  Data1.Recordset.Update
                  data_emi.Recordset.MoveNext
               Loop
            End If
         End If
      End If
   End If
End If
MsgBox "Carga de emisión terminado"

Dim Xorigen As String

'busca si no está el pendiente o baja para agregar
'acá-----
'data_arqueo.RecordSource = "Select * from arq0616 where arqueo in ('P','B')"
'data_arqueo.Refresh
'If data_arqueo.Recordset.RecordCount > 0 Then
'   data_arqueo.Recordset.MoveFirst
'   Do While Not data_arqueo.Recordset.EOF
'      Data1.RecordSource = "Select * from deudas where cliente =" & data_arqueo.Recordset("matricula") & " and mes =" & data_arqueo.Recordset("mes") & " and ano =" & data_arqueo.Recordset("ano")
'      Data1.Refresh
'      If Data1.Recordset.RecordCount > 0 Then
'      Else
'         Data1.Recordset.AddNew
'         Data1.Recordset("cod_cnv") = data_arqueo.Recordset("cat")
'         Data1.Recordset("nom_cnv") = Mid(data_arqueo.Recordset("nomcat"), 1, 20)
'         Data1.Recordset("tipocta") = "CC"
'         Data1.Recordset("cliente") = data_arqueo.Recordset("matricula")
'         Data1.Recordset("nombre") = data_arqueo.Recordset("nombre")
'         Data1.Recordset("fecha") = data_arqueo.Recordset("fecha")
'         Data1.Recordset("tipodoc") = "FAC"
'         Data1.Recordset("documento") = data_arqueo.Recordset("nrorec")
'         Data1.Recordset("importe") = data_arqueo.Recordset("importe")
'         Data1.Recordset("moneda") = 1
'         Xorigen = "Pendiente..." & Trim(Str(data_arqueo.Recordset("mes"))) & "/" & Trim(Str(data_arqueo.Recordset("ano")))
'         Data1.Recordset("origen") = Xorigen
 '        Data1.Recordset("nro_vende") = data_arqueo.Recordset("codpro")
'         Data1.Recordset("grupo") = data_arqueo.Recordset("codzon")
'         Data1.Recordset("saldo_cc") = 0
'         Data1.Recordset("mes") = data_arqueo.Recordset("mes")
'         Data1.Recordset("ano") = data_arqueo.Recordset("ano")
'         Data1.Recordset("nro_cobr") = data_arqueo.Recordset("cob")
'         Data1.Recordset("nom_cobr") = data_arqueo.Recordset("nomcob")
'         If data_arqueo.Recordset("arqueo") = "B" Then
'            Data1.Recordset("estado_cta") = 2
'         Else
'            Data1.Recordset("estado_cta") = 1
'         End If
'         Data1.Recordset("tiquet") = data_arqueo.Recordset("tiquet")
'         Data1.Recordset("deudas") = data_arqueo.Recordset("deudas")
'         Data1.Recordset("total") = data_arqueo.Recordset("total")
'         Data1.Recordset("servi") = data_arqueo.Recordset("servi")
'         Data1.Recordset("iva") = data_arqueo.Recordset("iva")
'         Data1.Recordset.Update
'      End If
'      data_arqueo.Recordset.MoveNext
'   Loop
'End If
''-----

SiNo = MsgBox("Terminado proceso de agregar pendientes/bajas, CONTINUA?", vbCritical + vbYesNo)
If SiNo = vbYes Then
Else
   End
End If
SiNo = ""

Data1.RecordSource = "Select * from deudas where estado_cta =" & 1 & " and tipodoc <>'" & "CRE" & "' order by cliente,ano,documento DESC"
Data1.Refresh
'busca en tabla arqueo si está pendiente o baja

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      If IsNull(Data1.Recordset("fecha_pago")) = False Then
         Data1.Recordset.Delete
      Else
        If Data1.Recordset("nro_cobr") = 616 Or _
           Data1.Recordset("nro_cobr") = 615 Or _
           Data1.Recordset("nro_cobr") = 636 Or _
           Data1.Recordset("nro_cobr") = 635 Or _
           Data1.Recordset("nro_cobr") = 602 Or _
           Data1.Recordset("nro_cobr") = 653 Or _
           Data1.Recordset("nro_cobr") = 672 Or _
           Data1.Recordset("nro_cobr") = 113 Or _
           Data1.Recordset("nro_cobr") = 685 Or _
           Data1.Recordset("nro_cobr") = 10 Or _
           Data1.Recordset("nro_cobr") = 1 Or _
           Data1.Recordset("nro_cobr") = 679 Or _
           Data1.Recordset("nro_cobr") = 685 Or _
           Data1.Recordset("nro_cobr") = 512 Or _
           Data1.Recordset("nro_cobr") = 201 Or _
           Data1.Recordset("nro_cobr") = 208 Or _
           Data1.Recordset("nro_cobr") = 209 Then
           If Int(Data1.Recordset("mes")) = Int(Val(txt_mes.Text)) And Int(Data1.Recordset("ano")) = Int(Val(txt_ano.Text)) Then
           Else
              Xorigen = "Pendiente..." & Trim(Str(Data1.Recordset("mes"))) & "/" & Trim(Str(Data1.Recordset("ano")))
              If Trim(Data1.Recordset("origen")) <> Trim(Xorigen) Then
                 Data1.Recordset.Edit
                 Data1.Recordset("origen") = Xorigen
                 Data1.Recordset.Update
              End If
           End If
        Else
           If Int(Data1.Recordset("mes")) = Int(Val(txt_mes.Text)) And Int(Data1.Recordset("ano")) = Int(Val(txt_ano.Text)) Then
           Else
              data_arqueo.RecordSource = "Select * from arq0616 where arqueo ='" & "P" & "' and matricula =" & Data1.Recordset("cliente") & " and mes =" & Data1.Recordset("mes") & " and ano =" & Data1.Recordset("ano")
              data_arqueo.Refresh
              If data_arqueo.Recordset.RecordCount > 0 Then
                 Xorigen = "Pendiente..." & Trim(Str(data_arqueo.Recordset("mes"))) & "/" & Trim(Str(data_arqueo.Recordset("ano")))
                 If Trim(Data1.Recordset("origen")) <> Trim(Xorigen) Then
                    Data1.Recordset.Edit
                    Data1.Recordset("origen") = Xorigen
                    Data1.Recordset.Update
                 End If
              Else
                 data_arqueo.RecordSource = "Select * from arq0616 where arqueo ='" & "B" & "' and matricula =" & Data1.Recordset("cliente") & " and mes =" & Data1.Recordset("mes") & " and ano =" & Data1.Recordset("ano")
                 data_arqueo.Refresh
                 If data_arqueo.Recordset.RecordCount > 0 Then
                    Xorigen = "Pendiente..." & Trim(Str(data_arqueo.Recordset("mes"))) & "/" & Trim(Str(data_arqueo.Recordset("ano")))
                    If Data1.Recordset("estado_cta") <> 2 Then
                       Data1.Recordset.Edit
                       Data1.Recordset("estado_cta") = 2
                       Data1.Recordset("origen") = Xorigen
                       Data1.Recordset.Update
                    End If
                 Else
                    Data1.Recordset.Delete
                 End If
              End If
           End If
        End If
      End If
      Data1.Recordset.MoveNext
   Loop
   Data1.Refresh
End If
MsgBox "Proceso de pendientes y Bajas, terminado"

Data1.RecordSource = "Select * from deudas where estado_cta =" & 1 & " and tipodoc ='" & "CRE" & "' and documento =" & 0
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data1.Recordset.Delete
      Data1.Recordset.MoveNext
   Loop
End If

SiNo = MsgBox("Terminado proceso de pendientes, CONTINUA?", vbCritical + vbYesNo)
If SiNo = vbYes Then
Else
   End
End If
SiNo = ""


DoEvents

'cobros de cuotas en base
'---
'If Data1.Recordset.RecordCount > 0 Then
'   Data1.Recordset.MoveFirst
'   Do While Not Data1.Recordset.EOF
'      data_arqueo.RecordSource = "Select * from linmmdd where cod_cli =" & Data1.Recordset("cliente") & " and mes_paga =" & Data1.Recordset("mes") & " and ano_paga =" & Data1.Recordset("ano") & " and cod_prod =" & 999
'      data_arqueo.Refresh
'      If data_arqueo.Recordset.RecordCount > 0 Then
'         Data1.Recordset.Delete
'      End If
'      Data1.Recordset.MoveNext
'   Loop
'End If
'---
'DoEvents

'SiNo = MsgBox("Terminado proceso de pagos en base, CONTINUA?", vbCritical + vbYesNo)
'If SiNo = vbYes Then
'Else
'   End
'End If
'SiNo = ""


frm_procdeu.MousePointer = 0
MsgBox "Comienza generación de Saldos"
Dim Xtotal As Double
Dim Xcant As Integer
Dim Xmat As Long
Dim Xmes, Xano As Integer
Dim Xfdlin As Date
Xfdlin = Date - 65

Xtotal = 0
frm_procdeu.MousePointer = 11
'inicializa deudas en cero de clientes activos
'data_cli.RecordSource = "Select * from clientes where estado =" & 1
'data_cli.Refresh
'If data_cli.Recordset.RecordCount > 0 Then
'   data_cli.Recordset.MoveFirst
'   Do While Not data_cli.Recordset.EOF
'      If IsNull(data_cli.Recordset("saldo_cc")) = False Then
'         If Format(data_cli.Recordset("saldo_cc"), "Standard") <> Format(Xtotal, "Standard") Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("saldo_cc") = 0
'            data_cli.Recordset.Update
'         End If
'      End If
'      If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
'         If Int(data_cli.Recordset("cl_ultmesp")) <> 0 Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("cl_ultmesp") = 0
'            data_cli.Recordset.Update
'         End If
'      End If
'      If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
'         If Int(data_cli.Recordset("cl_ultanop")) <> 0 Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("cl_ultanop") = 0
'            data_cli.Recordset.Update
'         End If
'      End If
'      If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
'         If Int(data_cli.Recordset("cl_atrasoa")) <> 0 Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("cl_atrasoa") = 0
'            data_cli.Recordset.Update
'         End If
'      End If
'      data_cli.Recordset.MoveNext
'   Loop

'End If

'DoEvents

'MsgBox "Inicialización terminada"


Dim Xelcli As Long
Xelcli = 0
'----
'data_arqueo.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xfdlin, "yyyy/mm/dd") & "# and cod_prod =" & 999 & " order by cod_cli,ano_paga,mes_paga"
'data_arqueo.Refresh
'If data_arqueo.Recordset.RecordCount > 0 Then
'   data_arqueo.Recordset.MoveFirst
'   Xelcli = data_arqueo.Recordset("cod_cli")
'   Do While Not data_arqueo.Recordset.EOF
'      If Int(Xelcli) = Int(data_arqueo.Recordset("cod_cli")) Then
'         Xelcli = data_arqueo.Recordset("cod_cli")
'         data_arqueo.Recordset.MoveNext
'      Else
'         data_arqueo.Recordset.MovePrevious
'         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_arqueo.Recordset("cod_cli")
'         data_cli.Refresh
'         If data_cli.Recordset.RecordCount > 0 Then
'            If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
'               If Int(data_cli.Recordset("cl_ultmesp")) <> Int(data_arqueo.Recordset("mes_paga")) Then
'                  data_cli.Recordset.Edit
'                  data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes_paga")
'                  data_cli.Recordset.Update
'               End If
'            Else
'               data_cli.Recordset.Edit
'               data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes_paga")
'               data_cli.Recordset.Update
'            End If
'            If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
'               If Int(data_cli.Recordset("cl_ultanop")) <> Int(data_arqueo.Recordset("ano_paga")) Then
'                  data_cli.Recordset.Edit
'                  data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano_paga")
'                  data_cli.Recordset.Update
'               End If
'            Else
'               data_cli.Recordset.Edit
'               data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano_paga")
'               data_cli.Recordset.Update
'            End If
'         End If
'         data_arqueo.Recordset.MoveNext
'         Xelcli = data_arqueo.Recordset("cod_cli")
'      End If
'   Loop
'   data_arqueo.Recordset.MovePrevious
'   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_arqueo.Recordset("cod_cli")
'   data_cli.Refresh
'   If data_cli.Recordset.RecordCount > 0 Then
'      If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
'         If Int(data_cli.Recordset("cl_ultmesp")) <> Int(data_arqueo.Recordset("mes_paga")) Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes_paga")
'            data_cli.Recordset.Update
'         End If
'      Else
'         data_cli.Recordset.Edit
'         data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes_paga")
'         data_cli.Recordset.Update
'      End If
'      If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
''         If Int(data_cli.Recordset("cl_ultanop")) <> Int(data_arqueo.Recordset("ano_paga")) Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano_paga")
'            data_cli.Recordset.Update
'         End If
'      Else
'         data_cli.Recordset.Edit
'         data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano_paga")
'         data_cli.Recordset.Update
'      End If
'   End If
'End If

'DoEvents
'MsgBox "Actualización de pagos en base terminado"

Xcant = 0
Xtotal = 0
Xmes = 0
Xano = 0
Data1.RecordSource = "Select * from deudas where estado_cta =" & 1 & " and fecha_pago is null order by cliente,ano,documento DESC"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Xmat = Data1.Recordset("cliente")
   Do While Not Data1.Recordset.EOF
      If Data1.Recordset("cliente") = Xmat Then
         If Data1.Recordset("tipodoc") = "CRE" Then
            Xtotal = Xtotal + Data1.Recordset("total")
         Else
            Xcant = Xcant + 1
            Xtotal = Xtotal + Data1.Recordset("total")
            Xmes = Data1.Recordset("mes")
            Xano = Data1.Recordset("ano")
         End If
      Else
         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            If Xmes = 1 Then
               Xmes = 12
               Xano = Xano - 1
            Else
               Xmes = Xmes - 1
            End If
            If IsNull(data_cli.Recordset("saldo_cc")) = False Then
               If Format(data_cli.Recordset("saldo_cc"), "Standard") <> Format(Xtotal, "Standard") Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("saldo_cc") = Xtotal
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("saldo_cc") = Xtotal
               data_cli.Recordset.Update
            End If
            
            If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
               If Int(data_cli.Recordset("cl_atrasoa")) <> Int(Xcant) Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_atrasoa") = Xcant
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_atrasoa") = Xcant
               data_cli.Recordset.Update
            End If
            
            If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
               If Int(data_cli.Recordset("cl_ultmesp")) <> Int(Xmes) Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_ultmesp") = Xmes
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_ultmesp") = Xmes
               data_cli.Recordset.Update
            End If
            
            If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
               If Int(data_cli.Recordset("cl_ultanop")) <> Int(Xano) Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_ultanop") = Xano
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_ultanop") = Xano
               data_cli.Recordset.Update
            End If
         End If
         Xcant = 1
         Xmes = Data1.Recordset("mes")
         Xano = Data1.Recordset("ano")
         Xtotal = Data1.Recordset("total")
      End If
      Xmat = Data1.Recordset("cliente")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      If Xmes = 1 Then
         Xmes = 12
         Xano = Xano - 1
      Else
         Xmes = Xmes - 1
      End If
      If IsNull(data_cli.Recordset("saldo_cc")) = False Then
         If Format(data_cli.Recordset("saldo_cc"), "Standard") <> Format(Xtotal, "Standard") Then
            data_cli.Recordset.Edit
            data_cli.Recordset("saldo_cc") = Xtotal
            data_cli.Recordset.Update
         End If
      Else
         data_cli.Recordset.Edit
         data_cli.Recordset("saldo_cc") = Xtotal
         data_cli.Recordset.Update
      End If
       
      If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
         If Int(data_cli.Recordset("cl_atrasoa")) <> Int(Xcant) Then
            data_cli.Recordset.Edit
            data_cli.Recordset("cl_atrasoa") = Xcant
            data_cli.Recordset.Update
         End If
      Else
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_atrasoa") = Xcant
         data_cli.Recordset.Update
      End If
      
      If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
         If Int(data_cli.Recordset("cl_ultmesp")) <> Int(Xmes) Then
            data_cli.Recordset.Edit
            data_cli.Recordset("cl_ultmesp") = Xmes
            data_cli.Recordset.Update
         End If
      Else
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_ultmesp") = Xmes
         data_cli.Recordset.Update
      End If
       
      If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
         If Int(data_cli.Recordset("cl_ultanop")) <> Int(Xano) Then
            data_cli.Recordset.Edit
            data_cli.Recordset("cl_ultanop") = Xano
            data_cli.Recordset.Update
         End If
      Else
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_ultanop") = Xano
         data_cli.Recordset.Update
      End If
   End If
End If
frm_procdeu.MousePointer = 0
MsgBox "Proceso terminado"


End Sub

Private Sub Command3_Click()
data_arqueo.Connect = "ODBC;DSN=sapp;"
'data_lin.DatabaseName = App.Path & "\sapp.mdb"
'data_arqueo.RecordSource = "Select * from arqueo where arqueo ='" & "P" & "' or arqueo ='" & "B" & "'"
'data_arqueo.Refresh
'data_emi.DatabaseName = ""
data_emi.Connect = "ODBC;DSN=sapp;"
Data1.Connect = "ODBC;DSN=sapp;"


Data1.RecordSource = "Select * from deudas where estado_cta =" & 1 & " and tipodoc <>'" & "CRE" & "' order by cliente,ano,documento DESC"
Data1.Refresh
'busca en tabla arqueo si está pendiente o baja

'If Data1.Recordset.RecordCount > 0 Then
'   Data1.Recordset.MoveFirst
'   Do While Not Data1.Recordset.EOF
'      If IsNull(Data1.Recordset("fecha_pago")) = False Then
'         Data1.Recordset.Delete
'      Else
 '       If Data1.Recordset("nro_cobr") = 616 Or _
'           Data1.Recordset("nro_cobr") = 615 Or _
'           Data1.Recordset("nro_cobr") = 636 Or _
'           Data1.Recordset("nro_cobr") = 635 Or _
'           Data1.Recordset("nro_cobr") = 602 Or _
'           Data1.Recordset("nro_cobr") = 653 Or _
'           Data1.Recordset("nro_cobr") = 672 Or _
'           Data1.Recordset("nro_cobr") = 113 Or _
'           Data1.Recordset("nro_cobr") = 685 Or _
'           Data1.Recordset("nro_cobr") = 10 Or _
'           Data1.Recordset("nro_cobr") = 1 Or _
'           Data1.Recordset("nro_cobr") = 679 Or _
'           Data1.Recordset("nro_cobr") = 685 Or _
'           Data1.Recordset("nro_cobr") = 512 Or _
'           Data1.Recordset("nro_cobr") = 201 Or _
'           Data1.Recordset("nro_cobr") = 208 Or _
'           Data1.Recordset("nro_cobr") = 209 Then
'           If Int(Data1.Recordset("mes")) = Int(Val(txt_mes.Text)) And Int(Data1.Recordset("ano")) = Int(Val(txt_ano.Text)) Then
'           Else
'              Xorigen = "Pendiente..." & Trim(Str(Data1.Recordset("mes"))) & "/" & Trim(Str(Data1.Recordset("ano")))
'              If Trim(Data1.Recordset("origen")) <> Trim(Xorigen) Then
'                 Data1.Recordset.Edit
'                 Data1.Recordset("origen") = Xorigen
'                 Data1.Recordset.Update
'              End If
'           End If
'        Else
'           If Int(Data1.Recordset("mes")) = Int(Val(txt_mes.Text)) And Int(Data1.Recordset("ano")) = Int(Val(txt_ano.Text)) Then
'           Else
'              data_arqueo.RecordSource = "Select * from arqueo where arqueo ='" & "P" & "' and matricula =" & Data1.Recordset("cliente") & " and mes =" & Data1.Recordset("mes") & " and ano =" & Data1.Recordset("ano")
'              data_arqueo.Refresh
'              If data_arqueo.Recordset.RecordCount > 0 Then
'                 Xorigen = "Pendiente..." & Trim(Str(data_arqueo.Recordset("mes"))) & "/" & Trim(Str(data_arqueo.Recordset("ano")))
'                 If Trim(Data1.Recordset("origen")) <> Trim(Xorigen) Then
'                    Data1.Recordset.Edit
'                    Data1.Recordset("origen") = Xorigen
'                    Data1.Recordset.Update
'                 End If
'              Else
'                 data_arqueo.RecordSource = "Select * from arqueo where arqueo ='" & "B" & "' and matricula =" & Data1.Recordset("cliente") & " and mes =" & Data1.Recordset("mes") & " and ano =" & Data1.Recordset("ano")
'                 data_arqueo.Refresh
'                 If data_arqueo.Recordset.RecordCount > 0 Then
'                    Xorigen = "Pendiente..." & Trim(Str(data_arqueo.Recordset("mes"))) & "/" & Trim(Str(data_arqueo.Recordset("ano")))
'                    If Data1.Recordset("estado_cta") <> 2 Then
'                       Data1.Recordset.Edit
'                       Data1.Recordset("estado_cta") = 2
'                       Data1.Recordset("origen") = Xorigen
'                       Data1.Recordset.Update
'                    End If
'                 Else
'                    Data1.Recordset.Delete
'                 End If
'              End If
'           End If
'        End If
'      End If
'      Data1.Recordset.MoveNext
'   Loop
'   Data1.Refresh
'End If
'MsgBox "Proceso de pendientes y Bajas, terminado"

'SiNo = MsgBox("Terminado proceso de pendientes, CONTINUA?", vbCritical + vbYesNo)
'If SiNo = vbYes Then
'Else
'   End
'End If
'SiNo = ""


'DoEvents

'If Data1.Recordset.RecordCount > 0 Then
'   Data1.Recordset.MoveFirst
'   Do While Not Data1.Recordset.EOF
'      data_arqueo.RecordSource = "Select * from linmmdd where cod_cli =" & Data1.Recordset("cliente") & " and mes_paga =" & Data1.Recordset("mes") & " and ano_paga =" & Data1.Recordset("ano") & " and cod_prod =" & 999
'      data_arqueo.Refresh
'      If data_arqueo.Recordset.RecordCount > 0 Then
'         Data1.Recordset.Delete
'      End If
'      Data1.Recordset.MoveNext
'   Loop
'End If

'DoEvents

'SiNo = MsgBox("Terminado proceso de pagos en base, CONTINUA?", vbCritical + vbYesNo)
'If SiNo = vbYes Then
'Else
'   End
'End If
'SiNo = ""


'frm_procdeu.MousePointer = 0
'MsgBox "Comienza generación de Saldos"
'Dim Xtotal As Double
'Dim Xcant As Integer
'Dim Xmat As Long
'Dim Xmes, Xano As Integer
'Dim Xfdlin As Date
'Xfdlin = Date - 65

'Xtotal = 0
'frm_procdeu.MousePointer = 11

'Dim Xelcli As Long
'Xelcli = 0

'data_arqueo.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xfdlin, "yyyy/mm/dd") & "# and cod_prod =" & 999 & " order by cod_cli,ano_paga,mes_paga"
'data_arqueo.Refresh
'If data_arqueo.Recordset.RecordCount > 0 Then
'   data_arqueo.Recordset.MoveFirst
'   Xelcli = data_arqueo.Recordset("cod_cli")
'   Do While Not data_arqueo.Recordset.EOF
'      If Int(Xelcli) = Int(data_arqueo.Recordset("cod_cli")) Then
'         Xelcli = data_arqueo.Recordset("cod_cli")
'         data_arqueo.Recordset.MoveNext
'      Else
'         data_arqueo.Recordset.MovePrevious
'         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_arqueo.Recordset("cod_cli")
'         data_cli.Refresh
'         If data_cli.Recordset.RecordCount > 0 Then
'            If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
'               If Int(data_cli.Recordset("cl_ultmesp")) <> Int(data_arqueo.Recordset("mes_paga")) Then
'                  data_cli.Recordset.Edit
'                  data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes_paga")
'                  data_cli.Recordset.Update
'               End If
'            Else
'               data_cli.Recordset.Edit
'               data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes_paga")
'               data_cli.Recordset.Update
'            End If
'            If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
'               If Int(data_cli.Recordset("cl_ultanop")) <> Int(data_arqueo.Recordset("ano_paga")) Then
'                  data_cli.Recordset.Edit
'                  data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano_paga")
'                  data_cli.Recordset.Update
'               End If
'            Else
'               data_cli.Recordset.Edit
'               data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano_paga")
'               data_cli.Recordset.Update
'            End If
'         End If
'         data_arqueo.Recordset.MoveNext
'         Xelcli = data_arqueo.Recordset("cod_cli")
'      End If
'   Loop
'   data_arqueo.Recordset.MovePrevious
'   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_arqueo.Recordset("cod_cli")
'   data_cli.Refresh
'   If data_cli.Recordset.RecordCount > 0 Then
'      If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
'         If Int(data_cli.Recordset("cl_ultmesp")) <> Int(data_arqueo.Recordset("mes_paga")) Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes_paga")
'            data_cli.Recordset.Update
'         End If
'      Else
'         data_cli.Recordset.Edit
'         data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes_paga")
'         data_cli.Recordset.Update
'      End If
'      If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
'         If Int(data_cli.Recordset("cl_ultanop")) <> Int(data_arqueo.Recordset("ano_paga")) Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano_paga")
'            data_cli.Recordset.Update
'         End If
'      Else
'         data_cli.Recordset.Edit
'         data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano_paga")
'         data_cli.Recordset.Update
'      End If
'   End If
'End If

'DoEvents
'MsgBox "Actualización de pagos en base terminado"

Xcant = 0
Xtotal = 0
Xmes = 0
Xano = 0
Data1.RecordSource = "Select * from deudas where estado_cta =" & 1 & " and fecha_pago is null order by cliente,ano,documento DESC"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Xmat = Data1.Recordset("cliente")
   Do While Not Data1.Recordset.EOF
      If Data1.Recordset("cliente") = Xmat Then
         If Data1.Recordset("tipodoc") = "CRE" Then
            Xtotal = Xtotal + Data1.Recordset("total")
         Else
            Xcant = Xcant + 1
            Xtotal = Xtotal + Data1.Recordset("total")
            Xmes = Data1.Recordset("mes")
            Xano = Data1.Recordset("ano")
         End If
      Else
         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            If Xmes = 1 Then
               Xmes = 12
               Xano = Xano - 1
            Else
               Xmes = Xmes - 1
            End If
            If IsNull(data_cli.Recordset("saldo_cc")) = False Then
               If Format(data_cli.Recordset("saldo_cc"), "Standard") <> Format(Xtotal, "Standard") Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("saldo_cc") = Xtotal
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("saldo_cc") = Xtotal
               data_cli.Recordset.Update
            End If

            If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
               If Int(data_cli.Recordset("cl_atrasoa")) <> Int(Xcant) Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_atrasoa") = Xcant
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_atrasoa") = Xcant
               data_cli.Recordset.Update
            End If
            
            If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
               If Int(data_cli.Recordset("cl_ultmesp")) <> Int(Xmes) Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_ultmesp") = Xmes
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_ultmesp") = Xmes
               data_cli.Recordset.Update
            End If

            If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
               If Int(data_cli.Recordset("cl_ultanop")) <> Int(Xano) Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_ultanop") = Xano
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_ultanop") = Xano
               data_cli.Recordset.Update
            End If
         End If
         Xcant = 1
         Xmes = Data1.Recordset("mes")
         Xano = Data1.Recordset("ano")
         Xtotal = Data1.Recordset("total")
      End If
      Xmat = Data1.Recordset("cliente")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      If Xmes = 1 Then
         Xmes = 12
         Xano = Xano - 1
      Else
         Xmes = Xmes - 1
      End If
      If IsNull(data_cli.Recordset("saldo_cc")) = False Then
         If Format(data_cli.Recordset("saldo_cc"), "Standard") <> Format(Xtotal, "Standard") Then
            data_cli.Recordset.Edit
            data_cli.Recordset("saldo_cc") = Xtotal
            data_cli.Recordset.Update
         End If
      Else
         data_cli.Recordset.Edit
         data_cli.Recordset("saldo_cc") = Xtotal
         data_cli.Recordset.Update
      End If

      If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
         If Int(data_cli.Recordset("cl_atrasoa")) <> Int(Xcant) Then
            data_cli.Recordset.Edit
            data_cli.Recordset("cl_atrasoa") = Xcant
            data_cli.Recordset.Update
         End If
      Else
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_atrasoa") = Xcant
         data_cli.Recordset.Update
      End If

      If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
         If Int(data_cli.Recordset("cl_ultmesp")) <> Int(Xmes) Then
            data_cli.Recordset.Edit
            data_cli.Recordset("cl_ultmesp") = Xmes
            data_cli.Recordset.Update
         End If
      Else
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_ultmesp") = Xmes
         data_cli.Recordset.Update
      End If
      
      If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
         If Int(data_cli.Recordset("cl_ultanop")) <> Int(Xano) Then
            data_cli.Recordset.Edit
            data_cli.Recordset("cl_ultanop") = Xano
            data_cli.Recordset.Update
         End If
      Else
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_ultanop") = Xano
         data_cli.Recordset.Update
      End If
   End If
End If
frm_procdeu.MousePointer = 0
MsgBox "Proceso terminado"


End Sub

Private Sub Form_Activate()
txt_mes.Text = Month(Date)
txt_ano.Text = Year(Date)

End Sub

Private Sub Form_Load()
data_cli.Connect = "ODBC;DSN=sapp;"
data_lin.Connect = "ODBC;DSN=sapp;"

'data_env.DatabaseName = App.Path & "\env_deu.mdb"
Data1.DatabaseName = ""
Data1.Connect = "ODBC;DSN=sapp;"
Data2.Connect = "ODBC;DSN=sapp;"


End Sub
