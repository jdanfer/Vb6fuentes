VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_pasaraRP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasar Deudas a RED PAGOS"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9075
   Icon            =   "frm_pasaraRP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9075
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccionar datos"
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox t_mat 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5880
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "odbc;dsn=sappnew;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Procesar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4440
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7680
         Picture         =   "frm_pasaraRP.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Consultar datos"
         Top             =   1920
         Width           =   615
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_pasaraRP.frx":0B14
         Height          =   2055
         Left            =   240
         OleObjectBlob   =   "frm_pasaraRP.frx":0B28
         TabIndex        =   8
         Top             =   2400
         Width           =   8055
      End
      Begin VB.TextBox t_h 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox t_d 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox t_cob 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_pasaraRP.frx":28CB
         Left            =   2640
         List            =   "frm_pasaraRP.frx":28D8
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Matrícula:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rango de e-ticket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cobrador:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seleccione opción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frm_pasaraRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Pasar todo un cobrador seleccionado
'Pasar por número de ticket emisión
'Pasar por número de ticket deuda
Dim Xelerror As Integer
Xelerror = 0

If Combo1.Text = "Pasar todo un cobrador seleccionado" Then
   If t_cob.Text <> "" Then
      If t_cob.Text > 0 Then
         If t_d.Text <> "" And t_h.Text <> "" Then
            Data1.RecordSource = "select * from deudas where fecha_pago is null and nro_cobr =" & t_cob.Text & " and documento >=" & t_d.Text & " and documento <=" & t_h.Text & " order by fecha"
         Else
            Data1.RecordSource = "select * from deudas where fecha_pago is null and nro_cobr =" & t_cob.Text & " order by fecha"
         End If
         Data1.Refresh
      Else
         Xelerror = 9
         MsgBox "Debe ingresar número de cobrador"
      End If
   Else
      Xelerror = 9
      MsgBox "Debe ingresar número de cobrador"
   End If
Else
   If Combo1.Text = "Pasar por número de ticket emisión" Or Combo1.Text = "Pasar por número de ticket deuda" Then
      If t_d.Text <> "" And t_h.Text <> "" Then
         Data1.RecordSource = "select * from deudas where fecha_pago is null and documento >=" & t_d.Text & " and documento <=" & t_h.Text & " order by fecha"
         Data1.Refresh
      Else
         Xelerror = 9
         MsgBox "No ingresó rango de números"
      End If
   Else
      If Combo1.Text = "Pasar por número de cliente" Then
         If t_mat.Text <> "" Then
            Data1.RecordSource = "select * from deudas where fecha_pago is null and cliente =" & t_mat.Text & " order by fecha"
            Data1.Refresh
         Else
            Xelerror = 9
            MsgBox "No ingresó matrícula"
         End If
      Else
         Xelerror = 9
         MsgBox "No seleccionó datos"
      
      End If
   End If
End If

If Xelerror = 0 Then
   Command2.Enabled = True
Else
   Command2.Enabled = False
End If

End Sub

Private Sub Command2_Click()
Dim Xsiproceso As String

Xsiproceso = MsgBox("Desea procesar los registros seleccionados para cobrador RedPagos?", vbInformation + vbYesNo)
If Xsiproceso = vbYes Then
   If Data1.Recordset.RecordCount > 0 Then
      frm_pasaraRP.MousePointer = 11
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         Data2.Recordset.AddNew
         Data2.Recordset("fecha") = Date
         Data2.Recordset("hora") = Format(Time, "HH:mm")
         Data2.Recordset("usuario") = WElusuario
         Data2.Recordset("matricula") = Data1.Recordset("cliente")
         Data2.Recordset("documento") = Data1.Recordset("documento")
         If IsNull(Data1.Recordset("nro_cobr")) = False Then
            Data2.Recordset("cobrador") = Data1.Recordset("nro_cobr")
         Else
            Data2.Recordset("cobrador") = 0
         End If
         Data2.Recordset.Update
                  
         Data1.Recordset.Edit
         Data1.Recordset("nro_cobr") = 221
         Data1.Recordset("nom_cobr") = "RED PAGOS"
         Data1.Recordset.Update
         
         Data3.RecordSource = "select * from arqueo where nrorec =" & Data1.Recordset("documento") & " and matricula =" & Data1.Recordset("cliente")
         Data3.Refresh
         If Data3.Recordset.RecordCount > 0 Then
            Data3.Recordset.Edit
            Data3.Recordset("cob") = 221
            Data3.Recordset("nomcob") = "RED PAGOS"
            Data3.Recordset.Update
         Else
            Data3.Recordset.AddNew
            Data3.Recordset("matricula") = Data1.Recordset("matricula")
            Data3.Recordset("nombre") = Data1.Recordset("nombre")
            If IsNull(Data1.Recordset("mes")) = False Then
               If Data1.Recordset("mes") > 0 Then
                  Data3.Recordset("mes") = Data1.Recordset("mes")
                  Data3.Recordset("ano") = Data1.Recordset("ano")
               Else
                  Data3.Recordset("mes") = Month(Data1.Recordset("fecha"))
                  Data3.Recordset("ano") = Year(Data1.Recordset("fecha"))
               End If
            Else
               Data3.Recordset("mes") = Month(Date)
               Data3.Recordset("ano") = Year(Date)
            End If
            Data3.Recordset("color") = "T"
            Data3.Recordset("cat") = Data1.Recordset("cod_cnv")
            Data3.Recordset("nomcat") = Data1.Recordset("nom_cnv")
            Data3.Recordset("arqueo") = "P"
            Data3.Recordset("importe") = Data1.Recordset("importe")
            Data3.Recordset("fecha") = Data1.Recordset("fecha")
            Data3.Recordset("nrorec") = Data1.Recordset("documento")
            Data3.Recordset("usuar") = WElusuario
            Data3.Recordset("moneda") = 2
            Data3.Recordset("cob") = 221
            Data3.Recordset("nomcob") = "RED PAGOS"
            Data3.Recordset("codzon") = 0
            Data3.Recordset("codpro") = 0
            Data3.Recordset("codsup") = 0
            Data3.Recordset("tiquet") = Data1.Recordset("tiquet")
            Data3.Recordset("total") = Data1.Recordset("total")
            Data3.Recordset("varia") = 0
            Data3.Recordset("iva") = Data1.Recordset("iva")
            Data3.Recordset("deudas") = Data1.Recordset("deudas")
            Data3.Recordset("servi") = 0
            Data3.Recordset.Update
            
         End If
         
         Data1.Recordset.MoveNext
      Loop
      frm_pasaraRP.MousePointer = 0
      Command2.Enabled = False
      MsgBox "Proceso terminado", vbInformation
   End If
End If

End Sub

Private Sub Form_Load()

Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data2.RecordSource = "select * from control_redpago where id=" & 2
Data2.Refresh

Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub
