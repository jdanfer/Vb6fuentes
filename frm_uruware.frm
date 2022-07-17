VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_uruware 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controles Facturación SAPP-URUWARE"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7515
   Icon            =   "frm_uruware.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7515
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   3000
      Top             =   2400
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
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin ComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_uruware.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Datos para Controles"
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
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.TextBox t_cod 
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
         Left            =   5280
         TabIndex        =   10
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox t_base 
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Controlar URUWARE con SAPP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Controlar SAPP con URUWARE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Value           =   -1  'True
         Width           =   3615
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Left            =   2400
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Código Uruware:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         ToolTipText     =   "Ejemplos: para base 3 SAPP-001, para base 2: SAPP-002"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese número de BASE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rango de fechas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frm_uruware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim Xrruu As Double
Dim Buscarut As String
Dim Xqueemision As String

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
    
MiBaseact.Execute "Delete * from infvtas"
   
data_inf.RecordSource = "infvtas"
data_inf.Refresh

If Option1.Value = True Then
    If md.Text = "__/__/____" Or mh.Text = "__/__/____" Then
       MsgBox "No ingresó fechas, verifique!", vbInformation
    Else
       Xqueemision = "emi"
       If Month(md.Text) > 9 Then
          Xqueemision = Xqueemision & Trim(str(Month(md.Text)))
       Else
          Xqueemision = Xqueemision & "0" & Trim(str(Month(md.Text)))
       End If
       Xqueemision = Xqueemision & Mid(Trim(str(Year(md.Text))), 3, 2)
       If Data1.Recordset.RecordCount > 0 Then
          frm_uruware.MousePointer = 11
          Data1.Recordset.MoveLast
          pb1.Max = Data1.Recordset.RecordCount
          pb1.Value = 0
          Data1.Recordset.MoveFirst
          DoEvents
          Do While Not Data1.Recordset.EOF
             If Data1.Recordset("Tipo de Cfe") = "EResguardo" Then
             Else
                If IsNull(Data1.Recordset("Documento del receptor")) = False Then
                   If Trim(Data1.Recordset("Documento del receptor")) <> "" Then
                      Xrruu = Val(Data1.Recordset("Documento del receptor"))
                      Buscarut = Data1.Recordset("Documento del receptor")
                      If Len(Trim(str(Val(Buscarut)))) >= 10 Then
                         Data2.RecordSource = "Select * from linmmdd where factura =" & Data1.Recordset("Número") & " and ruc ='" & Trim(Buscarut) & "'"
                      Else
                         Xrruu = 0
                         Data2.RecordSource = "Select * from linmmdd where factura =" & Data1.Recordset("Número") & " and cod_cli =" & Val(Buscarut)
                      End If
                   Else
                      Data2.RecordSource = "Select * from linmmdd where factura =" & Data1.Recordset("Número")
                   End If
                   Data2.Refresh
                   If Data2.Recordset.RecordCount > 0 Then
                   Else
                      Data2.RecordSource = "select * from " & Xqueemision & " where documento =" & Data1.Recordset("Número")
                      Data2.Refresh
                      If Data2.Recordset.RecordCount > 0 Then
                      Else
                         data_inf.Recordset.AddNew
                         data_inf.Recordset("factura") = Data1.Recordset("Número")
                         data_inf.Recordset("cod_cli") = Val(Buscarut)
                         data_inf.Recordset("nom_cli") = Mid(Trim(Data1.Recordset("Razón social del receptor")), 1, 30)
                         data_inf.Recordset("operador") = Data1.Recordset("Caja")
                         data_inf.Recordset("fecha") = Data1.Recordset("Fecha de comprobante")
                         data_inf.Recordset("nom_flia") = Data1.Recordset("Tipo de Cfe")
                         data_inf.Recordset("tot_lin") = Data1.Recordset("Monto total")
                         data_inf.Recordset.Update
                      End If
                   End If
                End If
             End If
             pb1.Value = pb1.Value + 1
             Data1.Recordset.MoveNext
          Loop
          DoEvents
          frm_uruware.MousePointer = 0
          MsgBox "Proceso terminado"
          data_inf.RecordSource = "select * from infvtas"
          data_inf.Refresh
          cr1.ReportFileName = App.path & "\infuruware.rpt"
          cr1.ReportTitle = "CONTROL DE DOCUMENTOS QUE ESTÁN EN URUWARE Y NO EN SAPP " & Format(md.Text, "dd/mm/yyyy") & "-" & Format(mh.Text, "dd/mm/yyyy")
          cr1.Action = 1
          
       End If
    End If
End If

If Option2.Value = True Then
    If md.Text = "__/__/____" Or mh.Text = "__/__/____" Then
       MsgBox "No ingresó fechas, verifique!", vbInformation
    Else
       Data2.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and tot_lin >0 and pendiente not in ('Z','R')"
       Data2.Refresh
       frm_uruware.MousePointer = 11
       If Data2.Recordset.RecordCount > 0 Then
          Data2.Recordset.MoveFirst
          Do While Not Data2.Recordset.EOF
             data_inf.Recordset.AddNew
             data_inf.Recordset("fecha") = Data2.Recordset("fecha")
             data_inf.Recordset("cod_cli") = Data2.Recordset("cod_cli")
             data_inf.Recordset("nom_cli") = Data2.Recordset("nom_cli")
             data_inf.Recordset("base") = Data2.Recordset("base")
             data_inf.Recordset("factura") = Data2.Recordset("factura")
             data_inf.Recordset("cod_prod") = Data2.Recordset("cod_prod")
             data_inf.Recordset("nom_prod") = Data2.Recordset("nom_prod")
             data_inf.Recordset("tot_lin") = Data2.Recordset("tot_lin")
             data_inf.Recordset("ruc") = Data2.Recordset("ruc")
             If IsNull(Data2.Recordset("ced_socio")) = False Then
                If IsNull(Data2.Recordset("fact")) = False Then
                   data_inf.Recordset("nom_superv") = Trim(str(Data2.Recordset("ced_socio"))) & Trim(str(Data2.Recordset("fact")))
                Else
                   data_inf.Recordset("nom_superv") = Trim(str(Data2.Recordset("ced_socio"))) & "0"
                End If
             Else
                data_inf.Recordset("nom_superv") = "0"
             End If
             data_inf.Recordset.Update
             Data2.Recordset.MoveNext
          Loop
       End If
       frm_uruware.MousePointer = 0
       MsgBox "Terminado proceso de bases. Continúa proceso de emisión."
       Xqueemision = "emi"
       If Month(md.Text) > 9 Then
          Xqueemision = Xqueemision & Trim(str(Month(md.Text)))
       Else
          Xqueemision = Xqueemision & "0" & Trim(str(Month(md.Text)))
       End If
       Xqueemision = Xqueemision & Mid(Trim(str(Year(md.Text))), 3, 2)
       Data2.RecordSource = "select * from " & Xqueemision
       Data2.Refresh
       If Data2.Recordset.RecordCount > 0 Then
          frm_uruware.MousePointer = 11
          Data2.Recordset.MoveFirst
          Do While Not Data2.Recordset.EOF
             data_inf.Recordset.AddNew
             data_inf.Recordset("fecha") = Data2.Recordset("fecha")
             data_inf.Recordset("cod_cli") = Data2.Recordset("cliente")
             data_inf.Recordset("nom_cli") = Mid(Data2.Recordset("apellidos"), 1, 30)
             data_inf.Recordset("base") = 206
             data_inf.Recordset("factura") = Data2.Recordset("documento")
             data_inf.Recordset("cod_prod") = 11111
             data_inf.Recordset("nom_prod") = Data2.Recordset("origen")
             data_inf.Recordset("tot_lin") = Data2.Recordset("total")
             data_inf.Recordset("ruc") = Data2.Recordset("ruc")
             If IsNull(Data2.Recordset("cedula")) = False Then
                If IsNull(Data2.Recordset("cod")) = False Then
                   data_inf.Recordset("nom_superv") = Trim(str(Data2.Recordset("cedula"))) & Trim(str(Data2.Recordset("cod")))
                Else
                   data_inf.Recordset("nom_superv") = Trim(str(Data2.Recordset("cedula"))) & "0"
                End If
             Else
                data_inf.Recordset("nom_superv") = "0"
             End If
             data_inf.Recordset.Update
             Data2.Recordset.MoveNext
          Loop
       End If
       frm_uruware.MousePointer = 0
       MsgBox "Terminado proceso de emisión. Comienzan controles..."
       data_inf.RecordSource = "select * from infvtas"
       data_inf.Refresh
       
       If Data1.Recordset.RecordCount > 0 Then
          frm_uruware.MousePointer = 11
          Data1.Recordset.MoveLast
          pb1.Max = Data1.Recordset.RecordCount
          pb1.Value = 0
          Data1.Recordset.MoveFirst
          DoEvents
          Do While Not Data1.Recordset.EOF
             If Data1.Recordset("Tipo de Cfe") = "EResguardo" Then
             Else
                If IsNull(Data1.Recordset("Documento del receptor")) = False Then
                   If Trim(Data1.Recordset("Documento del receptor")) <> "" Then
                      Xrruu = Val(Data1.Recordset("Documento del receptor"))
                      Buscarut = Data1.Recordset("Documento del receptor")
                      If Len(Trim(str(Val(Buscarut)))) >= 10 Then
                         data_inf.RecordSource = "Select * from infvtas where factura =" & Data1.Recordset("Número") & " and ruc ='" & Trim(Buscarut) & "'"
                      Else
                         If Data1.Recordset("Tipo de documento del receptor") = "CI" Then
                            Xrruu = 0
                            data_inf.RecordSource = "Select * from infvtas where factura =" & Data1.Recordset("Número") & " and nom_superv ='" & Trim(Buscarut) & "'"
                         Else
                            Xrruu = 0
                            data_inf.RecordSource = "Select * from infvtas where factura =" & Data1.Recordset("Número") & " and cod_cli =" & Val(Buscarut)
                         End If
                      End If
                   Else
                      data_inf.RecordSource = "Select * from infvtas where factura =" & Data1.Recordset("Número")
                   End If
                   data_inf.Refresh
                   If data_inf.Recordset.RecordCount > 0 Then
                      data_inf.Recordset.MoveFirst
                      Do While Not data_inf.Recordset.EOF
                         data_inf.Recordset.Edit
                         data_inf.Recordset("nro_flia") = 55
                         data_inf.Recordset.Update
                         data_inf.Recordset.MoveNext
                      Loop
                   End If
                End If
             End If
             pb1.Value = pb1.Value + 1
             Data1.Recordset.MoveNext
          Loop
          DoEvents
          data_inf.RecordSource = "infvtas"
          data_inf.Refresh
          MiBaseact.Execute "Delete * from infvtas where nro_flia =" & 55
          data_inf.RecordSource = "infvtas"
          data_inf.Refresh
          frm_uruware.MousePointer = 0
          MsgBox "Proceso terminado"
          cr1.ReportFileName = App.path & "\infuruware.rpt"
          cr1.ReportTitle = "CONTROL DE DOCUMENTOS QUE ESTÁN EN SAPP Y NO EN URUWARE " & Format(md.Text, "dd/mm/yyyy") & "-" & Format(mh.Text, "dd/mm/yyyy")
          cr1.Action = 1
          
       End If
    End If
End If


End Sub

Private Sub Form_Load()
Data1.DatabaseName = "C:\uruware\uruware.xls"
Data1.RecordSource = "Datos$"
Data1.Refresh
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_inf.DatabaseName = App.path & "\informes.mdb"

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option1.SetFocus
End If

End Sub
