VERSION 5.00
Begin VB.Form frm_pendcons 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Resumen de consultas por médico"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9570
   Icon            =   "frm_pendcons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List6 
      Height          =   2010
      Left            =   4920
      TabIndex        =   14
      Top             =   6360
      Width           =   4215
   End
   Begin VB.ListBox List5 
      Height          =   2010
      Left            =   240
      TabIndex        =   13
      Top             =   6360
      Width           =   4215
   End
   Begin VB.ListBox List4 
      Height          =   2010
      Left            =   4920
      TabIndex        =   10
      Top             =   3840
      Width           =   4215
   End
   Begin VB.ListBox List3 
      Height          =   2010
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   4920
      TabIndex        =   6
      Top             =   1080
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consultar"
      Height          =   375
      Left            =   3600
      Picture         =   "frm_pendcons.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Ingrese número de base o deje en blanco para todas"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      Caption         =   "Presionar el botón consultar para actualizar valores."
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
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "CMT DESPACHO PENDIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   6000
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00008000&
      Caption         =   "CMT DESPACHO REALIZADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6000
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Caption         =   "CMT POLICLÍNICAS PENDIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      Caption         =   "CMT POLICLÍNICAS REALIZADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "MED.GRAL. PRESENCIAL PENDIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "MED.GRAL. PRESENCIAL REALIZADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Ingrese BASE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frm_pendcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Codmed, Cantidadsi, Cantidadno, TotalCantsi, TotalCantno As Long
Codmed = 0
Cantidadsi = 0
Cantidadno = 0
TotalCantsi = 0
TotalCantno = 0
Data2.Connect = "odbc;dsn=sappnew;"

List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
frm_pendcons.MousePointer = 11
Data1.Connect = "odbc;dsn=sappnew;"
If Trim(Text1.Text) = "" Then
   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10001,10003,10005) and nro_med_a is not null order by nro_med_a"
Else
   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10001,10003,10005) and nro_med_a is not null and base =" & Text1.Text & " order by nro_med_a"
End If
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Codmed = Data1.Recordset("nro_med_a")
   Do While Not Data1.Recordset.EOF
      If Codmed = Data1.Recordset("nro_med_a") Then
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio") & " and tipo_consd in ('Consulta Policlínica')"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Cantidadsi = Cantidadsi + 1
            TotalCantsi = TotalCantsi + 1
         Else
            Cantidadno = Cantidadno + 1
            TotalCantno = TotalCantno + 1
         End If
         Data2.Recordset.Close
      Else
         Data1.Recordset.MovePrevious
         List1.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadsi
         List2.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadno
         Data1.Recordset.MoveNext
         Cantidadsi = 0
         Cantidadno = 0
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio") & " and tipo_consd in ('Consulta Policlínica')"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Cantidadsi = Cantidadsi + 1
            TotalCantsi = TotalCantsi + 1
         Else
            Cantidadno = Cantidadno + 1
            TotalCantno = TotalCantno + 1
         End If
         Data2.Recordset.Close
         
      End If
      Codmed = Data1.Recordset("nro_med_a")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   List1.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadsi
   List2.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadno
   Data1.Recordset.MoveNext
   Cantidadsi = 0
   Cantidadno = 0
   
   List1.AddItem "TOTALES.....................: " & TotalCantsi
   List2.AddItem "TOTALES.....................: " & TotalCantno

End If
Data1.Recordset.Close

'cmt polic
Codmed = 0
Cantidadsi = 0
Cantidadno = 0
TotalCantsi = 0
TotalCantno = 0
If Trim(Text1.Text) = "" Then
   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10050) and nro_med_a is not null order by nro_med_a"
Else
   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10050) and nro_med_a is not null and base =" & Text1.Text & " order by nro_med_a"
End If
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Codmed = Data1.Recordset("nro_med_a")
   Do While Not Data1.Recordset.EOF
      If Codmed = Data1.Recordset("nro_med_a") Then
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio") & " and tipo_consd in ('Consulta Policlínica','Orientación Telefónica')"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Cantidadsi = Cantidadsi + 1
            TotalCantsi = TotalCantsi + 1
         Else
            Cantidadno = Cantidadno + 1
            TotalCantno = TotalCantno + 1
         End If
         Data2.Recordset.Close
      Else
         Data1.Recordset.MovePrevious
         List3.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadsi
         List4.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadno
         Data1.Recordset.MoveNext
         Cantidadsi = 0
         Cantidadno = 0
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio") & " and tipo_consd in ('Consulta Policlínica','Orientación Telefónica')"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Cantidadsi = Cantidadsi + 1
            TotalCantsi = TotalCantsi + 1
         Else
            Cantidadno = Cantidadno + 1
            TotalCantno = TotalCantno + 1
         End If
         Data2.Recordset.Close
      End If
      Codmed = Data1.Recordset("nro_med_a")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   List3.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadsi
   List4.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadno
   Data1.Recordset.MoveNext
   Cantidadsi = 0
   Cantidadno = 0
   List3.AddItem "TOTALES.....................: " & TotalCantsi
   List4.AddItem "TOTALES.....................: " & TotalCantno

End If
Data1.Recordset.Close

'''cmt del despacho
Codmed = 0
Cantidadsi = 0
Cantidadno = 0
TotalCantsi = 0
TotalCantno = 0
Data1.RecordSource = "select * from llamado where fec_rea =#" & Format(Date, "yyyy/mm/dd") & "# and codmedcmt is not null and movilpas in (2015) order by codmedcmt"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Codmed = Data1.Recordset("codmedcmt")
   Do While Not Data1.Recordset.EOF
      If Codmed = Data1.Recordset("codmedcmt") Then
         Cantidadsi = Cantidadsi + 1
         TotalCantsi = TotalCantsi + 1
      Else
         Data1.Recordset.MovePrevious
         Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            List5.AddItem Data2.Recordset("nom_med") & " ------>" & Cantidadsi
         Else
            List5.AddItem "MEDICO OTROS " & " ------>" & Cantidadsi
         End If
         Data2.Recordset.Close
         Data1.Recordset.MoveNext
         Cantidadsi = 0
         Cantidadsi = Cantidadsi + 1
      End If
      Codmed = Data1.Recordset("codmedcmt")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      List5.AddItem Data2.Recordset("nom_med") & " ------>" & Cantidadsi
   Else
      List5.AddItem "MEDICO OTROS " & " ------>" & Cantidadsi
   End If
   Data2.Recordset.Close
   List5.AddItem "TOTALES.....................: " & TotalCantsi

End If
Data1.Recordset.Close

Codmed = 0
Cantidadsi = 0
Cantidadno = 0
TotalCantsi = 0
TotalCantno = 0
Dim Xdesdef As Date
Xdesdef = Date - 1
Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xdesdef, "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by codmedcmt"
'Data1.RecordSource = "select * from llamado where fecha >=#" & Format(Xdesdef, "yyyy/mm/dd") & "# and pend in (4) order by codmedcmt"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Codmed = Data1.Recordset("codmedcmt")
   Do While Not Data1.Recordset.EOF
      If Codmed = Data1.Recordset("codmedcmt") Then
         Cantidadno = Cantidadno + 1
         TotalCantno = TotalCantno + 1
      Else
         Data1.Recordset.MovePrevious
         Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            List6.AddItem Data2.Recordset("nom_med") & " ------>" & Cantidadno
         Else
            List6.AddItem "MEDICO OTROS " & " ------>" & Cantidadno
         End If
         Data2.Recordset.Close
         Data1.Recordset.MoveNext
         Cantidadno = 0
         Cantidadno = Cantidadno + 1
      End If
      Codmed = Data1.Recordset("codmedcmt")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      List6.AddItem Data2.Recordset("nom_med") & " ------>" & Cantidadno
   Else
      List6.AddItem "MEDICO OTROS " & " ------>" & Cantidadno
   End If
   Data2.Recordset.Close
   List6.AddItem "TOTALES.....................: " & Data1.Recordset.RecordCount

End If
Data1.Recordset.Close

frm_pendcons.MousePointer = 0


End Sub

Private Sub Form_Load()
Dim Codmed, Cantidadsi, Cantidadno, TotalCantsi, TotalCantno As Long
Codmed = 0
Cantidadsi = 0
Cantidadno = 0
TotalCantsi = 0
TotalCantno = 0
Data2.Connect = "odbc;dsn=sappnew;"

List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
frm_pendcons.MousePointer = 11
Data1.Connect = "odbc;dsn=sappnew;"
Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10001,10003,10005) and nro_med_a is not null order by nro_med_a"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Codmed = Data1.Recordset("nro_med_a")
   Do While Not Data1.Recordset.EOF
      If Codmed = Data1.Recordset("nro_med_a") Then
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio") & " and tipo_consd in ('Consulta Policlínica')"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Cantidadsi = Cantidadsi + 1
            TotalCantsi = TotalCantsi + 1
         Else
            Cantidadno = Cantidadno + 1
            TotalCantno = TotalCantno + 1
         End If
         Data2.Recordset.Close
      Else
         Data1.Recordset.MovePrevious
         List1.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadsi
         List2.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadno
         Data1.Recordset.MoveNext
         Cantidadsi = 0
         Cantidadno = 0
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio") & " and tipo_consd in ('Consulta Policlínica')"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Cantidadsi = Cantidadsi + 1
            TotalCantsi = TotalCantsi + 1
         Else
            Cantidadno = Cantidadno + 1
            TotalCantno = TotalCantno + 1
         End If
         Data2.Recordset.Close
      End If
      Codmed = Data1.Recordset("nro_med_a")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   List1.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadsi
   List2.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadno
   Data1.Recordset.MoveNext
   Cantidadsi = 0
   Cantidadno = 0
   
   List1.AddItem "TOTALES.....................: " & TotalCantsi
   List2.AddItem "TOTALES.....................: " & TotalCantno

End If
Data1.Recordset.Close

'cmt polic
Codmed = 0
Cantidadsi = 0
Cantidadno = 0
TotalCantsi = 0
TotalCantno = 0
Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10050) and nro_med_a is not null order by nro_med_a"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Codmed = Data1.Recordset("nro_med_a")
   Do While Not Data1.Recordset.EOF
      If Codmed = Data1.Recordset("nro_med_a") Then
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio") & " and tipo_consd in ('Consulta Policlínica','Orientación Telefónica') and hora >='" & Data1.Recordset("hora") & "'"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Cantidadsi = Cantidadsi + 1
            TotalCantsi = TotalCantsi + 1
         Else
            Cantidadno = Cantidadno + 1
            TotalCantno = TotalCantno + 1
         End If
         Data2.Recordset.Close
      Else
         Data1.Recordset.MovePrevious
         List3.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadsi
         List4.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadno
         Data1.Recordset.MoveNext
         Cantidadsi = 0
         Cantidadno = 0
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio") & " and tipo_consd in ('Consulta Policlínica','Orientación Telefónica')"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Cantidadsi = Cantidadsi + 1
            TotalCantsi = TotalCantsi + 1
         Else
            Cantidadno = Cantidadno + 1
            TotalCantno = TotalCantno + 1
         End If
         Data2.Recordset.Close
      End If
      Codmed = Data1.Recordset("nro_med_a")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   List3.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadsi
   List4.AddItem Data1.Recordset("nom_med_a") & " ------>" & Cantidadno
   Data1.Recordset.MoveNext
   Cantidadsi = 0
   Cantidadno = 0
   List3.AddItem "TOTALES.....................: " & TotalCantsi
   List4.AddItem "TOTALES.....................: " & TotalCantno

End If
Data1.Recordset.Close

'''cmt del despacho
Codmed = 0
Cantidadsi = 0
Cantidadno = 0
TotalCantsi = 0
TotalCantno = 0
Data1.RecordSource = "select * from llamado where fec_rea =#" & Format(Date, "yyyy/mm/dd") & "# and codmedcmt is not null and movilpas in (2015) order by codmedcmt"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Codmed = Data1.Recordset("codmedcmt")
   Do While Not Data1.Recordset.EOF
      If Codmed = Data1.Recordset("codmedcmt") Then
         Cantidadsi = Cantidadsi + 1
         TotalCantsi = TotalCantsi + 1
      Else
         Data1.Recordset.MovePrevious
         Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            List5.AddItem Data2.Recordset("nom_med") & " ------>" & Cantidadsi
         Else
            List5.AddItem "MEDICO OTROS " & " ------>" & Cantidadsi
         End If
         Data2.Recordset.Close
         Data1.Recordset.MoveNext
         Cantidadsi = 0
         Cantidadsi = Cantidadsi + 1
      End If
      Codmed = Data1.Recordset("codmedcmt")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      List5.AddItem Data2.Recordset("nom_med") & " ------>" & Cantidadsi
   Else
      List5.AddItem "MEDICO OTROS " & " ------>" & Cantidadsi
   End If
   Data2.Recordset.Close
   List5.AddItem "TOTALES.....................: " & TotalCantsi

End If
Data1.Recordset.Close

Codmed = 0
Cantidadsi = 0
Cantidadno = 0
TotalCantsi = 0
TotalCantno = 0
Dim Xdesdef As Date
Xdesdef = Date - 1
Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xdesdef, "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by codmedcmt"
'Data1.RecordSource = "select * from llamado where fecha >=#" & Format(Xdesdef, "yyyy/mm/dd") & "# and pend in (4) order by codmedcmt"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Codmed = Data1.Recordset("codmedcmt")
   Do While Not Data1.Recordset.EOF
      If Codmed = Data1.Recordset("codmedcmt") Then
         Cantidadno = Cantidadno + 1
         TotalCantno = TotalCantno + 1
      Else
         Data1.Recordset.MovePrevious
         Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            List6.AddItem Data2.Recordset("nom_med") & " ------>" & Cantidadno
         Else
            List6.AddItem "MEDICO OTROS " & " ------>" & Cantidadno
         End If
         Data2.Recordset.Close
         Data1.Recordset.MoveNext
         Cantidadno = 0
         Cantidadno = Cantidadno + 1
      End If
      Codmed = Data1.Recordset("codmedcmt")
      Data1.Recordset.MoveNext
   Loop
   Data1.Recordset.MovePrevious
   Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      List6.AddItem Data2.Recordset("nom_med") & " ------>" & Cantidadno
   Else
      List6.AddItem "MEDICO OTROS " & " ------>" & Cantidadno
   End If
   Data2.Recordset.Close
   
   List6.AddItem "TOTALES.....................: " & Data1.Recordset.RecordCount

End If
Data1.Recordset.Close

frm_pendcons.MousePointer = 0


End Sub
