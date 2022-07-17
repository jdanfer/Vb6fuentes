VERSION 5.00
Begin VB.Form frm_genarq 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar emisión y generar Arqueo"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "frm_genarq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Crear una tab"
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   480
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
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_arqueo 
      Caption         =   "data_arqueo"
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
      RecordSource    =   "ARQUEO"
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton btn_cierra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CERRAR"
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
      Left            =   4560
      Picture         =   "frm_genarq.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton btn_proc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROCESAR"
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
      Left            =   960
      Picture         =   "frm_genarq.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txt_anoa 
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
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txt_mesa 
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
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txt_anoe 
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
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txt_mese 
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
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MES Y AÑO DE ARQUEO A GENERAR:"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MES Y AÑO DE EMISION A CARGAR"
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
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Picture         =   "frm_genarq.frx":0CC6
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frm_genarq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CRViewer1_CloseButtonClicked(UseDefault As Boolean)

End Sub

Private Sub btn_cierra_Click()
frm_genarq.Hide

End Sub

Private Sub btn_proc_Click()
Dim MiBasear As Database
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Dim UnaSesionar As Workspace
Set UnaSesionar = Workspaces(0)
Dim Archivoar As String
Dim Xfec As Date
Dim Tabla1 As TableDef
Dim Recarq As Recordset
btn_proc.Enabled = False
btn_cierra.Enabled = False

'On Error GoTo Yaesta
Archivo = App.path & "\arqueos.mdb" 'Tabla local para guardar la generación
Set MiBasear = UnaSesionar.OpenDatabase(Archivo)

'''Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\sapp.mdb")

ConectarBD
ConbdSapp.Open

Dim Nomarq As String
Dim Nomemisa As String
Dim Nomemiuno As String
Dim Xmes As Integer
Dim Xano As Integer

'Dim Adocat As ADOX.Catalog

If txt_mese.Text <> "" And txt_anoe.Text <> "" Then
    Xfec = Date
    frm_genarq.MousePointer = 11
    Nomemisa = "em"
    Nomemiuno = "emi"
    Nomarq = "arq"
    Xmes = Val(txt_mese.Text)
    Xano = Val(txt_anoe.Text)
    If Xmes = 1 Then
       Xmes = 12
       Xano = Xano - 1
    Else
       Xmes = Xmes - 1
       Xano = Xano
    End If
    If Xmes < 10 Then
       Nomarq = Nomarq + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
    Else
       Nomarq = Nomarq + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
    End If
    
    If txt_mese.Text < 10 Then
       Nomemisa = Nomemisa + "0" + Trim(txt_mese.Text) + Mid(Trim(txt_anoe.Text), 3, 2)
       Nomemiuno = Nomemiuno + "0" + Trim(txt_mese.Text) + Mid(Trim(txt_anoe.Text), 3, 2)
    Else
       Nomemiuno = Nomemiuno + Trim(txt_mese.Text) + Mid(Trim(txt_anoe.Text), 3, 2)
    End If
    Set Tabla1 = MiBasear.CreateTableDef(Nomarq)
    Dim matricula As Field, Nombre As Field, mes As Field
    Dim ano As Field, color As Field, cat As Field
    Dim nomcat As Field, arqueo As Field, importe As Field
    Dim fecha As Field, nrorec As Field, usuar As Field
    Dim moneda As Field, cob As Field, Nomcob As Field
    Dim codzon As Field, codpro As Field, codsup As Field
    Dim tiquet As Field, total As Field, varia As Field
    Dim iva As Field, deudas As Field, servi As Field
'    Set matricula = Tabla1.CreateField("matricula", dbLong, 11)
    Set matricula = Tabla1.CreateField("matricula", dbLong)
    
    Set Nombre = Tabla1.CreateField("nombre", dbText, 60)
    Set mes = Tabla1.CreateField("mes", dbInteger)
    Set ano = Tabla1.CreateField("ano", dbInteger)
    Set color = Tabla1.CreateField("color", dbText, 1)
    Set cat = Tabla1.CreateField("cat", dbText, 6)
    Set nomcat = Tabla1.CreateField("nomcat", dbText, 60)
    Set arqueo = Tabla1.CreateField("arqueo", dbText, 1)
    Set importe = Tabla1.CreateField("importe", dbLong, 12)
    Set fecha = Tabla1.CreateField("fecha", dbDate)
    Set nrorec = Tabla1.CreateField("nrorec", dbLong, 12)
    Set usuar = Tabla1.CreateField("usuar", dbText, 15)
    Set moneda = Tabla1.CreateField("moneda", dbInteger)
    Set cob = Tabla1.CreateField("cob", dbInteger, 3)
    Set Nomcob = Tabla1.CreateField("nomcob", dbText, 35)
    Set codzon = Tabla1.CreateField("codzon", dbInteger, 3)
    Set codpro = Tabla1.CreateField("codpro", dbInteger, 3)
    Set codsup = Tabla1.CreateField("codsup", dbInteger, 3)
    Set tiquet = Tabla1.CreateField("tiquet", dbLong, 10)
    Set total = Tabla1.CreateField("total", dbLong, 12)
    Set varia = Tabla1.CreateField("varia", dbLong, 10)
    Set iva = Tabla1.CreateField("iva", dbLong, 10)
    Set deudas = Tabla1.CreateField("deudas", dbLong, 10)
    Set servi = Tabla1.CreateField("servi", dbLong, 10)
    With Tabla1
     .Fields.Append matricula
     .Fields.Append Nombre
     .Fields.Append mes
     .Fields.Append ano
     .Fields.Append color
     .Fields.Append cat
     .Fields.Append nomcat
     .Fields.Append arqueo
     .Fields.Append importe
     .Fields.Append fecha
     .Fields.Append nrorec
     .Fields.Append usuar
     .Fields.Append moneda
     .Fields.Append cob
     .Fields.Append Nomcob
     .Fields.Append codzon
     .Fields.Append codpro
     .Fields.Append codsup
     .Fields.Append tiquet
     .Fields.Append total
     .Fields.Append varia
     .Fields.Append iva
     .Fields.Append deudas
     .Fields.Append servi
    End With
    
    With MiBasear
     .TableDefs.Append Tabla1
     End With
    Set Recarq = MiBasear.OpenRecordset(Nomarq)
'    createTablearqueo (Nomarq) 'Función anterior que Crea la tabla arqueo en mysql
    ConbdSapp.Execute "create table " & Nomarq & " (matricula int(10) default NULL," & _
    " nombre varchar(60) default NULL, mes int(2) default NULL, ano int(4) default NULL," & _
    " color varchar(1) default NULL, cat varchar(6) default NULL, nomcat varchar(60) default NULL," & _
    " arqueo varchar(1) default NULL, importe double(15,2) default NULL, fecha datetime default NULL," & _
    " nrorec int(10) default NULL, usuar varchar(15) default NULL, tipo varchar(10) default NULL," & _
    " moneda int(1) default NULL, cob int(3) default NULL, nomcob varchar(35) default NULL," & _
    " codzon int(3) default NULL, codpro int(3) default NULL, codsup int(1) default NULL," & _
    " tiquet double(10,2) default NULL, total double(15,2) default NULL, varia double(15,2) default NULL," & _
    " iva double(15,2) default NULL, deudas double(15,2) default NULL, servi double(15,2) default NULL," & _
    " nro int(10) NOT NULL auto_increment, PRIMARY KEY (nro)) ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            
    data_arqueo.Recordset.MoveFirst
    Do While Not data_arqueo.Recordset.EOF
    'respalda el arqueo anterior en ARQMMAA
       Recarq.AddNew
       Recarq("matricula") = data_arqueo.Recordset("matricula")
       Recarq("nombre") = data_arqueo.Recordset("nombre")
       Recarq("mes") = data_arqueo.Recordset("mes")
       Recarq("ano") = data_arqueo.Recordset("ano")
       Recarq("color") = data_arqueo.Recordset("color")
       Recarq("cat") = data_arqueo.Recordset("cat")
       Recarq("nomcat") = data_arqueo.Recordset("nomcat")
       Recarq("arqueo") = data_arqueo.Recordset("arqueo")
       Recarq("importe") = data_arqueo.Recordset("importe")
       Recarq("fecha") = data_arqueo.Recordset("fecha")
       Recarq("nrorec") = data_arqueo.Recordset("nrorec")
       Recarq("usuar") = WElusuario
       Recarq("moneda") = data_arqueo.Recordset("moneda")
       Recarq("cob") = data_arqueo.Recordset("cob")
       Recarq("nomcob") = data_arqueo.Recordset("nomcob")
       Recarq("codzon") = data_arqueo.Recordset("codzon")
       Recarq("codpro") = data_arqueo.Recordset("codpro")
       Recarq("codsup") = data_arqueo.Recordset("codsup")
       Recarq("tiquet") = data_arqueo.Recordset("tiquet")
       Recarq("total") = data_arqueo.Recordset("total")
       Recarq("varia") = data_arqueo.Recordset("varia")
       Recarq("iva") = data_arqueo.Recordset("iva")
       Recarq("deudas") = data_arqueo.Recordset("deudas")
       Recarq("servi") = 0
       Recarq.Update
       data_arqueo.Recordset.MoveNext
    Loop
    data_arqueo.Recordset.MoveFirst
    DoEvents
    MsgBox "Terminado proceso de respaldo en tabla local. Continuar..."
    Data1.RecordSource = Nomarq
    Data1.Refresh
    Do While Not data_arqueo.Recordset.EOF
       Data1.Recordset.AddNew
       Data1.Recordset("matricula") = data_arqueo.Recordset("matricula")
       Data1.Recordset("nombre") = data_arqueo.Recordset("nombre")
       Data1.Recordset("mes") = data_arqueo.Recordset("mes")
       Data1.Recordset("ano") = data_arqueo.Recordset("ano")
       Data1.Recordset("color") = data_arqueo.Recordset("color")
       Data1.Recordset("cat") = data_arqueo.Recordset("cat")
       Data1.Recordset("nomcat") = data_arqueo.Recordset("nomcat")
       Data1.Recordset("arqueo") = data_arqueo.Recordset("arqueo")
       Data1.Recordset("importe") = data_arqueo.Recordset("importe")
       Data1.Recordset("fecha") = data_arqueo.Recordset("fecha")
       Data1.Recordset("nrorec") = data_arqueo.Recordset("nrorec")
       Data1.Recordset("usuar") = WElusuario
       Data1.Recordset("moneda") = data_arqueo.Recordset("moneda")
       Data1.Recordset("cob") = data_arqueo.Recordset("cob")
       Data1.Recordset("nomcob") = data_arqueo.Recordset("nomcob")
       Data1.Recordset("codzon") = data_arqueo.Recordset("codzon")
       Data1.Recordset("codpro") = data_arqueo.Recordset("codpro")
       Data1.Recordset("codsup") = data_arqueo.Recordset("codsup")
       Data1.Recordset("tiquet") = data_arqueo.Recordset("tiquet")
       Data1.Recordset("total") = data_arqueo.Recordset("total")
       Data1.Recordset("varia") = data_arqueo.Recordset("varia")
       Data1.Recordset("iva") = data_arqueo.Recordset("iva")
       Data1.Recordset("deudas") = data_arqueo.Recordset("deudas")
       Data1.Recordset("servi") = 0
       Data1.Recordset.Update
       data_arqueo.Recordset.MoveNext
    Loop
    DoEvents
    MsgBox "Terminado proceso de carga datos a tabla ARQMMAA"
    'MiBaseact.Execute "Delete from arqueo where arqueo <> 'P'"
    'MiBaseact.Execute "Update arqueo set arqueo ='P'"
    ConbdSapp.Execute "update arqueo set arqueo ='P' where arqueo ='E' and cob in (221,514,607,673,683,690)"
    ConbdSapp.Execute "Delete from arqueo where arqueo <> 'P'"
    ConbdSapp.Execute "Update arqueo set codpro =" & 0 & " where codpro=" & 98

''''    ConbdSapp.Execute "Update arqueo set arqueo ='P'"
    MsgBox "Terminado proceso de actualización de datos facturas Pend."
    data_emi.RecordSource = "Select * from " & Nomemiuno & " order by documento"
    data_emi.Refresh
    data_emi.Recordset.MoveFirst
'    data_arqueo.Refresh
'    If data_arqueo.Recordset.RecordCount > 0 Then
'       data_arqueo.Recordset.MoveFirst
'       Do While Not data_arqueo.Recordset.EOF
'          If data_arqueo.Recordset("usuar") = WElusuario Then
'          Else
'             data_arqueo.Recordset.Edit
'             data_arqueo.Recordset("usuar") = WElusuario
'             data_arqueo.Recordset.Update
'          End If
'          data_arqueo.Recordset.MoveNext
'       Loop
'    End If
    data_arqueo.Recordset.MoveFirst
    Do While Not data_emi.Recordset.EOF
       data_arqueo.Recordset.AddNew
       data_arqueo.Recordset("matricula") = data_emi.Recordset("cliente")
       data_arqueo.Recordset("nombre") = data_emi.Recordset("apellidos")
       data_arqueo.Recordset("mes") = data_emi.Recordset("mes")
       data_arqueo.Recordset("ano") = data_emi.Recordset("ano")
       data_arqueo.Recordset("color") = data_emi.Recordset("color_rec")
       data_arqueo.Recordset("cat") = data_emi.Recordset("cod_cnv")
       data_arqueo.Recordset("nomcat") = data_emi.Recordset("nom_cnv")
       data_arqueo.Recordset("arqueo") = "E"
       data_arqueo.Recordset("importe") = data_emi.Recordset("importe")
       data_arqueo.Recordset("fecha") = Date
       data_arqueo.Recordset("nrorec") = data_emi.Recordset("documento")
       data_arqueo.Recordset("usuar") = WElusuario
       data_arqueo.Recordset("moneda") = data_emi.Recordset("moneda")
       data_arqueo.Recordset("cob") = data_emi.Recordset("nro_cobr")
       data_arqueo.Recordset("nomcob") = data_emi.Recordset("nom_cobr")
       If IsNull(data_emi.Recordset("grupo")) = False Then
          data_arqueo.Recordset("codzon") = data_emi.Recordset("grupo")
       Else
          data_arqueo.Recordset("codzon") = 0
       End If
       data_arqueo.Recordset("codsup") = data_emi.Recordset("nro_superv")
       data_arqueo.Recordset("codpro") = data_emi.Recordset("nro_vende")
       data_arqueo.Recordset("tiquet") = data_emi.Recordset("tiquet")
       data_arqueo.Recordset("total") = data_emi.Recordset("total")
       data_arqueo.Recordset("varia") = data_emi.Recordset("deudas")
       data_arqueo.Recordset("iva") = data_emi.Recordset("iva")
       data_arqueo.Recordset("deudas") = data_emi.Recordset("deudas")
       data_arqueo.Recordset("servi") = 0
       data_arqueo.Recordset.Update
       data_emi.Recordset.MoveNext
    Loop
    MsgBox "Terminado carga de EMISION al arqueo"
    ConbdSapp.Execute "update arqueo set arqueo ='" & "E" & "' where arqueo ='" & "P" & "'"
        
    frm_genarq.MousePointer = 0
    btn_cierra.Enabled = True
    ConbdSapp.Close
    MsgBox "Proceso Terminado...", vbInformation, "Arqueos"
    Unload Me
Else
    MsgBox "Debe ingresar mes de emisión", vbCritical, "Arqueos"
    txt_mese.SetFocus
End If


Exit Sub
Yaesta:
       If Err.Number = 3010 Then
          MsgBox "Ya se generó arqueo " + Nomemi, vbInformation, "Arqueos"
          frm_genarq.MousePointer = 0
       Else
          MsgBox "Error al generar ERROR:" + str(Err.Number), vbCritical, "Arqueos"
          frm_genarq.MousePointer = 0
          End
       End If

End Sub

Private Sub Form_Initialize()
If Month(Date) = 1 Then
   txt_mese.Text = 12
   txt_mesa.Text = 12
   txt_anoe.Text = Year(Date) - 1
   txt_anoa.Text = Year(Date) - 1
Else
   txt_mese.Text = Month(Date) - 1
   txt_mesa.Text = Month(Date) - 1
   txt_anoe.Text = Year(Date)
   txt_anoa.Text = Year(Date)
End If

End Sub

Private Sub Form_Load()
'data_arqueo.DatabaseName = App.Path & "\sapp.mdb"
Dim Xsihay As Integer
Xsihay = 0
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_arqueo.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_arqueo.RecordSource = "select * from arqueo where codpro not in (98) order by cob"
'data_arqueo.Refresh
'If data_arqueo.Recordset.RecordCount > 0 Then
'   data_arqueo.Recordset.MoveFirst
'   Do While Not data_arqueo.Recordset.EOF
'      If data_arqueo.Recordset("cob") = 1 Or data_arqueo.Recordset("cob") = 113 Or _
'         data_arqueo.Recordset("cob") = 201 Or data_arqueo.Recordset("cob") = 208 Or _
'         data_arqueo.Recordset("cob") = 206 Or data_arqueo.Recordset("cob") = 209 Or _
'         data_arqueo.Recordset("cob") = 221 Or data_arqueo.Recordset("cob") = 224 Or _
'         data_arqueo.Recordset("cob") = 225 Or data_arqueo.Recordset("cob") = 514 Or _
'         data_arqueo.Recordset("cob") = 602 Or data_arqueo.Recordset("cob") = 607 Or _
'         data_arqueo.Recordset("cob") = 615 Or data_arqueo.Recordset("cob") = 616 Or _
'         data_arqueo.Recordset("cob") = 624 Or data_arqueo.Recordset("cob") = 635 Or _
'         data_arqueo.Recordset("cob") = 636 Or data_arqueo.Recordset("cob") = 641 Or _
'         data_arqueo.Recordset("cob") = 653 Or data_arqueo.Recordset("cob") = 673 Or _
'         data_arqueo.Recordset("cob") = 679 Or data_arqueo.Recordset("cob") = 683 Or _
'         data_arqueo.Recordset("cob") = 685 Or data_arqueo.Recordset("cob") = 690 Or _
'         data_arqueo.Recordset("cob") = 696 Or data_arqueo.Recordset("cob") = 701 Then
'      Else
'         Xsihay = 1
'      End If
'      data_arqueo.Recordset.MoveNext
'   Loop
'End If

'If Xsihay = 1 Then
'   MsgBox "ATENCION!!! Existen cobradores con arqueo sin cerrar, VERIFIQUE!!", vbCritical
'   End
'End If
data_arqueo.RecordSource = "select * from arqueo"
data_arqueo.Refresh

data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"
''data_emi.DatabaseName = App.Path & "\emisiones.mdb"
data_emi.RecordSource = ""
data_emi.Refresh


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_anoa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btn_proc.SetFocus
End If

End Sub

Private Sub txt_anoe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mesa.SetFocus
   txt_mesa.Text = txt_mese.Text
   txt_anoa.Text = txt_anoe.Text
End If

End Sub

Private Sub txt_mesa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_anoa.SetFocus
End If

End Sub

Private Sub txt_mese_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_anoe.SetFocus
End If

End Sub

Private Function createTablearqueo(tablename As String) As Boolean
    Dim conODBCDirect As DAO.Connection
    Dim rsODBCDirect As DAO.Recordset
    Dim strConn As String
    strConn = "ODBC;DSN=SAPP;"
    Set WrkODBC = CreateWorkspace("", "admin", "", dbUseODBC)
    Set conODBCDirect = WrkODBC.OpenConnection("", , , strConn)
    On Error Resume Next
'''    tablename = "ARQ0108"
    conODBCDirect.Execute ("call prcreatearq('" & tablename & "')")
    If Err <> 0 Then
        createTablearqueo = False
    Else
        createTablearqueo = True
    End If

End Function
