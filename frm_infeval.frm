VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infeval 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes evaluación de desempeño"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   Icon            =   "frm_infeval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5460
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr7 
      Left            =   2520
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr6 
      Left            =   3120
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr5 
      Left            =   2040
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr4 
      Left            =   1200
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr3 
      Left            =   600
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   120
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_persontot 
      Caption         =   "data_persontot"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data data_jefe 
      Caption         =   "data_jefe"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_inff 
      Caption         =   "data_inff"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salir"
      Height          =   735
      Left            =   3000
      Picture         =   "frm_infeval.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Procesar"
      Height          =   735
      Left            =   600
      Picture         =   "frm_infeval.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de informe"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Solo un formulario"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Incluir solo Evaluaciones Jefaturas"
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
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   4695
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frm_infeval.frx":109E
         Left            =   1440
         List            =   "frm_infeval.frx":10A0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Data data_buscap 
         Caption         =   "data_buscap"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data data_us 
         Caption         =   "data_us"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data data_inf2 
         Caption         =   "data_inf2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data data_pregunt 
         Caption         =   "data_pregunt"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data data_person 
         Caption         =   "data_person"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data data_cargo 
         Caption         =   "data_cargo"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data data_perio 
         Caption         =   "data_perio"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_infeval.frx":10A2
         Left            =   1440
         List            =   "frm_infeval.frx":10A4
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   3375
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
         ItemData        =   "frm_infeval.frx":10A6
         Left            =   1440
         List            =   "frm_infeval.frx":10D4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Jefatura:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Período:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Informe de:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   120
      Picture         =   "frm_infeval.frx":1227
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frm_infeval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.ListIndex = 11 Or Combo1.ListIndex = 9 Then
   Check1.Enabled = True
Else
   Check1.Enabled = False
End If

End Sub

Private Sub Combo2_Click()
'Combo2.Clear

'If Combo2.Text = "2015" Then
'   Data1.Connect = "ODBC;DSN=eval2015;"
'   data_us.Connect = "ODBC;DSN=sapp;"
'   data_persontot.Connect = "ODBC;DSN=eval2015;"
'   data_cargo.Connect = "ODBC;DSN=eval2015;"
'   data_buscap.Connect = "ODBC;DSN=eval2015;"
'   data_pregunt.Connect = "ODBC;DSN=eval2015;"
'   data_person.Connect = "ODBC;DSN=eval2015;"
    
'   data_jefe.Connect = "ODBC;DSN=eval2015;"
'   data_jefe.RecordSource = "jefes"
'   data_jefe.Refresh
    
'   Combo3.Clear
'   If data_jefe.Recordset.RecordCount > 0 Then
'      data_jefe.Recordset.MoveFirst
'      Do While Not data_jefe.Recordset.EOF
'         Combo3.AddItem data_jefe.Recordset("descrip")
'         data_jefe.Recordset.MoveNext
'      Loop
'   End If
'   data_perio.Connect = "ODBC;DSN=eval2015;"
'   data_perio.RecordSource = "periodo"
'   data_perio.Refresh
'   If data_perio.Recordset.RecordCount > 0 Then
'      data_perio.Recordset.MoveFirst
'      Do While Not data_perio.Recordset.EOF
'         Combo2.AddItem data_perio.Recordset("descrip")
'         data_perio.Recordset.MoveNext
'      Loop
'      Combo2.AddItem "2016"
       
'   End If
'Else
'    Data1.Connect = "ODBC;DSN=sappper;"
'    data_us.Connect = "ODBC;DSN=sapp;"
'    data_persontot.Connect = "ODBC;DSN=sappper;"
'    data_cargo.Connect = "ODBC;DSN=sappper;"
'    data_buscap.Connect = "ODBC;DSN=sappper;"
'    data_pregunt.Connect = "ODBC;DSN=sappper;"
'    data_person.Connect = "ODBC;DSN=sappper;"
    
'    data_jefe.Connect = "ODBC;DSN=sappper;"
'    data_jefe.RecordSource = "jefes"
'    data_jefe.Refresh
    
'    Combo3.Clear
'    If data_jefe.Recordset.RecordCount > 0 Then
'       data_jefe.Recordset.MoveFirst
'       Do While Not data_jefe.Recordset.EOF
'          Combo3.AddItem data_jefe.Recordset("descrip")
'          data_jefe.Recordset.MoveNext
'       Loop
'    End If
'    data_perio.Connect = "ODBC;DSN=sappper;"
'    data_perio.RecordSource = "periodo"
'    data_perio.Refresh
'    If data_perio.Recordset.RecordCount > 0 Then
'       data_perio.Recordset.MoveFirst
'       Do While Not data_perio.Recordset.EOF
'          Combo2.AddItem data_perio.Recordset("descrip")
'          data_perio.Recordset.MoveNext
'       Loop
'       Combo2.AddItem "2015"
       
'    End If
'End If

End Sub

Private Sub Command1_Click()
Dim Xelresul As String
Dim Xlace As Long
Dim Xlacejef As Long
Dim Xcomenta1, Xcomenta2, Xcomenta3, Xcomenta4 As String
Dim Xp1, Xp2, Xp3, Xp4, Xp5, Xp6, Xp7, Xp8 As Integer
Dim Xp9, Xp10, Xp11, Xp12, Xp13, Xp14, Xp15, Xp16 As Integer
Dim Xp17, Xp118, Xp19, Xp20, Xp21, Xp22, Xp23, Xp24 As Integer
Dim Xp25, Xp26, Xp27, Xp28, Xp29, Xp30, Xp31, Xp32 As Integer
Dim Xst1, Xst2, Xst3, Xst4 As Integer

'13
'On Error GoTo Errevalinf
Command1.Enabled = False
Command2.Enabled = False

If Combo2.Text = "2016" Then
   Data1.Connect = "ODBC;DSN=eval2015;"
   data_us.Connect = "odbc;dsn=" & Xconexrmt & ";"
   data_persontot.Connect = "ODBC;DSN=eval2015;"
   data_cargo.Connect = "ODBC;DSN=eval2015;"
   data_buscap.Connect = "ODBC;DSN=eval2015;"
   data_pregunt.Connect = "ODBC;DSN=eval2015;"
   data_person.Connect = "ODBC;DSN=eval2015;"
   data_jefe.Connect = "ODBC;DSN=eval2015;"
   data_jefe.RecordSource = "jefes"
   data_jefe.Refresh
   Combo3.Clear
   If data_jefe.Recordset.RecordCount > 0 Then
      data_jefe.Recordset.MoveFirst
      Do While Not data_jefe.Recordset.EOF
         Combo3.AddItem data_jefe.Recordset("descrip")
         data_jefe.Recordset.MoveNext
      Loop
   End If
End If


data_inf.RecordSource = "infcli"
data_inf.Refresh

If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If
frm_infeval.MousePointer = 11

If Combo1.ListIndex = 0 Then
   If Combo3.ListIndex < 0 Then
    '   Data1.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & Combo2.Text & "' and idtitulo =" & 1 & " and idjefe =" & Wxeljefeid & " order by idpregun"
      Data1.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & Combo2.Text & "' and id2 =" & Wxelnroid2 & " order by idjefe,idpregun"
       'wxelnroid2
    '   Data1.RecordSource = "Select * from evaluas order by idjefe,idpregun"
    '    Data1.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev
      Data1.Refresh
       If Data1.Recordset.RecordCount > 0 Then
    '      Data1.Recordset.MoveLast
          
       End If
       If Data1.Recordset.RecordCount > 0 Then
          Data1.Recordset.MoveFirst
          data_person.RecordSource = "Select * from personas where id =" & Wxelnrocedev & " and id2 =" & Wxelnroid2
          data_person.Refresh
          Do While Not Data1.Recordset.EOF
             If data_person.Recordset.RecordCount > 0 Then
                data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_person.Recordset("cargod") & "'"
                data_cargo.Refresh
                If data_cargo.Recordset.RecordCount > 0 Then
                   data_pregunt.RecordSource = "Select * from textos where codcargo =" & data_cargo.Recordset("codpreg") & " and pregunta =" & Data1.Recordset("idpregun")
                   data_pregunt.Refresh
                   If data_pregunt.Recordset.RecordCount > 0 Then
                      data_inf.Recordset.AddNew
                      data_inf.Recordset("info_debit") = data_pregunt.Recordset("descrip")
                      data_inf.Recordset("cl_codced") = data_pregunt.Recordset("pregunta")
                      data_inf.Recordset("cl_direcci") = Mid(Data1.Recordset("titulo"), 1, 80)
                      data_inf.Recordset("cl_forpago") = Data1.Recordset("idtitulo")
                      data_inf.Recordset("cl_telefon") = Data1.Recordset("periodo")
                      data_inf.Recordset("cl_apellid") = Mid(frm_evalua1.labnome.Caption, 1, 60)
                      data_inf.Recordset("cl_nombre") = Mid(data_person.Recordset("cargod"), 1, 30)
                      data_inf.Recordset("cl_codigo") = 0
                      data_inf.Recordset("cl_cedula") = 0
                      If Data1.Recordset("idempl") = Data1.Recordset("idjefe") Then
                         data_inf.Recordset("cl_codigo") = Data1.Recordset("puntos")
                         data_inf.Recordset("cl_nrocobr") = 7
                      Else
                         data_inf.Recordset("cl_cedula") = Data1.Recordset("puntos")
                         data_inf.Recordset("cl_nrocobr") = 8 'Jefatura
                      End If
                      data_inf.Recordset("obsp") = Data1.Recordset("obs")
                      data_inf.Recordset.Update
                   End If
                End If
             End If
             Data1.Recordset.MoveNext
          Loop
          data_inf.RecordSource = "Select * from infcli where cl_nrocobr =" & 8
          data_inf.Refresh
          If data_inf.Recordset.RecordCount > 0 Then
             data_inf2.RecordSource = "Select * from infcli"
             data_inf2.Refresh
             data_inf.Recordset.MoveLast
             data_inf.Recordset.MoveFirst
             Do While Not data_inf.Recordset.EOF
                data_inf2.RecordSource = "Select * from infcli where cl_nrocobr =" & 7 & " and cl_codced =" & data_inf.Recordset("cl_codced")
                data_inf2.Refresh
                If data_inf2.Recordset.RecordCount > 0 Then
                   data_inf.Recordset.Edit
                   data_inf.Recordset("cl_codigo") = data_inf2.Recordset("cl_codigo")
                   data_inf.Recordset.Update
                   data_inf2.Recordset.Edit
                   data_inf2.Recordset("cl_nrocobr") = 9
                   data_inf2.Recordset.Update
                
                End If
                data_inf.Recordset.MoveNext
             Loop
          End If
          data_inf.RecordSource = "Select * from infcli where cl_nrocobr in (7,8)"
          data_inf.Refresh
          If data_inf.Recordset.RecordCount > 0 Then
             data_inf.Recordset.MoveFirst
             Do While Not data_inf.Recordset.EOF
                data_inf.Recordset.Edit
                Xelresul = data_inf.Recordset("cl_codigo") + data_inf.Recordset("cl_cedula")
                Xelresul = Xelresul / 2
                data_inf.Recordset("cl_nrovend") = Xelresul
                data_inf.Recordset.Update
    
                data_inf.Recordset.MoveNext
             Loop
          End If
          data_inf.RecordSource = "Select * from infcli"
          data_inf.Refresh
          
          data_inf.Recordset.MoveFirst
          Do While Not data_inf.Recordset.EOF
             If data_inf.Recordset("cl_nrocobr") = 9 Then
                data_inf.Recordset.Delete
             End If
             data_inf.Recordset.MoveNext
          Loop
          data_inf.Refresh
          frm_infeval.MousePointer = 0
          MsgBox "Proceso terminado"
          cr1.ReportFileName = App.path & "\infformeval.rpt"
          cr1.Action = 1
          cr2.ReportFileName = App.path & "\infformevalc.rpt"
          cr2.Action = 1
          
          
       Else
          MsgBox "No hay registros de evaluación"
          
       End If
   Else
       Dim Xcuantosrpt As Integer
       Xcuantosrpt = 0
       data_persontot.RecordSource = "Select * from personas where jefed ='" & Combo3.Text & "'"
       data_persontot.Refresh
       If data_persontot.Recordset.RecordCount > 0 Then
          data_persontot.Recordset.MoveFirst
          Do While Not data_persontot.Recordset.EOF
             Data1.RecordSource = "Select * from evaluas where idempl =" & data_persontot.Recordset("id") & " and periodo ='" & Combo2.Text & "' and id2 =" & data_persontot.Recordset("id2") & " order by idjefe,idpregun"
             Data1.Refresh
             If Data1.Recordset.RecordCount > 0 Then
                Data1.Recordset.MoveFirst
                data_person.RecordSource = "Select * from personas where id =" & data_persontot.Recordset("id") & " and id2 =" & data_persontot.Recordset("id2")
                data_person.Refresh
                Do While Not Data1.Recordset.EOF
                   If data_person.Recordset.RecordCount > 0 Then
                      data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_person.Recordset("cargod") & "'"
                      data_cargo.Refresh
                      If data_cargo.Recordset.RecordCount > 0 Then
                         data_pregunt.RecordSource = "Select * from textos where codcargo =" & data_cargo.Recordset("codpreg") & " and pregunta =" & Data1.Recordset("idpregun")
                         data_pregunt.Refresh
                         If data_pregunt.Recordset.RecordCount > 0 Then
                            data_inf.Recordset.AddNew
                            data_inf.Recordset("info_debit") = data_pregunt.Recordset("descrip")
                            data_inf.Recordset("cl_codced") = data_pregunt.Recordset("pregunta")
                            data_inf.Recordset("cl_direcci") = Mid(Data1.Recordset("titulo"), 1, 80)
                            data_inf.Recordset("cl_forpago") = Data1.Recordset("idtitulo")
                            data_inf.Recordset("cl_apellid") = Mid(data_persontot.Recordset("ape1"), 1, 30) & " " & Mid(data_persontot.Recordset("nom1"), 1, 29)
                            data_inf.Recordset("cl_nombre") = Mid(data_person.Recordset("cargod"), 1, 30)
                            data_inf.Recordset("cl_codigo") = 0
                            data_inf.Recordset("cl_cedula") = 0
                            If Data1.Recordset("idempl") = Data1.Recordset("idjefe") Then
                               data_inf.Recordset("cl_codigo") = Data1.Recordset("puntos")
                               data_inf.Recordset("cl_nrocobr") = 7
                            Else
                               data_inf.Recordset("cl_cedula") = Data1.Recordset("puntos")
                               data_inf.Recordset("cl_nrocobr") = 8 'Jefatura
                            End If
                            data_inf.Recordset("obsp") = Data1.Recordset("obs")
                            data_inf.Recordset.Update
                         End If
                      End If
                   End If
                   Data1.Recordset.MoveNext
                Loop
                data_inf.RecordSource = "Select * from infcli where cl_nrocobr =" & 8
                data_inf.Refresh
                If data_inf.Recordset.RecordCount > 0 Then
                   data_inf2.RecordSource = "Select * from infcli"
                   data_inf2.Refresh
                   data_inf.Recordset.MoveLast
                   data_inf.Recordset.MoveFirst
                   Do While Not data_inf.Recordset.EOF
                      data_inf2.RecordSource = "Select * from infcli where cl_nrocobr =" & 7 & " and cl_codced =" & data_inf.Recordset("cl_codced")
                      data_inf2.Refresh
                      If data_inf2.Recordset.RecordCount > 0 Then
                         data_inf.Recordset.Edit
                         data_inf.Recordset("cl_codigo") = data_inf2.Recordset("cl_codigo")
                         data_inf.Recordset.Update
                         data_inf2.Recordset.Edit
                         data_inf2.Recordset("cl_nrocobr") = 9
                         data_inf2.Recordset.Update
                     
                      End If
                      data_inf.Recordset.MoveNext
                   Loop
                End If
                data_inf.RecordSource = "Select * from infcli where cl_nrocobr in (7,8)"
                data_inf.Refresh
                If data_inf.Recordset.RecordCount > 0 Then
                   data_inf.Recordset.MoveFirst
                   Do While Not data_inf.Recordset.EOF
                      data_inf.Recordset.Edit
                      Xelresul = data_inf.Recordset("cl_codigo") + data_inf.Recordset("cl_cedula")
                      Xelresul = Xelresul / 2
                      data_inf.Recordset("cl_nrovend") = Xelresul
                      data_inf.Recordset.Update
          
                      data_inf.Recordset.MoveNext
                   Loop
                End If
                data_inf.RecordSource = "Select * from infcli"
                data_inf.Refresh
                If data_inf.Recordset.RecordCount > 0 Then
                    data_inf.Recordset.MoveFirst
                    Do While Not data_inf.Recordset.EOF
                       If data_inf.Recordset("cl_nrocobr") = 9 Then
                          data_inf.Recordset.Delete
                       End If
                       data_inf.Recordset.MoveNext
                    Loop
                    data_inf.Refresh
                    frm_infeval.MousePointer = 0
                    Xcuantosrpt = Xcuantosrpt + 1
                    If Xcuantosrpt <= 14 Then
                       cr1.ReportFileName = App.path & "\infformeval.rpt"
                       cr1.Action = 1
                    Else
                       If Xcuantosrpt <= 28 Then
                          cr2.ReportFileName = App.path & "\infformeval2.rpt"
                          cr2.Action = 1
                       Else
                          If Xcuantosrpt <= 42 Then
                             cr3.ReportFileName = App.path & "\infformeval3.rpt"
                             cr3.Action = 1
                          Else
                             If Xcuantosrpt <= 56 Then
                                cr4.ReportFileName = App.path & "\infformeval4.rpt"
                                cr4.Action = 1
                             Else
                                If Xcuantosrpt <= 70 Then
                                   cr5.ReportFileName = App.path & "\infformeval5.rpt"
                                   cr5.Action = 1
                                Else
                                   If Xcuantosrpt <= 84 Then
                                      cr6.ReportFileName = App.path & "\infformeval6.rpt"
                                      cr6.Action = 1
                                   Else
                                      If Xcuantosrpt <= 98 Then
                                         cr7.ReportFileName = App.path & "\infformeval7.rpt"
                                         cr7.Action = 1
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                End If
             End If
             data_persontot.Recordset.MoveNext
             If data_inf.Recordset.RecordCount > 0 Then
                data_inf.Recordset.MoveFirst
                Do While Not data_inf.Recordset.EOF
                   data_inf.Recordset.Delete
                   data_inf.Recordset.MoveNext
                Loop
                data_inf.Refresh
             End If
          Loop
          MsgBox "Proceso terminado"
       End If
   End If
End If
Dim Xnombb As String

If Combo1.ListIndex = 3 Then
'   Data1.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & Combo2.Text & "' and idtitulo =" & 1 & " and idjefe =" & Wxeljefeid & " order by idpregun"
   
   data_us.RecordSource = "Select * from cap_ciap where des_cap ='" & WElusuario & "'"
   data_us.Refresh
   If data_us.Recordset.RecordCount > 0 Then
      data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
      data_buscap.Refresh
      If data_buscap.Recordset.RecordCount > 0 Then
         Wxeljefeid = data_buscap.Recordset("id")
      Else
         MsgBox "No se encontró el usuario, comunique al administrador"
         frm_infeval.MousePointer = 0
         Unload Me
      End If
   Else
      MsgBox "No se encuentra usuario registrado, comunique al administrador", vbInformation
      frm_infeval.MousePointer = 0
      Unload Me
   End If
   If IsNull(data_buscap.Recordset("cargod")) = False Then
      data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_buscap.Recordset("cargod") & "'"
      data_cargo.Refresh
      If data_cargo.Recordset("tipo") = 2 Or data_cargo.Recordset("tipo") = 3 Or WElusuario = "DARIOH" Or WElusuario = "JFERNAN" Then
         If data_cargo.Recordset("tipo") = 2 And WElusuario <> "JFERNAN" Then
            data_person.RecordSource = "Select * from personas where jefed ='" & data_cargo.Recordset("descrip") & "' or cargod ='" & data_cargo.Recordset("descrip") & "' order by id"
            data_person.Refresh
         Else
            data_person.RecordSource = "Select * from personas"
            data_person.Refresh
         End If
      Else
         MsgBox "Usuario no registrado para informes"
         Unload Me
      End If
      If data_person.Recordset.RecordCount > 0 Then
         data_person.Recordset.MoveFirst
         Do While Not data_person.Recordset.EOF
            Data1.RecordSource = "Select * from evaluas where idempl =" & data_person.Recordset("id")
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
'               Data1.Recordset.MoveLast
'               If Data1.Recordset.RecordCount >= 32 Then
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_cedula") = Data1.Recordset("idempl")
                  Xnombb = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
                  data_inf.Recordset("cl_apellid") = Mid(Xnombb, 1, 60)
                  data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
                  data_inf.Recordset("cl_nombre") = Data1.Recordset("periodo")
                  data_inf.Recordset("cl_nomcobr") = Mid(data_person.Recordset("cargod"), 1, 20)
                  If IsNull(Data1.Recordset("firma")) = False Then
                     If Data1.Recordset("firma") = 5 Then
                        data_inf.Recordset("cl_codconv") = "SI"
                     Else
                        data_inf.Recordset("cl_codconv") = "NO"
                     End If
                  Else
                     data_inf.Recordset("cl_codconv") = "NO"
                  End If
                  data_inf.Recordset.Update
'               End If
            End If
            data_person.Recordset.MoveNext
         Loop
         frm_infeval.MousePointer = 0
         MsgBox "Terminado"
         cr1.ReportFileName = App.path & "\infpereval.rpt"
         cr1.Action = 1
      End If
   Else
      MsgBox "No se encuentra cargo para informes"
      frm_infeval.MousePointer = 0
      Unload Me
      
   End If
End If

If Combo1.ListIndex = 4 Then
'   Data1.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & Combo2.Text & "' and idtitulo =" & 1 & " and idjefe =" & Wxeljefeid & " order by idpregun"
   
   data_us.RecordSource = "Select * from cap_ciap where des_cap ='" & WElusuario & "'"
   data_us.Refresh
   If data_us.Recordset.RecordCount > 0 Then
      data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
      data_buscap.Refresh
      If data_buscap.Recordset.RecordCount > 0 Then
         Wxeljefeid = data_buscap.Recordset("id")
      Else
         MsgBox "No se encontró el usuario, comunique al administrador"
         frm_infeval.MousePointer = 0
         Unload Me
      End If
   Else
      MsgBox "No se encuentra usuario registrado, comunique al administrador", vbInformation
      frm_infeval.MousePointer = 0
      Unload Me
   End If
   If IsNull(data_buscap.Recordset("cargod")) = False Then
      data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_buscap.Recordset("cargod") & "'"
      data_cargo.Refresh
      If data_cargo.Recordset("tipo") = 2 Or data_cargo.Recordset("tipo") = 3 Or WElusuario = "DARIOH" Or WElusuario = "JFERNAN" Then
         If data_cargo.Recordset("tipo") = 2 And WElusuario <> "JFERNAN" Then
            data_person.RecordSource = "Select * from personas where jefed ='" & data_cargo.Recordset("descrip") & "' or cargod ='" & data_cargo.Recordset("descrip") & "' order by id"
            data_person.Refresh
         Else
            data_person.RecordSource = "Select * from personas"
            data_person.Refresh
         End If
      Else
         MsgBox "Usuario no registrado para informes"
         frm_infeval.MousePointer = 0
         Unload Me
      End If
      If data_person.Recordset.RecordCount > 0 Then
         data_person.Recordset.MoveFirst
         Do While Not data_person.Recordset.EOF
            Data1.RecordSource = "Select * from evaluas where idempl =" & data_person.Recordset("id")
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
            Else
'               Data1.Recordset.MoveLast
'               If Data1.Recordset.RecordCount >= 32 Then
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_cedula") = data_person.Recordset("id")
                  Xnombb = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
                  data_inf.Recordset("cl_apellid") = Mid(Xnombb, 1, 60)
'                  data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
                  data_inf.Recordset("cl_nombre") = Mid(data_person.Recordset("cargod"), 1, 20)
                  data_inf.Recordset.Update
'               End If
            End If
            data_person.Recordset.MoveNext
         Loop
         frm_infeval.MousePointer = 0
         MsgBox "Terminado"
         cr1.ReportFileName = App.path & "\infperneval.rpt"
         cr1.Action = 1
      End If
   Else
      MsgBox "No se encuentra cargo para informes"
      frm_infeval.MousePointer = 0
      Unload Me
      
   End If
End If


If Combo1.ListIndex = 13 Then
   Dim Xcierre As Integer
   Xcierre = 0
'   Data1.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & Combo2.Text & "' and idtitulo =" & 1 & " and idjefe =" & Wxeljefeid & " order by idpregun"
   If Combo3.ListIndex >= 0 Then
      data_person.RecordSource = "Select * from personas where jefed ='" & Combo3.Text & "' order by id"
      data_person.Refresh
   Else
      data_person.RecordSource = "Select * from personas"
      data_person.Refresh
   End If
   If data_person.Recordset.RecordCount > 0 Then
      data_person.Recordset.MoveLast
      data_person.Recordset.MoveFirst
      Do While Not data_person.Recordset.EOF
         Data1.RecordSource = "Select * from evaluas where idempl =" & data_person.Recordset("id")
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
               If IsNull(Data1.Recordset("cierre")) = False Then
                  If Data1.Recordset("cierre") = "SI" Then
                     Xcierre = Xcierre + 1
                  Else
                     Xcierre = 0
                  End If
               Else
                  Xcierre = 0
               End If
               Data1.Recordset.MoveNext
            Loop
            If Xcierre >= 64 Then
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_cedula") = data_person.Recordset("id")
               Xnombb = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
               data_inf.Recordset("cl_apellid") = Mid(Xnombb, 1, 60)
               data_inf.Recordset("cl_nombre") = Mid(data_person.Recordset("cargod"), 1, 20)
               data_inf.Recordset("cl_codconv") = "SI"
               data_inf.Recordset.Update
            Else
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_cedula") = data_person.Recordset("id")
               Xnombb = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
               data_inf.Recordset("cl_apellid") = Mid(Xnombb, 1, 60)
               data_inf.Recordset("cl_nombre") = Mid(data_person.Recordset("cargod"), 1, 20)
               data_inf.Recordset("cl_codconv") = "NO"
               data_inf.Recordset.Update
            End If
         End If
         Xcierre = 0
         data_person.Recordset.MoveNext
      Loop
      frm_infeval.MousePointer = 0
      MsgBox "Terminado"
      cr1.ReportFileName = App.path & "\infperneval22.rpt"
      cr1.Action = 1
   End If
End If

If Combo1.ListIndex = 5 Then
'   Data1.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & Combo2.Text & "' and idtitulo =" & 1 & " and idjefe =" & Wxeljefeid & " order by idpregun"
   
   data_us.RecordSource = "Select * from cap_ciap where des_cap ='" & WElusuario & "'"
   data_us.Refresh
   If data_us.Recordset.RecordCount > 0 Then
      data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
      data_buscap.Refresh
      If data_buscap.Recordset.RecordCount > 0 Then
         Wxeljefeid = data_buscap.Recordset("id")
      Else
         MsgBox "No se encontró el usuario, comunique al administrador"
         frm_infeval.MousePointer = 0
         Unload Me
      End If
   Else
      MsgBox "No se encuentra usuario registrado, comunique al administrador", vbInformation
      frm_infeval.MousePointer = 0
      Unload Me
   End If
   If IsNull(data_buscap.Recordset("cargod")) = False Then
      data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_buscap.Recordset("cargod") & "'"
      data_cargo.Refresh
      If data_cargo.Recordset("tipo") = 2 Or data_cargo.Recordset("tipo") = 3 Or WElusuario = "DARIOH" Or WElusuario = "JFERNAN" Then
         If data_cargo.Recordset("tipo") = 2 And WElusuario <> "JFERNAN" Then
            data_person.RecordSource = "Select * from personas where jefed ='" & data_cargo.Recordset("descrip") & "' or cargod ='" & data_cargo.Recordset("descrip") & "' order by id"
            data_person.Refresh
         Else
            data_person.RecordSource = "Select * from personas"
            data_person.Refresh
         End If
      Else
         MsgBox "Usuario no registrado para informes"
         Unload Me
      End If
      If data_person.Recordset.RecordCount > 0 Then
         data_person.Recordset.MoveFirst
         Do While Not data_person.Recordset.EOF
            Data1.RecordSource = "Select * from evaluas where idempl =" & data_person.Recordset("id")
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.MoveLast
               If Data1.Recordset.RecordCount >= 64 Then
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_cedula") = Data1.Recordset("idempl")
                  Xnombb = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
                  data_inf.Recordset("cl_apellid") = Mid(Xnombb, 1, 60)
                  data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
                  data_inf.Recordset("cl_nombre") = Data1.Recordset("periodo")
                  data_inf.Recordset("cl_nomcobr") = Mid(data_person.Recordset("cargod"), 1, 20)
                  If IsNull(Data1.Recordset("firma")) = False Then
                     If Data1.Recordset("firma") = 5 Then
                        data_inf.Recordset("cl_codconv") = "SI"
                     Else
                        data_inf.Recordset("cl_codconv") = "NO"
                     End If
                  Else
                     data_inf.Recordset("cl_codconv") = "NO"
                  End If
                  data_inf.Recordset.Update
               End If
            End If
            data_person.Recordset.MoveNext
         Loop
         frm_infeval.MousePointer = 0
         MsgBox "Terminado"
         cr1.ReportFileName = App.path & "\infperevalj.rpt"
         cr1.Action = 1
      End If
   Else
      MsgBox "No se encuentra cargo para informes"
      frm_infeval.MousePointer = 0
      Unload Me
      
   End If
End If


If Combo1.ListIndex = 9 Then
   Dim Xprom1, Xprom2, Xprom3, Xprom4, Xcanp, Xdif As Integer
   Dim Xcomjef, Xcomxpreg As String
   Xcomjef = ""
   Xcomxpreg = ""
   frm_infeval.MousePointer = 11
   data_inff.RecordSource = "infor1"
   data_inff.Refresh
   If data_inff.Recordset.RecordCount > 0 Then
      data_inff.Recordset.MoveFirst
      Do While Not data_inff.Recordset.EOF
         data_inff.Recordset.Delete
         data_inff.Recordset.MoveNext
      Loop
   End If
   If Combo3.ListIndex >= 0 Then
      data_person.RecordSource = "Select * from personas where jefed ='" & Combo3.Text & "'"
      data_person.Refresh
   Else
      data_person.RecordSource = "Select * from personas order by jefed"
      data_person.Refresh
   End If
''''''' BUSCAR POR EL ID2 PARA LOS REPETIDOS
   If data_person.Recordset.RecordCount > 0 Then
      data_person.Recordset.MoveLast
      pb1.Max = data_person.Recordset.RecordCount
      pb1.Value = 0
      data_person.Recordset.MoveFirst
      Xlace = data_person.Recordset("id2")
      DoEvents
      Xcanp = 0
      Do While Not data_person.Recordset.EOF
         If data_person.Recordset("cargod") = Combo3.Text Then
         Else
'            Data1.RecordSource = "Select * from evaluas where id2 =" & data_person.Recordset("id2") & " and idjefe = idempl order by idpregun"
            If Check1.Value = 1 Then
               Data1.RecordSource = "Select * from evaluas where id2 =" & data_person.Recordset("id2") & " and idempl <> idjefe and idpregun not in (33,34,35,36,37,38,39,40) order by idpregun"
               Data1.Refresh
            Else
               Data1.RecordSource = "Select * from evaluas where id2 =" & data_person.Recordset("id2") & " and idpregun not in (33,34,35,36,37,38,39,40) order by idpregun"
               Data1.Refresh
            End If
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.MoveFirst
               Do While Not Data1.Recordset.EOF
                  If Data1.Recordset("idtitulo") = 1 Then
                     Xprom1 = Xprom1 + Data1.Recordset("puntos")
                  End If
                  If Data1.Recordset("idtitulo") = 2 Then
                     Xprom2 = Xprom2 + Data1.Recordset("puntos")
                  End If
                  If Data1.Recordset("idtitulo") = 3 Then
                     Xprom3 = Xprom3 + Data1.Recordset("puntos")
                  End If
                  If Data1.Recordset("idtitulo") = 4 Then
                     Xprom4 = Xprom4 + Data1.Recordset("puntos")
                  End If
                  If IsNull(Data1.Recordset("obs")) = False Then
                     If Trim(Data1.Recordset("obs")) <> "" Then
                        data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_person.Recordset("cargod") & "'"
                        data_cargo.Refresh
                        If data_cargo.Recordset.RecordCount > 0 Then
                           data_pregunt.RecordSource = "Select * from textos where codcargo =" & data_cargo.Recordset("codpreg") & " and pregunta =" & Data1.Recordset("idpregun")
                           data_pregunt.Refresh
                           If data_pregunt.Recordset.RecordCount > 0 Then
                              If Xcomxpreg = "" Then
                                 Xcomxpreg = Trim(data_pregunt.Recordset("descrip"))
                              Else
                                 Xcomxpreg = Xcomxpreg & Chr(13) & Chr(10) & Trim(data_pregunt.Recordset("descrip"))
                              End If
                              If Xcomjef = "" Then
                                 Xcomjef = Trim(data_pregunt.Recordset("descrip")) & Chr(13) + Chr(10) & Data1.Recordset("obs")
                              Else
                                 Xcomjef = Xcomjef & Chr(13) & Chr(10) & Trim(data_pregunt.Recordset("descrip")) & Chr(13) + Chr(10) & Data1.Recordset("obs")
                              End If
                           End If
                        End If
                     End If
                  End If
                  Xlace = Data1.Recordset("idempl")
                  Xcanp = Xcanp + 1
                  Data1.Recordset.MoveNext
               Loop
               DoEvents
               If Check1.Value = 1 Then
                  If Xcanp > 30 Then
                     data_inff.Recordset.AddNew
                     data_inff.Recordset("id") = data_person.Recordset("id2")
                     data_inff.Recordset("nombre") = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
                     data_inff.Recordset("cargo") = data_person.Recordset("cargod")
                     data_inff.Recordset("jefe") = data_person.Recordset("jefed")
                     data_inff.Recordset("prom1") = Xprom1 / 8
                     data_inff.Recordset("prom2") = Xprom2 / 8
                     data_inff.Recordset("prom3") = Xprom3 / 8
                     data_inff.Recordset("prom4") = Xprom4 / 8
                     data_inff.Recordset("cantper") = data_person.Recordset.RecordCount
                     data_inff.Recordset("tot") = Xprom1 + Xprom2 + Xprom3 + Xprom4
                     data_inff.Recordset("tot") = data_inff.Recordset("tot") / 32
                     data_inff.Recordset("obs") = Xcomjef
                     data_inff.Recordset.Update
                  Else
                     Xdif = Xdif + 1
                  End If
               Else
                  If Xcanp > 35 Then
                     data_inff.Recordset.AddNew
                     data_inff.Recordset("id") = data_person.Recordset("id2")
                     data_inff.Recordset("nombre") = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
                     data_inff.Recordset("cargo") = data_person.Recordset("cargod")
                     data_inff.Recordset("jefe") = data_person.Recordset("jefed")
                     data_inff.Recordset("prom1") = Xprom1 / 16
                     data_inff.Recordset("prom2") = Xprom2 / 16
                     data_inff.Recordset("prom3") = Xprom3 / 16
                     data_inff.Recordset("prom4") = Xprom4 / 16
                     data_inff.Recordset("cantper") = data_person.Recordset.RecordCount
                     data_inff.Recordset("tot") = Xprom1 + Xprom2 + Xprom3 + Xprom4
                     data_inff.Recordset("tot") = data_inff.Recordset("tot") / 64
                     data_inff.Recordset("obs") = Xcomjef
                     data_inff.Recordset.Update
                  Else
                     Xdif = Xdif + 1
                  End If
               End If
            End If
         End If
         Xprom1 = 0
         Xprom2 = 0
         Xprom3 = 0
         Xprom4 = 0
         Xcanp = 0
         Xcomjef = ""
         data_person.Recordset.MoveNext
         pb1.Value = pb1.Value + 1
      Loop
      frm_infeval.MousePointer = 0
      MsgBox "Terminado"
      If Combo3.ListIndex >= 0 Then
         cr1.ReportFileName = App.path & "\infprom1.rpt"
      Else
         cr1.ReportFileName = App.path & "\infprom1t.rpt"
      End If
      cr1.Action = 1
   End If
End If


If Combo1.ListIndex = 10 Then
   Dim Xprom11, Xprom22, Xcanpp, Xdiff As Integer
   Dim Xelcar As String
   frm_infeval.MousePointer = 11
   data_inff.RecordSource = "porcargo"
   data_inff.Refresh
   If data_inff.Recordset.RecordCount > 0 Then
      data_inff.Recordset.MoveFirst
      Do While Not data_inff.Recordset.EOF
         data_inff.Recordset.Delete
         data_inff.Recordset.MoveNext
      Loop
   End If
   If Combo3.ListIndex >= 0 Then
      data_person.RecordSource = "Select * from personas order by cargod"
      data_person.Refresh
   Else
      data_person.RecordSource = "Select * from personas order by cargod"
      data_person.Refresh
   End If
''''''' BUSCAR POR EL ID2 PARA LOS REPETIDOS
   Xelcar = ""
   If data_person.Recordset.RecordCount > 0 Then
      data_person.Recordset.MoveLast
      pb1.Max = data_person.Recordset.RecordCount
      pb1.Value = 0
      data_person.Recordset.MoveFirst
      DoEvents
      Xcanpp = 0
      Xelcar = data_person.Recordset("cargod")
      Do While Not data_person.Recordset.EOF
         If Xelcar = data_person.Recordset("cargod") Then
            Data1.RecordSource = "Select * from evaluas where id2 =" & data_person.Recordset("id2") & " and idjefe = idempl and idpregun not in (33,34,35,36,37,38,39,40) order by idpregun"
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.MoveFirst
               Do While Not Data1.Recordset.EOF
                  Xprom11 = Xprom11 + Data1.Recordset("puntos")
                  Data1.Recordset.MoveNext
               Loop
               DoEvents
               Xcanpp = Xcanpp + 1
            End If
            Data1.RecordSource = "Select * from evaluas where id2 =" & data_person.Recordset("id2") & " and idjefe <> idempl order by idpregun"
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.MoveFirst
               Do While Not Data1.Recordset.EOF
                  Xprom22 = Xprom22 + Data1.Recordset("puntos")
                  Data1.Recordset.MoveNext
               Loop
               DoEvents
            End If
            Xelcar = data_person.Recordset("cargod")
            data_person.Recordset.MoveNext
         Else
            If Xprom11 > 4 And Xprom22 > 4 Then
               data_person.Recordset.MovePrevious
               data_inff.Recordset.AddNew
               data_inff.Recordset("id") = 1
               data_inff.Recordset("cargo") = data_person.Recordset("cargod")
               data_inff.Recordset("promauto") = Xprom11 / 32
               data_inff.Recordset("promauto") = data_inff.Recordset("promauto") / Xcanpp
               data_inff.Recordset("promeval") = Xprom22 / 32
               data_inff.Recordset("promeval") = data_inff.Recordset("promeval") / Xcanpp
               data_inff.Recordset("promedio") = Xprom11 + Xprom22
               data_inff.Recordset("promedio") = data_inff.Recordset("promedio") / 64
               data_inff.Recordset("promedio") = data_inff.Recordset("promedio") / Xcanpp
'               data_inff.Recordset("promedio") = data_inff.Recordset("promedio") / 2
               data_inff.Recordset.Update
'               MsgBox "PERSONAS:" & Xcanpp & " PUNTOS AUTO:" & Xprom11
            End If
            data_person.Recordset.MoveNext
            If data_person.Recordset.EOF = False Then
               Xelcar = data_person.Recordset("cargod")
            End If
            
'            Xelcar = data_person.Recordset("cargod")
            Xprom11 = 0
            Xprom22 = 0
            Xcanpp = 0
         End If
         If pb1.Value < pb1.Max Then
            pb1.Value = pb1.Value + 1
         End If
      Loop
      frm_infeval.MousePointer = 0
      MsgBox "Terminado"
      cr1.ReportFileName = App.path & "\infprom2.rpt"
      cr1.Action = 1
   End If
End If


If Combo1.ListIndex = 1 Then
   
   frm_infeval.MousePointer = 11
   data_inff.RecordSource = "item"
   data_inff.Refresh
   If data_inff.Recordset.RecordCount > 0 Then
      data_inff.Recordset.MoveFirst
      Do While Not data_inff.Recordset.EOF
         data_inff.Recordset.Delete
         data_inff.Recordset.MoveNext
      Loop
   End If
'   If Combo3.ListIndex >= 0 Then
'      data_person.RecordSource = "Select * from personas where jefed ='" & Combo3.Text & "'"
'      data_person.Refresh
'   Else
'      data_person.RecordSource = "Select * from personas order by jefed"
'      data_person.Refresh
'   End If
''''''' BUSCAR POR EL ID2 PARA LOS REPETIDOS
   Data1.RecordSource = "Select * from evaluas where idjefe <> idempl order by id2,idpregun"
   Data1.Refresh
   Xst1 = 0
   Xst2 = 0
   Xst3 = 0
   Xst4 = 0
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveLast
      pb1.Max = Data1.Recordset.RecordCount
      pb1.Value = 0
      Data1.Recordset.MoveFirst
      Xlace = Data1.Recordset("id2")
      DoEvents
      Do While Not Data1.Recordset.EOF
         If Xlace = Data1.Recordset("id2") Then
            If Data1.Recordset("idpregun") = 1 Then
               Xp1 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 2 Then
               Xp2 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 3 Then
               Xp3 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 4 Then
               Xp4 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 5 Then
               Xp5 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 6 Then
               Xp6 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 7 Then
               Xp7 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 8 Then
               Xp8 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 9 Then
               Xp9 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 10 Then
               Xp10 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 11 Then
               Xp11 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 12 Then
               Xp12 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 13 Then
               Xp13 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 14 Then
               Xp14 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 15 Then
               Xp15 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 16 Then
               Xp16 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 17 Then
               Xp17 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 18 Then
               Xp18 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 19 Then
               Xp19 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 20 Then
               Xp20 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 21 Then
               Xp21 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 22 Then
               Xp22 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 23 Then
               Xp23 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 24 Then
               Xp24 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 25 Then
               Xp25 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 26 Then
               Xp26 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 27 Then
               Xp27 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 28 Then
               Xp28 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 29 Then
               Xp29 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 30 Then
               Xp30 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 31 Then
               Xp31 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 32 Then
               Xp32 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            Xlace = Data1.Recordset("id2")
            Data1.Recordset.MoveNext
         Else
            Data1.Recordset.MovePrevious
            If Combo3.ListIndex >= 0 Then
'      data_person.RecordSource = "Select * from personas where jefed ='" & Combo3.Text & "'"
               data_person.RecordSource = "Select * from personas where id2 =" & Data1.Recordset("id2") & " and jefed ='" & Combo3.Text & "'"
            Else
               data_person.RecordSource = "Select * from personas where id2 =" & Data1.Recordset("id2")
            End If
            data_person.Refresh
            If data_person.Recordset.RecordCount > 0 Then
               data_inff.Recordset.AddNew
               data_inff.Recordset("nombre") = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
               data_inff.Recordset("cargod") = data_person.Recordset("cargod")
               data_inff.Recordset("jefed") = data_person.Recordset("jefed")
               data_inff.Recordset("idempl") = Data1.Recordset("idempl")
               data_inff.Recordset("p1") = Xp1
               data_inff.Recordset("p2") = Xp2
               data_inff.Recordset("p3") = Xp3
               data_inff.Recordset("p4") = Xp4
               data_inff.Recordset("p5") = Xp5
               data_inff.Recordset("p6") = Xp6
               data_inff.Recordset("p7") = Xp7
               data_inff.Recordset("p8") = Xp8
               data_inff.Recordset("p9") = Xp9
               data_inff.Recordset("p10") = Xp10
               data_inff.Recordset("p11") = Xp11
               data_inff.Recordset("p12") = Xp12
               data_inff.Recordset("p13") = Xp13
               data_inff.Recordset("p14") = Xp14
               data_inff.Recordset("p15") = Xp15
               data_inff.Recordset("p16") = Xp16
               data_inff.Recordset("p17") = Xp17
               data_inff.Recordset("p18") = Xp18
               data_inff.Recordset("p19") = Xp19
               data_inff.Recordset("p20") = Xp20
               data_inff.Recordset("p21") = Xp21
               data_inff.Recordset("p22") = Xp22
               data_inff.Recordset("p23") = Xp23
               data_inff.Recordset("p24") = Xp24
               data_inff.Recordset("p25") = Xp25
               data_inff.Recordset("p26") = Xp26
               data_inff.Recordset("p27") = Xp27
               data_inff.Recordset("p28") = Xp28
               data_inff.Recordset("p29") = Xp29
               data_inff.Recordset("p30") = Xp30
               data_inff.Recordset("p31") = Xp31
               data_inff.Recordset("p32") = Xp32
               data_inff.Recordset("st1") = Xst1 / 8
               data_inff.Recordset("st2") = Xst2 / 8
               data_inff.Recordset("st3") = Xst3 / 8
               data_inff.Recordset("st4") = Xst4 / 8
               data_inff.Recordset("tot") = Xst1 + Xst2 + Xst3 + Xst4
               data_inff.Recordset("tot") = data_inff.Recordset("tot") / 32
               data_inff.Recordset.Update
               
               Xst1 = 0
               Xst2 = 0
               Xst3 = 0
               Xst4 = 0
            Else
               Xst1 = 0
               Xst2 = 0
               Xst3 = 0
               Xst4 = 0
            End If
            Data1.Recordset.MoveNext
            Xlace = Data1.Recordset("id2")
         End If
         If pb1.Value >= pb1.Max Then
         Else
           pb1.Value = pb1.Value + 1
         End If
      Loop
      frm_infeval.MousePointer = 0
      MsgBox "Proceso terminado"
'     If Combo3.ListIndex >= 0 Then
      cr1.ReportFileName = App.path & "\infitem2.rpt"
'      Else
'         cr1.ReportFileName = App.Path & "\infprom1t.rpt"
'      End If
      cr1.Action = 1
   End If
End If


If Combo1.ListIndex = 11 Then
   Dim Xpromm1, Xpromm2, Xpromm3, Xpromm4 As Integer
   Dim Xpromm11, Xpromm22, Xpromm33, Xpromm44, Xcanpreg, Xpromauto, Xpromeva As Integer
   Xcanpreg = 0
   Xpromauto = 0
   Xpromeva = 0
   frm_infeval.MousePointer = 11
   data_inff.RecordSource = "porpersona"
   data_inff.Refresh
   If data_inff.Recordset.RecordCount > 0 Then
      data_inff.Recordset.MoveFirst
      Do While Not data_inff.Recordset.EOF
         data_inff.Recordset.Delete
         data_inff.Recordset.MoveNext
      Loop
   End If
   If Combo3.ListIndex >= 0 Then
      data_person.RecordSource = "Select * from personas where jefed ='" & Combo3.Text & "'"
      data_person.Refresh
   Else
      data_person.RecordSource = "Select * from personas order by jefed"
      data_person.Refresh
   End If
''''''' BUSCAR POR EL ID2 PARA LOS REPETIDOS
   If data_person.Recordset.RecordCount > 0 Then
      data_person.Recordset.MoveLast
      pb1.Max = data_person.Recordset.RecordCount
      pb1.Value = 0
      data_person.Recordset.MoveFirst
      Xlace = data_person.Recordset("id2")
      DoEvents
      Do While Not data_person.Recordset.EOF
         Data1.RecordSource = "Select * from evaluas where id2 =" & data_person.Recordset("id2") & " and idjefe = idempl and idpregun not in (33,34,35,36,37,38,39,40) order by idpregun"
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
               If Data1.Recordset("idtitulo") = 1 Then
                  Xpromm1 = Xpromm1 + Data1.Recordset("puntos")
               End If
               If Data1.Recordset("idtitulo") = 2 Then
                  Xpromm2 = Xpromm2 + Data1.Recordset("puntos")
               End If
               If Data1.Recordset("idtitulo") = 3 Then
                  Xpromm3 = Xpromm3 + Data1.Recordset("puntos")
               End If
               If Data1.Recordset("idtitulo") = 4 Then
                  Xpromm4 = Xpromm4 + Data1.Recordset("puntos")
               End If
               Xpromauto = Xpromm1 + Xpromm2 + Xpromm3 + Xpromm4
               Data1.Recordset.MoveNext
               Xcanpreg = Xcanpreg + 1
            Loop
            DoEvents
         End If
         Data1.RecordSource = "Select * from evaluas where id2 =" & data_person.Recordset("id2") & " and idjefe <> idempl and idpregun not in (33,34,35,36,37,38,39,40) order by idpregun"
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
               If Data1.Recordset("idtitulo") = 1 Then
                  Xpromm11 = Xpromm11 + Data1.Recordset("puntos")
               End If
               If Data1.Recordset("idtitulo") = 2 Then
                  Xpromm22 = Xpromm22 + Data1.Recordset("puntos")
               End If
               If Data1.Recordset("idtitulo") = 3 Then
                  Xpromm33 = Xpromm33 + Data1.Recordset("puntos")
               End If
               If Data1.Recordset("idtitulo") = 4 Then
                  Xpromm44 = Xpromm44 + Data1.Recordset("puntos")
               End If
               Xpromeva = Xpromm11 + Xpromm22 + Xpromm33 + Xpromm44
               Data1.Recordset.MoveNext
               Xcanpreg = Xcanpreg + 1
            Loop
            DoEvents
         End If
         If Xpromm1 > 0 And Xpromm11 > 0 Then
            data_inff.Recordset.AddNew
            data_inff.Recordset("id") = data_person.Recordset("id2")
            data_inff.Recordset("nombre") = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
            data_inff.Recordset("cargo") = data_person.Recordset("cargod")
            data_inff.Recordset("jefe") = data_person.Recordset("jefed") 'aca
            
            data_inff.Recordset("promauto1") = Xpromm1 / 8
            data_inff.Recordset("promeva1") = Xpromm11 / 8
            data_inff.Recordset("tot1") = Xpromm1 + Xpromm11
            data_inff.Recordset("tot1") = data_inff.Recordset("tot1") / 16
            
            data_inff.Recordset("promauto2") = Xpromm2 / 8
            data_inff.Recordset("promeva2") = Xpromm22 / 8
            data_inff.Recordset("tot2") = Xpromm2 + Xpromm22
            data_inff.Recordset("tot2") = data_inff.Recordset("tot2") / 16
            
            data_inff.Recordset("promauto3") = Xpromm3 / 8
            data_inff.Recordset("promeva3") = Xpromm33 / 8
            data_inff.Recordset("tot3") = Xpromm3 + Xpromm33
            data_inff.Recordset("tot3") = data_inff.Recordset("tot3") / 16
                  
            data_inff.Recordset("promauto4") = Xpromm4 / 8
            data_inff.Recordset("promeva4") = Xpromm44 / 8
            data_inff.Recordset("tot4") = Xpromm4 + Xpromm44
            data_inff.Recordset("tot4") = data_inff.Recordset("tot4") / 16
                  
            data_inff.Recordset("totfin") = Xpromm1 + Xpromm11 + Xpromm2 + Xpromm22 + Xpromm3 + Xpromm33 + Xpromm4 + Xpromm44
            data_inff.Recordset("totfin") = data_inff.Recordset("totfin") / 64
            data_inff.Recordset("totauto") = Xpromauto / 32
            data_inff.Recordset("toteva") = Xpromeva / 32
            data_inff.Recordset.Update
         End If
         Xpromeva = 0
         Xpromauto = 0
         Xpromm1 = 0
         Xpromm11 = 0
         Xpromm2 = 0
         Xpromm22 = 0
         Xpromm3 = 0
         Xpromm33 = 0
         Xpromm4 = 0
         Xpromm44 = 0
         data_person.Recordset.MoveNext
         If pb1.Value < pb1.Max Then
            pb1.Value = pb1.Value + 1
         End If
      Loop
      frm_infeval.MousePointer = 0
      MsgBox "Terminado"
      If Check1.Value = 1 Then
         cr1.ReportFileName = App.path & "\infprom33.rpt"
      Else
         cr1.ReportFileName = App.path & "\infprom3.rpt"
      End If
      cr1.Action = 1
   End If
End If



Xlace = 0

If Combo1.ListIndex = 8 Then
   frm_infeval.MousePointer = 11
   data_inff.RecordSource = "item"
   data_inff.Refresh
   If data_inff.Recordset.RecordCount > 0 Then
      data_inff.Recordset.MoveFirst
      Do While Not data_inff.Recordset.EOF
         data_inff.Recordset.Delete
         data_inff.Recordset.MoveNext
      Loop
   End If
   If Combo3.ListIndex >= 0 Then
'      data_person.RecordSource = "Select * from personas where cargod ='" & Combo3.Text & "'"
'      data_person.Refresh
'      If data_person.Recordset.RecordCount > 0 Then
'         Xlacejef = data_person.Recordset("id")
'      Else
'         Xlacejef = 0
'      End If
'      Data1.RecordSource = "Select * from evaluas where idjefe =" & Xlacejef & " order by idempl,idpregun"
'      Data1.Refresh
      Data1.RecordSource = "Select * from evaluas where idjefe = idempl and idpregun not in (33,34,35,36,37,38,39,40) order by id2,idpregun"
      Data1.Refresh
   Else
      Data1.RecordSource = "Select * from evaluas where idjefe = idempl and idpregun not in (33,34,35,36,37,38,39,40) order by id2,idpregun"
      Data1.Refresh
   End If
''''''' BUSCAR POR EL ID2 PARA LOS REPETIDOS
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveLast
      pb1.Max = Data1.Recordset.RecordCount
      pb1.Value = 0
      Data1.Recordset.MoveFirst
      Xlace = Data1.Recordset("id2")
      DoEvents
      Do While Not Data1.Recordset.EOF
         If Xlace = Data1.Recordset("id2") Then
            If Data1.Recordset("idpregun") = 1 Then
               Xp1 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 2 Then
               Xp2 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 3 Then
               Xp3 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 4 Then
               Xp4 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 5 Then
               Xp5 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 6 Then
               Xp6 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 7 Then
               Xp7 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 8 Then
               Xp8 = Data1.Recordset("puntos")
               Xst1 = Xst1 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 9 Then
               Xp9 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 10 Then
               Xp10 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 11 Then
               Xp11 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 12 Then
               Xp12 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 13 Then
               Xp13 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 14 Then
               Xp14 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 15 Then
               Xp15 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 16 Then
               Xp16 = Data1.Recordset("puntos")
               Xst2 = Xst2 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 17 Then
               Xp17 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 18 Then
               Xp18 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 19 Then
               Xp19 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 20 Then
               Xp20 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 21 Then
               Xp21 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 22 Then
               Xp22 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 23 Then
               Xp23 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 24 Then
               Xp24 = Data1.Recordset("puntos")
               Xst3 = Xst3 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 25 Then
               Xp25 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 26 Then
               Xp26 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 27 Then
               Xp27 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 28 Then
               Xp28 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 29 Then
               Xp29 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 30 Then
               Xp30 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 31 Then
               Xp31 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            If Data1.Recordset("idpregun") = 32 Then
               Xp32 = Data1.Recordset("puntos")
               Xst4 = Xst4 + Data1.Recordset("puntos")
            End If
            Xlace = Data1.Recordset("idempl")
            Data1.Recordset.MoveNext
         Else
            Data1.Recordset.MovePrevious
            data_person.RecordSource = "Select * from personas where id2 =" & Data1.Recordset("id2")
            data_person.Refresh
            If data_person.Recordset.RecordCount > 0 Then
               data_inff.Recordset.AddNew
               data_inff.Recordset("nombre") = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
               data_inff.Recordset("cargod") = data_person.Recordset("cargod")
               data_inff.Recordset("jefed") = data_person.Recordset("jefed")
               data_inff.Recordset("idempl") = Data1.Recordset("id2")
               data_inff.Recordset("p1") = Xp1
               data_inff.Recordset("p2") = Xp2
               data_inff.Recordset("p3") = Xp3
               data_inff.Recordset("p4") = Xp4
               data_inff.Recordset("p5") = Xp5
               data_inff.Recordset("p6") = Xp6
               data_inff.Recordset("p7") = Xp7
               data_inff.Recordset("p8") = Xp8
               data_inff.Recordset("p9") = Xp9
               data_inff.Recordset("p10") = Xp10
               data_inff.Recordset("p11") = Xp11
               data_inff.Recordset("p12") = Xp12
               data_inff.Recordset("p13") = Xp13
               data_inff.Recordset("p14") = Xp14
               data_inff.Recordset("p15") = Xp15
               data_inff.Recordset("p16") = Xp16
               data_inff.Recordset("p17") = Xp17
               data_inff.Recordset("p18") = Xp18
               data_inff.Recordset("p19") = Xp19
               data_inff.Recordset("p20") = Xp20
               data_inff.Recordset("p21") = Xp21
               data_inff.Recordset("p22") = Xp22
               data_inff.Recordset("p23") = Xp23
               data_inff.Recordset("p24") = Xp24
               data_inff.Recordset("p25") = Xp25
               data_inff.Recordset("p26") = Xp26
               data_inff.Recordset("p27") = Xp27
               data_inff.Recordset("p28") = Xp28
               data_inff.Recordset("p29") = Xp29
               data_inff.Recordset("p30") = Xp30
               data_inff.Recordset("p31") = Xp31
               data_inff.Recordset("p32") = Xp32
               data_inff.Recordset("st1") = Xst1 / 8
               data_inff.Recordset("st2") = Xst2 / 8
               data_inff.Recordset("st3") = Xst3 / 8
               data_inff.Recordset("st4") = Xst4 / 8
               data_inff.Recordset("tot") = Xst1 + Xst2 + Xst3 + Xst4
               data_inff.Recordset("tot") = data_inff.Recordset("tot") / 32
               data_inff.Recordset.Update
               
               Xst1 = 0
               Xst2 = 0
               Xst3 = 0
               Xst4 = 0
               
            End If
            Data1.Recordset.MoveNext
            Xlace = Data1.Recordset("id2")
         End If
         If pb1.Value >= pb1.Max Then
         Else
           pb1.Value = pb1.Value + 1
         End If
      Loop
      frm_infeval.MousePointer = 0
      MsgBox "Proceso terminado"
      cr1.ReportFileName = App.path & "\infitem1.rpt"
      cr1.Action = 1
      
   End If
End If


Xlace = 0

If Combo1.ListIndex = 7 Then
   frm_infeval.MousePointer = 11
   data_inff.RecordSource = "infcomenta"
   data_inff.Refresh
   If data_inff.Recordset.RecordCount > 0 Then
      data_inff.Recordset.MoveFirst
      Do While Not data_inff.Recordset.EOF
         data_inff.Recordset.Delete
         data_inff.Recordset.MoveNext
      Loop
   End If
   If Combo3.ListIndex >= 0 Then
      data_person.RecordSource = "Select * from personas where cargod ='" & Combo3.Text & "'"
      data_person.Refresh
      If data_person.Recordset.RecordCount > 0 Then
         Xlacejef = data_person.Recordset("id")
      Else
         Xlacejef = 0
      End If
      Data1.RecordSource = "Select * from evaluas where idjefe =" & Xlacejef & " and obs is not null order by idempl,idpregun"
      Data1.Refresh
   Else
      Data1.RecordSource = "Select * from evaluas where obs is not null order by idempl,idpregun"
      Data1.Refresh
   End If
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Xlace = Data1.Recordset("idempl")
      Do While Not Data1.Recordset.EOF
         If Xlace = Data1.Recordset("idempl") Then
            If Data1.Recordset("idpregun") > 1 And Data1.Recordset("idpregun") <= 8 Then
               If Xcomenta1 = "" Then
                  Xcomenta1 = Trim(str(Data1.Recordset("idpregun"))) & "." & Data1.Recordset("obs")
               Else
                  Xcomenta1 = Xcomenta1 & Chr(13) & Chr(10) & Trim(str(Data1.Recordset("idpregun"))) & "." & Data1.Recordset("obs")
               End If
            Else
               If Data1.Recordset("idpregun") > 8 And Data1.Recordset("idpregun") <= 16 Then
                  If Xcomenta2 = "" Then
                     Xcomenta2 = Trim(str(Data1.Recordset("idpregun"))) & "." & Data1.Recordset("obs")
                  Else
                     Xcomenta2 = Xcomenta2 & Chr(13) & Chr(10) & Trim(str(Data1.Recordset("idpregun"))) & "." & Data1.Recordset("obs")
                  End If
               Else
                  If Data1.Recordset("idpregun") > 16 And Data1.Recordset("idpregun") <= 24 Then
                     If Xcomenta3 = "" Then
                        Xcomenta3 = Trim(str(Data1.Recordset("idpregun"))) & "." & Data1.Recordset("obs")
                     Else
                        Xcomenta3 = Xcomenta3 & Chr(13) & Chr(10) & Trim(str(Data1.Recordset("idpregun"))) & "." & Data1.Recordset("obs")
                     End If
                  Else
                     If Data1.Recordset("idpregun") > 25 And Data1.Recordset("idpregun") <= 32 Then
                        If Xcomenta4 = "" Then
                           Xcomenta4 = Trim(str(Data1.Recordset("idpregun"))) & "." & Data1.Recordset("obs")
                        Else
                           Xcomenta4 = Xcomenta4 & Chr(13) & Chr(10) & Trim(str(Data1.Recordset("idpregun"))) & "." & Data1.Recordset("obs")
                        End If
                     End If
                  End If
               End If
            End If
            Xlace = Data1.Recordset("idempl")
            Data1.Recordset.MoveNext
         Else
            Data1.Recordset.MovePrevious
            data_inff.Recordset.AddNew
            data_inff.Recordset("obs1") = Xcomenta1
            data_inff.Recordset("obs2") = Xcomenta2
            data_inff.Recordset("obs3") = Xcomenta3
            data_inff.Recordset("obs4") = Xcomenta4
            data_inff.Recordset("idgpo") = Data1.Recordset("idtitulo")
            data_inff.Recordset("idempl") = Data1.Recordset("idempl")
            data_person.RecordSource = "Select * from personas where id =" & Data1.Recordset("idempl")
            data_person.Refresh
            If data_person.Recordset.RecordCount > 0 Then
               data_inff.Recordset("nombre") = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
               data_inff.Recordset("cargo") = data_person.Recordset("cargod")
               data_inff.Recordset("cargoid") = data_person.Recordset("cargo")
               data_inff.Recordset("jefe") = data_person.Recordset("jefed")
            End If
            data_inff.Recordset.Update
            Data1.Recordset.MoveNext
            Xlace = Data1.Recordset("idempl")
            Xcomenta1 = ""
            Xcomenta2 = ""
            Xcomenta3 = ""
            Xcomenta4 = ""
         End If
      Loop
      frm_infeval.MousePointer = 0
      MsgBox "Proceso terminado"
      cr1.ReportFileName = App.path & "\infcoment.rpt"
      cr1.Action = 1
      
   Else
      frm_infeval.MousePointer = 0
      MsgBox "No existen registros"
   End If
   frm_infeval.MousePointer = 0
End If

If Combo1.ListIndex = 6 Then
'   Data1.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & Combo2.Text & "' and idtitulo =" & 1 & " and idjefe =" & Wxeljefeid & " order by idpregun"
   
   data_us.RecordSource = "Select * from cap_ciap where des_cap ='" & WElusuario & "'"
   data_us.Refresh
   If data_us.Recordset.RecordCount > 0 Then
      data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
      data_buscap.Refresh
      If data_buscap.Recordset.RecordCount > 0 Then
         Wxeljefeid = data_buscap.Recordset("id")
      Else
         MsgBox "No se encontró el usuario, comunique al administrador"
         frm_infeval.MousePointer = 0
         Unload Me
      End If
   Else
      MsgBox "No se encuentra usuario registrado, comunique al administrador", vbInformation
      frm_infeval.MousePointer = 0
      Unload Me
   End If
   If IsNull(data_buscap.Recordset("cargod")) = False Then
      data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_buscap.Recordset("cargod") & "'"
      data_cargo.Refresh
      If data_cargo.Recordset("tipo") = 2 Or data_cargo.Recordset("tipo") = 3 Or WElusuario = "DARIOH" Or WElusuario = "JFERNAN" Then
         If data_cargo.Recordset("tipo") = 2 And WElusuario <> "JFERNAN" Then
            data_person.RecordSource = "Select * from personas where jefed ='" & data_cargo.Recordset("descrip") & "' or cargod ='" & data_cargo.Recordset("descrip") & "' order by id"
            data_person.Refresh
         Else
            data_person.RecordSource = "Select * from personas"
            data_person.Refresh
         End If
      Else
         MsgBox "Usuario no registrado para informes"
         Unload Me
      End If
      If data_person.Recordset.RecordCount > 0 Then
         data_person.Recordset.MoveFirst
         Do While Not data_person.Recordset.EOF
            Data1.RecordSource = "Select * from evaluas where idempl =" & data_person.Recordset("id")
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.MoveLast
               If Data1.Recordset.RecordCount >= 64 Then
               Else
                  If Data1.Recordset.RecordCount >= 32 Then
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_cedula") = Data1.Recordset("idempl")
                     Xnombb = data_person.Recordset("nom1") & " " & data_person.Recordset("ape1")
                     data_inf.Recordset("cl_apellid") = Mid(Xnombb, 1, 60)
                     data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
                     data_inf.Recordset("cl_nombre") = Data1.Recordset("periodo")
                     data_inf.Recordset("cl_nomcobr") = Mid(data_person.Recordset("cargod"), 1, 20)
                     If IsNull(Data1.Recordset("firma")) = False Then
                        If Data1.Recordset("firma") = 5 Then
                           data_inf.Recordset("cl_codconv") = "SI"
                        Else
                           data_inf.Recordset("cl_codconv") = "NO"
                        End If
                     Else
                        data_inf.Recordset("cl_codconv") = "NO"
                     End If
                     data_inf.Recordset.Update
                  End If
               End If
            End If
            data_person.Recordset.MoveNext
         Loop
         frm_infeval.MousePointer = 0
         MsgBox "Terminado"
         cr1.ReportFileName = App.path & "\infperevalj2.rpt"
         cr1.Action = 1
      End If
   Else
      MsgBox "No se encuentra cargo para informes"
      frm_infeval.MousePointer = 0
      Unload Me
      
   End If
End If



Combo2.Clear

data_perio.Connect = "ODBC;DSN=sappper;"
data_perio.RecordSource = "periodo"
data_perio.Refresh
If data_perio.Recordset.RecordCount > 0 Then
   data_perio.Recordset.MoveFirst
   Do While Not data_perio.Recordset.EOF
      Combo2.AddItem data_perio.Recordset("descrip")
      data_perio.Recordset.MoveNext
   Loop
   Combo2.AddItem "2016"
   Combo2.ListIndex = 0
   
End If

Command1.Enabled = True
Command2.Enabled = True
frm_infeval.MousePointer = 0

'Exit Sub

'Errevalinf:
'           If Err.Number = 3155 Then
'              frm_infeval.MousePointer = 0
'              MsgBox "Error al grabar"
'              Command1.Enabled = True
'              Command2.Enabled = True
'           Else
'              frm_infeval.MousePointer = 0
'              MsgBox "Hay un error en el informe, cierre la pantalla y vuelva a intentar."
'              Command1.Enabled = True
'              Command2.Enabled = True
'           End If
           

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf2.DatabaseName = App.path & "\informes.mdb"

data_inff.DatabaseName = App.path & "\infcomen.mdb"


Data1.Connect = "ODBC;DSN=sappper;"

data_us.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_persontot.Connect = "ODBC;DSN=sappper;"

data_cargo.Connect = "ODBC;DSN=sappper;"

data_buscap.Connect = "ODBC;DSN=sappper;"

data_pregunt.Connect = "ODBC;DSN=sappper;"

data_person.Connect = "ODBC;DSN=sappper;"

data_jefe.Connect = "ODBC;DSN=sappper;"
data_jefe.RecordSource = "jefes"
data_jefe.Refresh

Combo3.Clear
If data_jefe.Recordset.RecordCount > 0 Then
   data_jefe.Recordset.MoveFirst
   Do While Not data_jefe.Recordset.EOF
      Combo3.AddItem data_jefe.Recordset("descrip")
      data_jefe.Recordset.MoveNext
   Loop
End If
Combo2.Clear

data_perio.Connect = "ODBC;DSN=sappper;"
data_perio.RecordSource = "periodo"
data_perio.Refresh
If data_perio.Recordset.RecordCount > 0 Then
   data_perio.Recordset.MoveFirst
   Do While Not data_perio.Recordset.EOF
      Combo2.AddItem data_perio.Recordset("descrip")
      data_perio.Recordset.MoveNext
   Loop
   Combo2.AddItem "2016"
   Combo2.ListIndex = 0
   
End If

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
