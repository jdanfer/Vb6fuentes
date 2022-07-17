VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_entsale 
   BackColor       =   &H00C0C000&
   Caption         =   "REGISTROS DE ENTRADAS Y SALIDAS"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8070
   Icon            =   "frm_entsale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_medicos 
      Height          =   735
      Left            =   2520
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      Caption         =   "data_medicos"
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
   Begin MSAdodcLib.Adodc data_medicos2 
      Height          =   735
      Left            =   360
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
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
      Caption         =   "data_medicos2"
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
   Begin MSAdodcLib.Adodc data_hor 
      Height          =   735
      Left            =   3240
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
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
      Caption         =   "data_hor"
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
   Begin MSAdodcLib.Adodc data_hor2 
      Height          =   735
      Left            =   3120
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
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
      Caption         =   "data_hor2"
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
   Begin VB.TextBox t_movil 
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
      Left            =   3000
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox t_codmsapp 
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox t_nom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   4815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1620
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   6495
   End
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      ItemData        =   "frm_entsale.frx":0442
      Left            =   5040
      List            =   "frm_entsale.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Picture         =   "frm_entsale.frx":0461
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Graba los datos ingresaados"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox t_ced 
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
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   5
      ToolTipText     =   "Ingrese todos los números (para cédula 3480884-4 digitar 34808844"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7440
      Top             =   3240
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   120
      Picture         =   "frm_entsale.frx":09EB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MÓVIL/BASE:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Entrada/Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INGRESE SU DOCUMENTO:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label labhora 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hora actual:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label labfec 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha Actual:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   6840
      Picture         =   "frm_entsale.frx":10AED
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frm_entsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xcedtot As String

Command1.Enabled = False
If t_ced.Text <> "" And t_nom.Text <> "" And t_movil.Text <> "" Then
   Xcedtot = Trim(t_ced.Text)
   If t_codmsapp.Text = "" Then
      data_hor.RecordSource = "Select * from hc_archotro where hc_mat =" & t_ced.Text & " order by hc_fecha"
      data_hor.Refresh
   Else
      data_hor.RecordSource = "Select * from hc_archotro where hc_mat =" & t_codmsapp.Text & " order by hc_fecha"
      data_hor.Refresh
   End If
   data_par.Recordset.Edit
   data_par.Recordset("nro_reg") = data_par.Recordset("nro_reg") + 1
   data_par.Recordset.Update
   data_par.Refresh
   data_hor2.RecordSource = "Select * from hc_viaae where id =" & data_par.Recordset("nro_reg")
   data_hor2.Refresh
   
   data_hor.Recordset.AddNew
   data_hor.Recordset("id") = data_par.Recordset("nro_reg")
   If t_codmsapp.Text = "" Then
      data_hor.Recordset("hc_mat") = t_ced.Text
   Else
      data_hor.Recordset("hc_mat") = t_codmsapp.Text
   End If
   data_hor.Recordset("hc_nro") = Combo1.ListIndex
   data_hor.Recordset("hc_fecha") = Format(labfec.Caption, "dd/mm/yyyy")
   data_hor.Recordset("hc_hora") = labhora.Caption
   data_hor.Recordset("hc_descrip") = Xcedtot
   data_hor.Recordset("hc_lugar") = t_nom.Text
   data_hor.Recordset.Update
   data_hor2.Recordset.AddNew
   data_hor2.Recordset("id") = data_par.Recordset("nro_reg")
   data_hor2.Recordset("hc_cod") = t_movil.Text
   data_hor2.Recordset.Update
   data_hor2.Refresh
   
   data_hor.Refresh
   t_ced.Text = ""
   t_nom.Text = ""
   t_codmsapp.Text = ""
   List1.Clear
   data_hor.Recordset.MoveFirst
   Do While Not data_hor.Recordset.EOF
      If IsNull(data_hor.Recordset("hc_nro")) = False Then
         If data_hor.Recordset("hc_nro") = 0 Then
            EntoSale = "ENTRADA"
         Else
            EntoSale = "SALIDA"
         End If
      Else
         EntoSale = "ENTRADA"
      End If
      List1.AddItem data_hor.Recordset("hc_fecha") & " | " & data_hor.Recordset("hc_hora") & " | " & EntoSale
      data_hor.Recordset.MoveNext
   Loop
Else
   MsgBox "Verifique si están todos los datos"
End If
Command1.Enabled = True



End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then
   MsgBox "Ya está abierto el programa", vbCritical
   End
End If
labfec.Caption = Format(Date, "dd/mm/yyyy")
Combo1.ListIndex = 0
'data_medicos.Connect = "ODBC;DSN=sappnew;"
data_medicos.ConnectionString = "dsn=sappnew"
'data_medicos2.Connect = "ODBC;DSN=sappnew;"
data_medicos2.ConnectionString = "dsn=sappnew"
'data_hor.Connect = "ODBC;DSN=sappnew;"
data_hor.ConnectionString = "dsn=sappnew"
'data_hor2.Connect = "ODBC;DSN=sappnew;"
data_hor2.ConnectionString = "dsn=sappnew"
data_par.DatabaseName = App.Path & "\paramhoras.mdb"
data_par.RecordSource = "parsec0"
data_par.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_ced_LostFocus()
Dim Lacemed As String
Dim EntoSale As String
List1.Clear
If t_ced.Text <> "" Then
   Lacemed = Trim(t_ced.Text)
   data_medicos.RecordSource = "Select * from meta_tres where m_nrofrm ='" & Lacemed & "'"
   data_medicos.Refresh
   If data_medicos.Recordset.RecordCount > 0 Then
      data_medicos2.RecordSource = "Select * from medicos where med_cod =" & data_medicos.Recordset("m_mat")
      data_medicos2.Refresh
      If data_medicos2.Recordset.RecordCount > 0 Then
         t_codmsapp.Text = data_medicos.Recordset("m_mat")
         If IsNull(data_medicos2.Recordset("med_nombre")) = False Then
            t_nom.Text = data_medicos2.Recordset("med_nombre")
         Else
            t_nom.Text = ""
         End If
         If t_codmsapp.Text = "" Then
            data_hor.RecordSource = "Select * from hc_archotro where hc_mat =" & t_ced.Text & " order by hc_fecha"
            data_hor.Refresh
         Else
            data_hor.RecordSource = "Select * from hc_archotro where hc_mat =" & t_codmsapp.Text & " order by hc_fecha"
            data_hor.Refresh
         End If
         If data_hor.Recordset.RecordCount > 0 Then
            data_hor.Recordset.MoveFirst
            Do While Not data_hor.Recordset.EOF
               If IsNull(data_hor.Recordset("hc_nro")) = False Then
                  If data_hor.Recordset("hc_nro") = 0 Then
                     EntoSale = "ENTRADA"
                  Else
                     EntoSale = "SALIDA"
                  End If
               Else
                  EntoSale = "ENTRADA"
               End If
               List1.AddItem data_hor.Recordset("hc_fecha") & " | " & data_hor.Recordset("hc_hora") & " | " & EntoSale
               data_hor.Recordset.MoveNext
            Loop
         End If
      Else
         t_nom.Text = ""
      End If
   Else
      t_nom.Text = ""
   End If
Else
   t_nom.Text = ""
End If

End Sub



Private Sub t_movil_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_movil.SetFocus
End If

End Sub

Private Sub Timer1_Timer()
labhora.Caption = Format(Time, "HH:mm:ss")
labfec.Caption = Format(Date, "dd/mm/yyyy")

End Sub
