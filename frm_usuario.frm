VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_usuario 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Ingreso al sistema"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_versionlocal 
      Caption         =   "data_versionlocal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc adoversion 
      Height          =   330
      Left            =   3360
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      Caption         =   "adoversion"
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4560
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data_usuario 
      Height          =   330
      Left            =   1560
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
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
      Caption         =   "data_usuario"
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Conectar remoto"
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
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data data_ctrl 
      Caption         =   "data_ctrl"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_ctrol 
      Caption         =   "data_ctrol"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\usapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarioact"
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton btn_can 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      MouseIcon       =   "frm_usuario.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frm_usuario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton btn_acep 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MaskColor       =   &H0080FF80&
      MouseIcon       =   "frm_usuario.frx":0894
      MousePointer    =   99  'Custom
      Picture         =   "frm_usuario.frx":0B9E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txt_pass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txt_usua 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frm_usuario.frx":1128
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "CONTRASEÑA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "USUARIO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "SEGURIDAD DEL SISTEMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   240
      Picture         =   "frm_usuario.frx":156A
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "frm_usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub btn_acep_Click()
Dim Xdias As Long
Dim lafecha As Date
If Check1.Value = 1 Then
   Xipsrv = "201.221.0.76"
   Xconexrmt = "sappnewrmt"
Else
 '''  Xipsrv = "192.168.10.50"
   Xipsrv = "192.168.10.134"
''   Xipsrv = "localhost"
   Xconexrmt = "sappnew"
End If
data_usuario.ConnectionString = "dsn=" & Xconexrmt
'data_usuario.Recordset.FindFirst "usuario = '" & txt_usua.Text & "'"
data_usuario.RecordSource = "Select * from usuarios where usuario ='" & txt_usua.Text & "'"
data_usuario.Refresh
If data_usuario.Recordset.RecordCount > 0 Then
   lafecha = adoversion.Recordset("fecha")
'   If data_versionlocal.Recordset("ultfecact") < lafecha Then
'       MsgBox "Existen actualizaciones del sistema para descargar."
'       btn_acep.Enabled = False
'       btn_can.Enabled = False
'       Command1_Click
'   End If
   If data_usuario.Recordset("clave") = txt_pass.Text Then
      If data_usuario.Recordset("tipo") = "ADMINISTRADOR" Then
         data_ctrol.Recordset.Edit
         data_ctrol.Recordset("nombre") = data_usuario.Recordset("usuario")
         data_ctrol.Recordset("nomb2") = data_usuario.Recordset("usuario")
         data_ctrol.Recordset.Update
      Else
         data_ctrol.Recordset.Edit
         data_ctrol.Recordset("nombre") = data_usuario.Recordset("usuario")
         data_ctrol.Recordset.Update
      End If
      WElusuario = UCase(txt_usua.Text)
      XWeltipoU = data_usuario.Recordset("tipo")
      Welnombredu = data_usuario.Recordset("nombre")
      Welnrou = data_usuario.Recordset("id")
      WxclaveU = txt_pass.Text
      Xcolesp = 0 ' para anotacion especialistas
'      Unload Me
      frm_usuario.Hide
      If data_ctrl.Recordset("base") = 100 Then
         Xdias = DateDiff("d", data_ctrl.Recordset("fecha"), Date)
         If Xdias > 40 Then
            MsgBox "SE HA VENCIDO EL PLAZO PARA EL MANTENIMIENTO DE SU NOTEBOOK. EL SISTEMA NO SE PUEDE EJECUTAR!", vbCritical
            End
         Else
            If Xdias > 30 Then
               MsgBox "ESTÁ VENCIDO EL PLAZO PARA ENVIAR SU NOTEBOOK PARA MATENIMIENTO. COORDINE CON INFORMÁTICA!!", vbCritical
               frm_menu.Show
            Else
               If Xdias > 25 Then
                  MsgBox "FALTAN 5 DÍAS PARA ENVIAR SU NOTEBOOK A RESPALDAR. COORDINE CON INFORMÁTICA", vbInformation
                  frm_menu.Show
               Else
                  frm_menu.Show
               End If
            End If
         End If
      Else
         frm_menu.Show
      End If
   Else
      MsgBox "Clave incorrecta", vbInformation, "Seguridad"
      txt_pass.SetFocus
   End If
Else
   MsgBox "Usuario no registrado", vbInformation, "Seguridad"
   txt_usua.SetFocus
End If

End Sub

Private Sub btn_can_Click()

End

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   MsgBox "Marcando ésta opción se conectará por internet", vbInformation
End If

End Sub

Private Sub Command1_Click()

 'Dim El_Host As String
 'Dim Xlocal, Xremoto As String
 'Xlocal = App.path & "\SAPP\sapp.zip"
 'Xremoto = "sapp.zip"
'frm_usuario.MousePointer = 11
'    El_Host = "192.168.10.22"

'    With Inet1
'        .url = "http://192.168.10.22"
'        .Protocol = icFTP
'        .RemoteHost = "192.168.10.22"
'        .UserName = "Administrator"
'        .PassWord = "Sapp1987"
'        .Execute , "Get " + Xremoto + " " + Xlocal
'        Do While .StillExecuting
'           DoEvents
'        Loop
'     End With
'     Inet1.Execute , "quit"
'     frm_usuario.MousePointer = 0
'     MsgBox "Descarga de archivo terminado. Se actualizará el sistema.", vbInformation
'     Shell (App.path & "\actualiza.exe"), vbNormalFocus
'     End


End Sub


Private Sub Form_Load()
'Set Cnn = New ADODB.Connection
'Cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51; " & _
'"Data Source=" & sBase & ";" & _
'"Jet OLEDB:Database Password=laclave"
adoversion.ConnectionString = "dsn=sappnew"
adoversion.RecordSource = "select * from version"
adoversion.Refresh

data_ctrl.DatabaseName = App.path & "\ctrabre.mdb"
data_ctrl.RecordSource = "ctrabre"
data_ctrl.Refresh

If App.PrevInstance = True Then
   MsgBox "Ya está abierto el programa SAPP", vbCritical
   End
End If
data_versionlocal.DatabaseName = App.path & "\ctrf.mdb"
data_versionlocal.RecordSource = "ctrf"
data_versionlocal.Refresh



End Sub

Private Sub Form_Resize()
With Image2
   .Left = 0
   .Top = 0
   .Height = Me.Height
   .Width = Me.Width
End With

End Sub

Private Sub txt_pass_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(VBA.UCase(VBA.Chr(KeyAscii)))

If KeyAscii = 13 Then
   KeyAscii = 0
   btn_acep.SetFocus
   btn_acep_Click
End If

End Sub

Private Sub txt_usua_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(VBA.UCase(VBA.Chr(KeyAscii)))

If KeyAscii = 13 Then
   KeyAscii = 0
   txt_pass.SetFocus
End If

End Sub



