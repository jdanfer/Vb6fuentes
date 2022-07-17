VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_inicia 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Inicio de SAPP"
   ClientHeight    =   2760
   ClientLeft      =   1560
   ClientTop       =   2775
   ClientWidth     =   5925
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frm_inicia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2760
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4920
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5280
      Top             =   1200
   End
   Begin VB.Data data_rec 
      Caption         =   "data_rec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_env 
      Caption         =   "data_env"
      Connect         =   "dBASE IV;"
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
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2160
      Top             =   1080
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   1680
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   2160
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock wincli 
      Left            =   4200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   888
   End
   Begin MSWinsockLib.Winsock sockserv 
      Left            =   3360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   888
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picGancho 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "frm_inicia.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Menu mnuBar 
      Caption         =   ""
      Enabled         =   0   'False
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu mnuAcerca 
         Caption         =   "Programa SAPP"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frm_inicia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  
'Constantes para determinar que tipo de Red estamos conectados
  
Const NETWORK_ALIVE_AOL = &H4
Const NETWORK_ALIVE_LAN = &H1
Const NETWORK_ALIVE_WAN = &H2
  
'Función Api IsNetworkAlive para detectar _
 si estamos conectados y a que tipo de red
Private Declare Function IsNetworkAlive Lib "SENSAPI.DLL" ( _
    ByRef lpdwFlags As Long) As Long
  
Public Xcontar As Integer


' El PictureBox picGancho sirve como gancho de los
' mensajes CallBack del API Shell_NotifyIcon. Tiene
' que ser un control con un hWnd. Todo lo interesante
' esta en el picGancho_MouseMove . Como pueden ver, un
' control MsgHook o MsgBlaster aqui sobra...
'---------------
Private Type TIPONOTIFICARICONO
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'------------------
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
'--------------------
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    pnid As TIPONOTIFICARICONO) As Boolean
'--------------------
Private Declare Function WinExec& Lib "kernel32" _
    (ByVal lpCmdLine As String, ByVal nCmdShow As Long)
'--------------------
Dim t As TIPONOTIFICARICONO

Private Sub Command1_Click()
MsgBox "Que desea hacer?", vbOKCancel

End Sub

Private Sub Command2_Click()
'Me.Hide
    Dim ret As Long
  
    'Si la Api retorna 0 quiere decir que no hay ningun tipo de conexión de Red
        If IsNetworkAlive(ret) = 0 Then
  
            MsgBox "El sistema no está conectado a una NetWork!", vbInformation
  
        Else
            ' hay conexión , y muestra el tipo
            MsgBox "El sistema está conectado a: " + _
                   IIf(ret = NETWORK_ALIVE_AOL, "AOL", _
                   IIf(ret = NETWORK_ALIVE_LAN, "LAN", "WAN")) + " network", vbInformation
  
    End If

End Sub

Private Sub Form_Click()
'    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then  ' Como tener "Cancel"
'        Me.Hide
    End If
End Sub

Private Sub Form_Load()
'Kill ("C:\Datos\env_cli.mdb")

    If App.PrevInstance Then
        mnuAcerca_Click
        Unload Me
        End
    End If
'---------------------------------
    
    t.cbSize = Len(t)
    t.hwnd = picGancho.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
'---------------------------------
    t.szTip = "Programa SAPP iniciado" & Chr$(0) ' Es un string de "C" ( \0 )
    Shell_NotifyIcon NIM_ADD, t
    Me.Hide
    App.TaskVisible = False
    data_parsec.DatabaseName = App.Path & "\parsec0.mdb"
    data_parsec.RecordSource = "parsec0"
    data_parsec.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hwnd = picGancho.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub


Private Sub Form_Unload(Cancel As Integer)

    End
End Sub

Private Sub mnuAcerca_Click()
' Un consejo, mover un Form en estado minimizado
' da un GPF...
Dim ValDev As Long
'With frm_reinicia
'    picGancho.Picture = Me.Icon
'    Top = Screen.Height / 2 - Height / 2
'    Left = Screen.Width / 2 - Width / 2
'    Show
'End With
Shell (App.Path & "\sapp.exe"), vbNormalFocus

End Sub
Private Sub mnuSalir_Click(Index As Integer)
'    Unload Me

MsgBox "No puede cerrar el programa", vbInformation

End

End Sub
Private Sub picGancho_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, Msg As Long, ValDev As Long
    Msg = X / Screen.TwipsPerPixelX

    If rec = False Then
        rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                ValDev = WinExec(App.Path & "\sapp.exe", 1)
'                 ValDev = Form1.Show
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                 ' PopUp menu,2 significa Izq/Der botones en el menu, mnuAbout es BOLD
                 Me.PopupMenu mnuBar, 2, , , mnuAcerca
            End Select
        rec = False
    End If
End Sub

Private Sub sockserv_ConnectionRequest(ByVal requestID As Long)
sockserv.Close
sockserv.Accept requestID

End Sub

Private Sub Timer1_Timer()
Xcontar = Xcontar + 1
If Xcontar = 2 Then
    Dim ret As Long
  
    'Si la Api retorna 0 quiere decir que no hay ningun tipo de conexión de Red
        If IsNetworkAlive(ret) = 0 Then
  
'            MsgBox "El sistema no está conectado a una NetWork!", vbInformation
           Timer1.Enabled = False
           Timer5.Enabled = True
        Else
            ' hay conexión , y muestra el tipo
'            MsgBox "El sistema está conectado a: " + _
'                   IIf(Ret = NETWORK_ALIVE_AOL, "AOL", _
'                   IIf(Ret = NETWORK_ALIVE_LAN, "LAN", "WAN")) + " network", vbInformation
  
           Timer1.Enabled = False
           Timer2.Enabled = True
        
        End If

'   MsgBox "Comienza", vbCritical, "Mensaje"
   
   
'   Text2.Text = sockserv.LocalIP
'   sockserv.Close
'   sockserv.Listen
   Xcontar = 0


'   MsgBox "Esta ok "
End If

End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Private Sub Timer2_Timer()
Dim Fecenv As Date
Dim Xdias, Xdiaste As Long
Dim XCanD As Integer
Dim XFecrecib, XFecenvte As Date

' Preparar en el archivo ENVIOS los datos para enviar
'----------------------------------------------------
FileCopy "c:\datos\vacios\env_clia.dbf", "C:\datos\env_clia.dbf"

FileCopy "c:\datos\vacios\env_clib.dbf", "C:\datos\env_clib.dbf"

FileCopy "c:\datos\vacios\env_clim.dbf", "C:\datos\env_clim.dbf"

FileCopy "c:\datos\vacios\env_lin.dbf", "C:\datos\env_lin.dbf"

FileCopy "c:\datos\vacios\env_caja.dbf", "C:\datos\env_caja.dbf"

FileCopy "c:\datos\vacios\env_abm.dbf", "C:\datos\env_abm.dbf"

FileCopy "c:\datos\vacios\env_tes.dbf", "C:\datos\env_tes.dbf"

FileCopy "c:\datos\vacios\env_lla.dbf", "C:\datos\env_lla.dbf"

FileCopy "c:\datos\vacios\env_conv.dbf", "C:\datos\env_conv.dbf"

FileCopy "c:\datos\vacios\env_estu.dbf", "C:\datos\env_estu.dbf"

FileCopy "c:\datos\vacios\env_arq.dbf", "C:\datos\env_arq.dbf"
    
FileCopy "c:\datos\vacios\env_cob.dbf", "C:\datos\env_cob.dbf"

FileCopy "c:\datos\vacios\env_codc.dbf", "C:\datos\env_codc.dbf"

FileCopy "c:\datos\vacios\env_rubt.dbf", "C:\datos\env_rubt.dbf"

If data_parsec.Recordset("base") <> 3 Then
   Fecenv = data_parsec.Recordset("ult_env") + 1
   Xdias = Date - data_parsec.Recordset("ult_env")
   XCanD = 1
Else
   Fecenv = data_parsec.Recordset("ult_env") + 1
   Xdias = Date - data_parsec.Recordset("ult_env")
   XFecenvte = data_parsec.Recordset("ult_env3") + 1
   Xdiaste = Date - data_parsec.Recordset("ult_env3")
   XFecrecib = data_parsec.Recordset("ult_env") + 1
End If

If Xdias > 1 Then
    data_env.DatabaseName = "c:\Datos"
    data_env.RecordSource = "env_clia"
    data_env.Refresh
    XLafecha = Fecenv
    If data_env.Recordset.RecordCount > 0 Then
       data_env.Recordset.MoveFirst
       Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
       Loop
    End If
    data_cli.DatabaseName = App.Path & "\sapp.mdb"
    data_cli.RecordSource = "caja"
    data_cli.Refresh
    data_env.RecordSource = "env_caja"
    data_env.Refresh
    If data_env.Recordset.RecordCount > 0 Then
       data_env.Recordset.MoveFirst
       Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
       Loop
    End If
    If data_parsec.Recordset("base") = 3 Then
       data_cli.RecordSource = "Select * from caja where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "# and base =" & 3
       data_cli.Refresh
    End If
    If data_cli.Recordset.RecordCount > 0 Then
       data_cli.Recordset.MoveFirst
       Do While Not data_cli.Recordset.EOF
           data_env.Recordset.AddNew
            data_env.Recordset("fecha") = data_cli.Recordset("fecha")
            data_env.Recordset("moneda") = data_cli.Recordset("moneda")
            data_env.Recordset("numero") = data_cli.Recordset("numero")
            data_env.Recordset("nombre") = data_cli.Recordset("nombre")
            data_env.Recordset("movimiento") = data_cli.Recordset("movimiento")
            data_env.Recordset("imp_fact") = data_cli.Recordset("imp_fact")
            data_env.Recordset("observ") = data_cli.Recordset("observ")
            data_env.Recordset("saldo") = data_cli.Recordset("saldo")
            data_env.Recordset("usuario") = data_cli.Recordset("usuario")
            data_env.Recordset("hora") = data_cli.Recordset("hora")
            data_env.Recordset("saldo_user") = data_cli.Recordset("saldo_user")
            data_env.Recordset("base") = data_cli.Recordset("base")
            data_env.Recordset("cod_serv") = data_cli.Recordset("cod_serv")
            data_env.Recordset("nom_serv") = data_cli.Recordset("nom_serv")
            data_env.Recordset("cod_socio") = data_cli.Recordset("cod_socio")
            data_env.Recordset("nom_socio") = data_cli.Recordset("nom_socio")
            data_env.Recordset("documento") = data_cli.Recordset("documento")
            data_env.Recordset("caja_mesp") = data_cli.Recordset("caja_mesp")
            data_env.Recordset("caja_anop") = data_cli.Recordset("caja_anop")
            data_env.Recordset("imp_iva") = data_cli.Recordset("imp_iva")
            data_env.Recordset("opiva") = data_cli.Recordset("opiva")
            data_env.Recordset.Update
            data_cli.Recordset.MoveNext
       Loop
    End If
    data_env.RecordSource = "env_lin"
    data_env.Refresh
    If data_env.Recordset.RecordCount > 0 Then
       data_env.Recordset.MoveFirst
       Do While Not data_env.Recordset.EOF
          data_env.Recordset.Delete
          data_env.Recordset.MoveNext
       Loop
    End If
    If data_parsec.Recordset("base") = 3 Then
       data_cli.RecordSource = "Select * from linmmdd where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "# And base =" & 3
       data_cli.Refresh
    Else
       data_cli.RecordSource = "Select * from linmmdd where fecha =#" & Format(Fecenv, "yyyy/mm/dd") & "# And base =" & data_parsec.Recordset("base")
       data_cli.Refresh
    End If
    If data_cli.Recordset.RecordCount > 0 Then
       data_cli.Recordset.MoveFirst
       Do While Not data_cli.Recordset.EOF
            data_env.Recordset.AddNew
            data_env.Recordset("tipo_mov") = data_cli.Recordset("tipo_mov")
            data_env.Recordset("factura") = data_cli.Recordset("factura")
            data_env.Recordset("tipo") = data_cli.Recordset("tipo")
            data_env.Recordset("realizada") = data_cli.Recordset("realizada")
            data_env.Recordset("fecha") = data_cli.Recordset("fecha")
            data_env.Recordset("cod_cli") = data_cli.Recordset("cod_cli")
            data_env.Recordset("nom_cli") = data_cli.Recordset("nom_cli")
            data_env.Recordset("cod_prod") = data_cli.Recordset("cod_prod")
            data_env.Recordset("nom_prod") = data_cli.Recordset("nom_prod")
            data_env.Recordset("cantidad") = data_cli.Recordset("cantidad")
            data_env.Recordset("moneda") = data_cli.Recordset("moneda")
            data_env.Recordset("operador") = data_cli.Recordset("operador")
            data_env.Recordset("hora") = data_cli.Recordset("hora")
            data_env.Recordset("nro_flia") = data_cli.Recordset("nro_flia")
            data_env.Recordset("nom_flia") = data_cli.Recordset("nom_flia")
            data_env.Recordset("linea") = data_cli.Recordset("linea")
            data_env.Recordset("convenio") = data_cli.Recordset("convenio")
            data_env.Recordset("rub_cont") = data_cli.Recordset("rub_cont")
            data_env.Recordset("usa_timbre") = data_cli.Recordset("usa_timbre")
            data_env.Recordset("imp_timbre") = data_cli.Recordset("imp_timbre")
            data_env.Recordset("tot_lin") = data_cli.Recordset("tot_lin")
            data_env.Recordset("rub_nomb") = data_cli.Recordset("rub_nomb")
            data_env.Recordset("nro_med_a") = data_cli.Recordset("nro_med_a")
            data_env.Recordset("nom_med_a") = data_cli.Recordset("nom_med_a")
            data_env.Recordset("precio_est") = data_cli.Recordset("precio_est")
            data_env.Recordset("mes_paga") = data_cli.Recordset("mes_paga")
            data_env.Recordset("ano_paga") = data_cli.Recordset("ano_paga")
            data_env.Recordset("base") = data_cli.Recordset("base")
            data_env.Recordset("imp_iva") = data_cli.Recordset("imp_iva")
            data_env.Recordset("ruc") = data_cli.Recordset("ruc")
            data_env.Recordset.Update
            data_cli.Recordset.MoveNext
       Loop
    End If
       
' Arqueo
' Deudas
    Timer2.Enabled = False
    Timer6.Enabled = True
Else
   correo.Show
   correo.Timer2.Enabled = True
    
    Timer2.Enabled = False
'    Timer1.Enabled = True

End If

'borrar envios
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
'''frm_inicia.Hide
End
'frmMain.Timer10.Enabled = True
'''frmMain.Show
'''frmMain.Timer10 = True

End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
Timer1.Enabled = True


End Sub

Private Sub Timer5_Timer()
'MsgBox "No hay conexión a internet para transmitir, VERIFIQUE", vbCritical, "Mensaje"
Timer5.Enabled = False
Timer1.Enabled = True

End Sub

Private Sub Timer6_Timer()
   data_env.DatabaseName = ""
   data_env.RecordSource = ""
   data_env.Refresh
   data_cli.DatabaseName = ""
   data_cli.RecordSource = ""
   data_cli.Refresh
   If Dir("C:\Datos\envios.zip") <> "" Then
      Kill "c:\datos\envios.zip"
   End If
   Shell (App.Path & "\pkzip c:\datos\envios.zip c:\datos\env_*.*"), vbNormalFocus
   Timer6.Enabled = False
   Timer3.Enabled = True

End Sub
