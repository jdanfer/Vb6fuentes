VERSION 5.00
Begin VB.Form frm_uno 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reinicio de PC"
   ClientHeight    =   2610
   ClientLeft      =   1560
   ClientTop       =   2775
   ClientWidth     =   5595
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "formuno.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2610
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\control.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tcontrol"
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3600
      Top             =   360
   End
   Begin VB.PictureBox picGancho 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
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
         Caption         =   "Ejecutar Horas"
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
Attribute VB_Name = "frm_uno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    t.szTip = "Horas iniciado" & Chr$(0) ' Es un string de "C" ( \0 )
    Shell_NotifyIcon NIM_ADD, t
    Me.Hide
    App.TaskVisible = False
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
Shell ("C:\Horas\DiscoC\Horas.exe"), vbNormalFocus

End Sub
Private Sub mnuSalir_Click(Index As Integer)
'    Unload Me
MsgBox "No puede cerrar el programa", vbInformation

End Sub
Private Sub picGancho_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, Msg As Long, ValDev As Long
    Msg = X / Screen.TwipsPerPixelX

    If rec = False Then
        rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                ValDev = WinExec("c:\Horas\DiscoC\Horas.exe", 1)
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


Private Sub Timer1_Timer()
Dim aa As Long
Dim bb As Long
Dim Dif As Long
If Data1.Recordset.RecordCount = 0 Then
   Data1.Recordset.AddNew
   Data1.Recordset("fecha") = Date
   Data1.Recordset("hora") = Time
   Data1.Recordset("nrohora") = Timer
   Data1.Recordset("error") = 0
   Data1.Recordset("ejecuta") = Timer
   Data1.Recordset.Update
   Data1.Refresh
End If
If Data1.Recordset("error") <> 1 Then
    aa = Data1.Recordset("nrohora")
    bb = Timer
    Dif = bb - aa
    
    If Dif <= -180 Then
       If Dif <= -5000 Then
    '      MsgBox "Verifique", vbInformation
            Data1.Recordset.Edit
            Data1.Recordset("fecha") = Date
            Data1.Recordset("hora") = Time
            Data1.Recordset("nrohora") = Timer
            Data1.Recordset("error") = 0
            Data1.Recordset("ejecuta") = Timer
            Data1.Recordset.Update
       Else
    '      MsgBox "Error", vbCritical
          Data1.Recordset.Edit
          Data1.Recordset("fecha") = Date
          Data1.Recordset("hora") = Time
          Data1.Recordset("nrohora") = Timer
          Data1.Recordset("error") = 1
          Data1.Recordset("ejecuta") = Timer
          Data1.Recordset.Update
'          End
       End If
    Else
       Data1.Recordset.Edit
       Data1.Recordset("fecha") = Date
       Data1.Recordset("hora") = Time
       Data1.Recordset("nrohora") = Timer
       Data1.Recordset("error") = 0
       Data1.Recordset("ejecuta") = Timer
       Data1.Recordset.Update
    End If
End If

End Sub
