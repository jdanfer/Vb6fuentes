VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_deudasarq 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmar deudas del socio"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   Icon            =   "frm_deudasarq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3836
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   5280
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label labnom 
      BackColor       =   &H00FF0000&
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
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label labcli 
      BackColor       =   &H00FF0000&
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
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3840
      Picture         =   "frm_deudasarq.frx":058A
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "frm_deudasarq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If data1.Recordset.RecordCount > 0 Then
   data1.Recordset.MoveFirst
   Do While Not data1.Recordset.EOF
      If data1.Recordset("arqueo") <> "B" Then
'         data1.Recordset.Edit
         data1.Recordset("arqueo") = "B"
         data1.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
         data1.Recordset("usuar") = WElusuario
         data1.Recordset.Update
      End If
      data1.Recordset.MoveNext
   Loop
   MsgBox "Los recibos fueron pasados cómo BAJA", vbInformation
End If
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
labcli.Caption = frmabm.txt_mat.Caption
labnom.Caption = frmabm.txt_apellid.Text
data1.ConnectionString = "dsn=" & Xconexrmt

If labcli.Caption <> "" Then
   data1.RecordSource = "Select * from arqueo where matricula =" & Val(labcli.Caption) & " and arqueo <>'" & "C" & "'"
   data1.Refresh
End If
DBGrid1.Rows = 2
DBGrid1.Cols = 8
DBGrid1.TextMatrix(0, 0) = "MES"
DBGrid1.ColWidth(0) = 500
DBGrid1.TextMatrix(0, 1) = "AÑO"
DBGrid1.ColWidth(1) = 500
DBGrid1.TextMatrix(0, 2) = "COLOR"
DBGrid1.ColWidth(2) = 500
DBGrid1.TextMatrix(0, 3) = "CATEG."
DBGrid1.ColWidth(3) = 900
DBGrid1.TextMatrix(0, 4) = "ARQ."
DBGrid1.ColWidth(4) = 500
DBGrid1.TextMatrix(0, 5) = "NRO.REC"
DBGrid1.ColWidth(5) = 1200
DBGrid1.TextMatrix(0, 6) = "COBR."
DBGrid1.ColWidth(6) = 900
DBGrid1.TextMatrix(0, 7) = "TOTAL $."
DBGrid1.ColWidth(7) = 1500

Dim Xcann As Integer
Xcann = 1
If data1.Recordset.RecordCount > 0 Then
    data1.Recordset.MoveFirst
    Do While Not data1.Recordset.EOF
       If IsNull(data1.Recordset("mes")) = False Then
          DBGrid1.TextMatrix(Xcann, 0) = data1.Recordset("mes")
       End If
       If IsNull(data1.Recordset("ano")) = False Then
          DBGrid1.TextMatrix(Xcann, 1) = data1.Recordset("ano")
       End If
       If IsNull(data1.Recordset("color")) = False Then
          DBGrid1.TextMatrix(Xcann, 2) = data1.Recordset("color")
       End If
       If IsNull(data1.Recordset("cat")) = False Then
          DBGrid1.TextMatrix(Xcann, 3) = data1.Recordset("cat")
       End If
       If IsNull(data1.Recordset("arqueo")) = False Then
          DBGrid1.TextMatrix(Xcann, 4) = data1.Recordset("arqueo")
       End If
       If IsNull(data1.Recordset("nrorec")) = False Then
          DBGrid1.TextMatrix(Xcann, 5) = data1.Recordset("nrorec")
       End If
       If IsNull(data1.Recordset("cob")) = False Then
          DBGrid1.TextMatrix(Xcann, 6) = data1.Recordset("cob")
       End If
       If IsNull(data1.Recordset("total")) = False Then
          DBGrid1.TextMatrix(Xcann, 7) = data1.Recordset("total")
       End If
       
       DBGrid1.Rows = DBGrid1.Rows + 1
       data1.Recordset.MoveNext
       Xcann = Xcann + 1
    Loop
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
