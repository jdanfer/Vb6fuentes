VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_verdesconv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ver Descripción de convenio"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   3240
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
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
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7200
      Picture         =   "frm_verdesconv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cerrar"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Detalle del convenio:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   7455
   End
End
Attribute VB_Name = "frm_verdesconv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "dsn=" & Xconexrmt

If frm_largador.txt_cat.Text <> "" Then
    Adodc1.RecordSource = "Select * from convenio where cnv_codigo ='" & frm_largador.txt_cat.Text & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
       Adodc1.Recordset.MoveFirst
       If IsNull(Adodc1.Recordset("cnv_desc")) = False Then
          Label2.Caption = Adodc1.Recordset("cnv_desc")
       Else
          Label2.Caption = "Sin Datos"
       End If
       If IsNull(Adodc1.Recordset("cnv_motbaj")) = False Then
          Label1.Caption = Adodc1.Recordset("cnv_motbaj")
       Else
          Label1.Caption = "Sin Datos de convenio"
       End If
    Else
       MsgBox "No se encuentra el convenio, verifique!", vbInformation
    End If
End If


End Sub
