VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_paspen 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasaje de recibos COBRADOS"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frm_paspen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc data_cob 
      Height          =   375
      Left            =   3000
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Caption         =   "data_cob"
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
   Begin MSAdodcLib.Adodc data_arqueo 
      Height          =   375
      Left            =   5280
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Caption         =   "data_arqueo"
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
   Begin MSAdodcLib.Adodc data_arq2 
      Height          =   375
      Left            =   5280
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Caption         =   "data_arq2"
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
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   840
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
      Caption         =   "data_cli"
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
   Begin VB.CommandButton btn_salir 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Picture         =   "frm_paspen.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   840
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Pasaje de recibos COBRADOS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6975
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   3840
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Caption         =   "data_lin"
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
      Begin VB.CommandButton btn_fin 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Terminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "frm_paspen.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txt_rec 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
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
         Left            =   1800
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txt_mat 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
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
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label labnom 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label labcantp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Left            =   4800
         TabIndex        =   12
         Top             =   2160
         Width           =   1680
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "Total de importe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label labcantr 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Left            =   4920
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cantidad de Recibos"
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
         Left            =   4560
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   4320
         X2              =   4320
         Y1              =   120
         Y2              =   3000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Nro. Recibo:"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Matrícula:"
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
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton btn_acep 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "frm_paspen.frx":0E0E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txt_nrocobr 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   1
      EndProperty
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
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label labnomcob 
      BackColor       =   &H00C0E0FF&
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
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "COBRADOR:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   2640
      Picture         =   "frm_paspen.frx":1398
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frm_paspen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_acep_Click()
If txt_nrocobr.Text <> "" Then
   data_cob.RecordSource = "Select * from cobrador where cb_numero =" & txt_nrocobr.Text
   data_cob.Refresh
   If data_cob.Recordset.RecordCount > 0 Then
      txt_nrocobr.Text = data_cob.Recordset("cb_numero")
      labnomcob.Caption = data_cob.Recordset("cb_nombre")
      btn_acep.Enabled = False
      txt_nrocobr.Enabled = False
      Frame1.Enabled = True
      txt_mat.SetFocus
      data_arqueo.RecordSource = "Select * from arqueo where cob =" & txt_nrocobr.Text & " and codpro =" & 98
      data_arqueo.Refresh
      If data_arqueo.Recordset.RecordCount > 0 Then
         MsgBox "EL ARQUEO PARA ESTE COBRADOR ESTA CERRADO.", vbInformation
         Unload Me
      End If
   Else
      MsgBox "No existe cobrador, Verifique", vbInformation, "Arqueos"
      btn_acep.Enabled = True
      txt_nrocobr.Enabled = True
      Frame1.Enabled = False
      txt_nrocobr.SetFocus
   End If
   data_cob.Recordset.Close
Else
   MsgBox "No ingresó cobrador, verifique", vbCritical, "Arqueos"
   btn_acep.Enabled = True
   txt_nrocobr.Enabled = True
   Frame1.Enabled = False
   txt_nrocobr.SetFocus
End If

End Sub

Private Sub btn_fin_Click()
Frame1.Enabled = False
btn_acep.Enabled = True
txt_nrocobr.Enabled = True
txt_nrocobr.Text = ""
txt_nrocobr.SetFocus
labnomcob.Caption = ""
End Sub

Private Sub btn_salir_Click()
'frm_paspen.Hide
Unload Me

End Sub


Private Sub Form_Load()
data_arqueo.ConnectionString = "dsn=" & Xconexrmt
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_cob.ConnectionString = "dsn=" & Xconexrmt
data_arq2.ConnectionString = "dsn=" & Xconexrmt
data_lin.ConnectionString = "dsn=" & Xconexrmt

Dim XnuevaF As Date
Dim XmesR, XanioR As Integer


frm_paspen.MousePointer = 11

data_arqueo.RecordSource = "Select * from deudas where origen >'" & "Refinancia" & "' and fecha_pago is null and mes_r is null"
data_arqueo.Refresh
If data_arqueo.Recordset.RecordCount > 0 Then
   data_arqueo.Recordset.MoveFirst
   Do While Not data_arqueo.Recordset.EOF
      XnuevaF = data_arqueo.Recordset("fecha") + data_arqueo.Recordset("nro_superv")
      XmesR = Month(XnuevaF)
      XanioR = Year(XnuevaF)
      data_arqueo.Recordset("mes_r") = XmesR
      data_arqueo.Recordset("anio_r") = XanioR
      data_arqueo.Recordset.Update
      data_arqueo.Recordset.MoveNext
   Loop
End If
frm_paspen.MousePointer = 0

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

    
End Sub

Private Sub txt_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_rec.SetFocus
End If
   
End Sub

Private Sub txt_mat_LostFocus()
If txt_mat.Text <> "" Then
'   data_cli.Recordset.FindFirst "cl_codigo =" & txt_mat.Text
   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & txt_mat.Text
   data_cli.Refresh
'   If Not data_cli.Recordset.NoMatch Then
   If data_cli.Recordset.RecordCount > 0 Then
      labnom.Caption = data_cli.Recordset("cl_apellid")
   Else
      labnom.Caption = ""
   End If
   data_cli.Recordset.Close
   
End If

End Sub

Private Sub txt_nrocobr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_nrocobr.Text <> "" Then
      data_cob.RecordSource = "Select * from cobrador where cb_numero =" & txt_nrocobr.Text
      data_cob.Refresh
      If data_cob.Recordset.RecordCount > 0 Then
         txt_nrocobr.Text = data_cob.Recordset("cb_numero")
         labnomcob.Caption = data_cob.Recordset("cb_nombre")
      Else
         MsgBox "No existe cobrador, Verifique", vbInformation, "Arqueos"
         txt_nrocobr.SetFocus
      End If
      data_cob.Recordset.Close
      
   End If
   btn_acep.SetFocus
End If

End Sub

Private Sub txt_rec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mat.SetFocus
End If

End Sub

Private Sub txt_rec_LostFocus()
On Error GoTo Yaestapas
Dim TextFec1, TexFec2 As String
Dim Fec1, Fec2, Fecrefin As Date
Dim Xmesref, Xanoref As Integer
Xmesref = 0
Xanoref = 0
If txt_rec.Text <> "" Then
'   data_arqueo.Recordset.FindFirst "Matricula =" & txt_mat.Text & " And Nrorec =" & txt_rec.Text
'   data_arqueo.RecordSource = "Select * from arqueo where matricula =" & txt_mat.Text & " and nrorec =" & txt_rec.Text
   data_arq2.RecordSource = "Select * from arqueo where matricula =" & txt_mat.Text & " and nrorec =" & txt_rec.Text
   data_arq2.Refresh
   If data_arq2.Recordset.RecordCount > 0 Then
      data_arqueo.RecordSource = "Select * from deudas where cliente =" & txt_mat.Text & " and documento =" & txt_rec.Text
      data_arqueo.Refresh
      If data_arqueo.Recordset.RecordCount > 0 Then
         Xmesref = data_arqueo.Recordset("mes")
         Xanoref = data_arqueo.Recordset("ano")
         If IsNull(data_arqueo.Recordset("fecha_pago")) = False Then
            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Text & " and mes_paga =" & data_arq2.Recordset("mes") & " and ano_paga =" & data_arq2.Recordset("ano") & " and cod_prod =" & 999
            data_lin.Refresh
            If data_lin.Recordset.RecordCount > 0 Then
               MsgBox "Ya se encuentra registrado el pago en BASE, el recibo pasará cómo DEVOLUCION"
               If data_arq2.Recordset.RecordCount > 0 Then
                  If data_arq2.Recordset("arqueo") = "D" Then
                  Else
'                     data_arq2.Recordset.Edit
                     data_arq2.Recordset("arqueo") = "D"
                     data_arq2.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                     data_arq2.Recordset("usuar") = WElusuario
                     data_arq2.Recordset.Update
                  End If
               End If
            Else
               If data_arq2.Recordset.RecordCount > 0 Then
                  If data_arq2.Recordset("arqueo") = "C" Then
                  Else
'                     data_arq2.Recordset.Edit
                     data_arq2.Recordset("arqueo") = "C"
                     data_arq2.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                     data_arq2.Recordset("usuar") = WElusuario
                     data_arq2.Recordset.Update
                  End If
               End If
            End If
            data_lin.Recordset.Close
         Else
            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Text & " and mes_paga =" & data_arq2.Recordset("mes") & " and ano_paga =" & data_arq2.Recordset("ano") & " and cod_prod =" & 999
            data_lin.Refresh
            If data_lin.Recordset.RecordCount > 0 Then
               MsgBox "Ya se encuentra registrado el pago en BASE, el recibo pasará cómo DEVOLUCION"
               data_arqueo.Recordset("fecha_pago") = Date
               data_arqueo.Recordset.Update
               If data_arq2.Recordset.RecordCount > 0 Then
                  If data_arq2.Recordset("arqueo") = "D" Then
                  Else
'                     data_arq2.Recordset.Edit
                     data_arq2.Recordset("arqueo") = "D"
                     data_arq2.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                     data_arq2.Recordset("usuar") = WElusuario
                     data_arq2.Recordset.Update
                  End If
               End If
            Else
               If data_arq2.Recordset("arqueo") = "C" Then
               Else
'                  data_arq2.Recordset.Edit
                  data_arq2.Recordset("arqueo") = "C"
                  data_arq2.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                  data_arq2.Recordset("usuar") = WElusuario
                  data_arq2.Recordset.Update
         
'                  data_arqueo.Recordset.Edit
                  data_arqueo.Recordset("fecha_pago") = Date
                  data_arqueo.Recordset.Update
                  If data_arqueo.Recordset("mes") > 9 Then
                     TextFec1 = "01/" & Trim(str(data_arqueo.Recordset("mes"))) & "/" & Trim(str(data_arqueo.Recordset("ano")))
                  Else
                     TextFec1 = "01/" & "0" & Trim(str(data_arqueo.Recordset("mes"))) & "/" & Trim(str(data_arqueo.Recordset("ano")))
                  End If
                  Fec1 = CDate(TextFec1)
                  data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_arqueo.Recordset("cliente")
                  data_cli.Refresh
                  If data_cli.Recordset.RecordCount > 0 Then
                     If IsNull(data_cli.Recordset("saldo_cc")) = False Then
'                        data_cli.Recordset.Edit
                        data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") - data_arqueo.Recordset("total")
                        If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
                           data_cli.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa") - 1
                        End If
                        If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
                           If data_cli.Recordset("cl_ultmesp") > 9 Then
                              If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
                                 TexFec2 = "01/" & Trim(str(data_cli.Recordset("cl_ultmesp"))) & "/" & Trim(str(data_cli.Recordset("cl_ultanop")))
                              Else
                                 TexFec2 = "01/" & Trim(str(data_cli.Recordset("cl_ultmesp"))) & "/" & Trim(str(data_arqueo.Recordset("ano")))
                              End If
                           Else
                              If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
                                 TexFec2 = "01/" & "0" & Trim(str(data_cli.Recordset("cl_ultmesp"))) & "/" & Trim(str(data_cli.Recordset("cl_ultanop")))
                              Else
                                 TexFec2 = "01/" & "0" & Trim(str(data_cli.Recordset("cl_ultmesp"))) & "/" & Trim(str(data_arqueo.Recordset("ano")))
                              End If
                           End If
                           Fec2 = CDate(TextFec2)
                        Else
                           Fec2 = CDate(TextFec1)
                           Fec1 = CDate(TextFec1)
                        End If
                        If Format(Fec2, "yyyy/mm/dd") >= Format(Fec1, "yyyy/mm/dd") Then
                        Else
                           data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes")
                           data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano")
                        End If
                        data_cli.Recordset.Update
                     Else
'                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_ultmesp") = data_arqueo.Recordset("mes")
                        data_cli.Recordset("cl_ultanop") = data_arqueo.Recordset("ano")
                        data_cli.Recordset.Update
                     End If
                  End If
                  data_cli.Recordset.Close
                  If labcantr.Caption <> "" Then
                     labcantr.Caption = labcantr.Caption + 1
                  Else
                     labcantr.Caption = 1
                  End If
                  If labcantp.Caption <> "" Then
                     labcantp.Caption = labcantp.Caption + data_arq2.Recordset("total")
                  Else
                     labcantp.Caption = data_arq2.Recordset("total")
                  End If
               End If
            End If
            data_lin.Recordset.Close
         End If
         data_arqueo.RecordSource = "Select * from deudas where origen >='" & "Refinancia" & "' and fecha_pago is null and cliente =" & txt_mat.Text & " and mes_r =" & data_arq2.Recordset("mes") & " and anio_r =" & data_arq2.Recordset("ano")
         data_arqueo.Refresh
         If data_arqueo.Recordset.RecordCount > 0 Then
            data_arqueo.Recordset("fecha_pago") = Date
            data_arqueo.Recordset.Update
         End If
         data_arqueo.Recordset.Close
      Else
         If data_arq2.Recordset("arqueo") = "C" Then
         Else
'            data_arq2.Recordset.Edit
            data_arq2.Recordset("arqueo") = "C"
            data_arq2.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
            data_arq2.Recordset("usuar") = WElusuario
            data_arq2.Recordset.Update
         End If
         data_arqueo.Recordset.Close
      End If
      
   Else
      MsgBox "No se encuentra documento"
   End If
   data_arq2.Recordset.Close
   
End If
txt_mat.Text = ""
txt_rec.Text = ""
data_arqueo.Refresh

Exit Sub

Yaestapas:
          If Err.Number = 3197 Then
             MsgBox "Verifique si el recibo ya fue pasado.", vbInformation, "Pendientes"
          Else
             MsgBox "Hay un error en los datos, VERIFIQUE!!", vbInformation, "Pendientes"
          End If

End Sub
