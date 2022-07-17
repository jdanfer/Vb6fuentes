VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_solhc 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Historia Clínica"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_solhc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9240
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   3240
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8400
      MaskColor       =   &H00FF8080&
      Picture         =   "frm_solhc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7080
      Picture         =   "frm_solhc.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Aceptar"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de la solicitud"
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.TextBox t_obs 
         Height          =   375
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3360
         Width           =   6135
      End
      Begin VB.TextBox t_pres 
         Height          =   375
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   12
         Top             =   2760
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_solhc.frx":0F56
         Left            =   2520
         List            =   "frm_solhc.frx":0F75
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox t_tel 
         Height          =   360
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox ctras 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TRASLADOS"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox cesp 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ESPECIALISTA"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox cpol 
         BackColor       =   &H00C0FFC0&
         Caption         =   "POLICLINICA"
         Height          =   255
         Left            =   6840
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox cmov 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MOVIL"
         Height          =   255
         Left            =   4800
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox ctot 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TODAS"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Observaciones:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "A presentar en:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Solicitante:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Telef. de contacto:"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   9480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Tipo de asistencias:"
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "frm_solhc.frx":0FC4
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2775
   End
End
Attribute VB_Name = "frm_solhc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_pres.SetFocus
End If

End Sub

Private Sub Command1_Click()
On Error GoTo Alsolhc

data1.ConnectionString = "dsn=" & Xconexrmt
If ctot.value = 1 Or cmov.value = 1 Or cpol.value = 1 Or cesp.value = 1 Or ctras.value = 1 Then
   If t_tel.Text <> "" Then
      If t_pres.Text <> "" Then
         data1.RecordSource = "Select * from provdeu where cliente =" & 301395
         data1.Refresh
         data1.Recordset.AddNew
         data1.Recordset("cod_cnv") = frmabm.txt_codcnv.Text
         data1.Recordset("nom_cnv") = Mid(frmabm.txt_nomcnv.Text, 1, 25)
         If frmabm.txt_matmut.Text = "" Then
         Else
            data1.Recordset("cliente") = Val(frmabm.txt_matmut.Text)
         End If
         data1.Recordset("origen") = t_tel.Text
         If ctot.value = 1 Then
            data1.Recordset("mes") = 1
         Else
            If cpol.value = 1 Then
               data1.Recordset("mes") = 2
            Else
               If cesp.value = 1 Then
                  data1.Recordset("mes") = 3
               Else
                  If ctras.value = 1 Then
                     data1.Recordset("mes") = 4
                  Else
                     If cmov.value = 1 Then
                        data1.Recordset("mes") = 5
                     Else
                        data1.Recordset("mes") = 1
                     End If
                  End If
               End If
            End If
         End If
         data1.Recordset("fecha") = Date
         data1.Recordset("documento") = frm_factura.labfac.Caption
         If Combo1.ListIndex >= 0 Then
            data1.Recordset("moneda") = Combo1.ListIndex
         End If
         data1.Recordset("nom_cobr") = t_pres.Text
         If t_obs.Text <> "" Then
            data1.Recordset("nombre") = t_obs.Text
         End If
         data1.Recordset("nro_cobr") = frmabm.txt_mat.Caption
         data1.Recordset.Update
         Xestaok = 1
         Unload Me
      Else
         MsgBox "No ingresó a quién se debe presentar", vbCritical, "Mensaje"
         Xestaok = 0
      End If
   Else
      MsgBox "Debe ingresar teléfono de contacto", vbInformation, "Mensaje"
      Xestaok = 0
   End If
Else
   MsgBox "Debe marcar una opción de Historia a buscar", vbCritical, "Mensaje"
   Xestaok = 0
End If

Exit Sub

Alsolhc:
        If Err.Number = 3155 Then
           MsgBox "Error al grabar, verifique datos o presione el botón de CANCELAR para terminar la factura.", vbInformation
        Else
           MsgBox "Error al grabar, verifique datos o presione el botón de CANCELAR para terminar la factura.", vbInformation
        End If

End Sub

Private Sub Command2_Click()
Xestaok = 1
Unload Me

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub t_pres_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_obs.SetFocus
End If

End Sub

Private Sub t_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
