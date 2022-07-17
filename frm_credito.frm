VERSION 5.00
Begin VB.Form frm_credito 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Venta a crédito"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   6720
      Picture         =   "frm_credito.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   240
      Picture         =   "frm_credito.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingrese datos de responsable del crédito"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.TextBox t_cod 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4320
         MaxLength       =   1
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox t_ced 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   3
         Top             =   855
         Width           =   1695
      End
      Begin VB.TextBox t_trab 
         Height          =   375
         Left            =   2640
         MaxLength       =   25
         TabIndex        =   11
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox t_correo 
         Height          =   375
         Left            =   2640
         MaxLength       =   70
         TabIndex        =   10
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox t_telef 
         Height          =   360
         Left            =   2640
         MaxLength       =   35
         TabIndex        =   8
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox t_domi 
         Height          =   360
         Left            =   2640
         MaxLength       =   35
         TabIndex        =   6
         Top             =   1320
         Width           =   5295
      End
      Begin VB.TextBox t_apel 
         Height          =   360
         Left            =   2640
         MaxLength       =   35
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Cédula:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Lugar de trabajo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Otros teléfonos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Celular:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Domicilio:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Apellidos/Nombres:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   2760
      Picture         =   "frm_credito.frx":0B14
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frm_credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Alingcredito

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from cabezal where factura =" & 1515
Data1.Refresh
If Len(t_apel.Text) > 10 Then
   If Len(t_ced.Text) > 5 Then
      If Len(t_cod.Text) >= 1 Then
         If Len(t_domi.Text) > 10 Then
            If Len(t_telef.Text) > 4 Then
               Data1.Recordset.AddNew
               Data1.Recordset("factura") = frm_factura.labfac.Caption
               Data1.Recordset("fecha") = Date
               Data1.Recordset("nom_cli") = t_apel.Text
               Data1.Recordset("cod_cli") = Val(t_ced.Text)
               Data1.Recordset("dias") = Val(t_cod.Text)
               Data1.Recordset("dir_cli") = t_domi.Text
               Data1.Recordset("loc_cli") = t_telef.Text
               If t_correo.Text <> "" Then
                  Data1.Recordset("impalfa") = t_correo.Text
               End If
               If t_trab.Text <> "" Then
                  Data1.Recordset("ruc") = t_trab.Text
               End If
               Data1.Recordset("danulada") = 0
               Data1.Recordset("impreso") = 0
               Data1.Recordset.Update
               Xestaok = 1
               t_apel.Text = ""
               t_ced.Text = ""
               t_cod.Text = ""
               t_domi.Text = ""
               t_telef.Text = ""
               t_correo.Text = ""
               t_trab.Text = ""
               Unload Me
            Else
               MsgBox "No ingresó TELEFONO", vbCritical, "Datos créditos"
               t_telef.SetFocus
            End If
         Else
            MsgBox "No ingresó DOMICILIO", vbCritical, "Datos créditos"
            t_domi.SetFocus
         End If
      Else
         MsgBox "No ingresó dígito de CEDULA", vbCritical, "Datos créditos"
         t_cod.SetFocus
      End If
   Else
      MsgBox "No ingresó CEDULA", vbCritical, "Datos crédito"
      t_ced.SetFocus
   End If
Else
   MsgBox "No ingresó NOMBRES", vbCritical, "Datos crédito"
   t_apel.SetFocus
End If

Exit Sub

Alingcredito:
              If Err.Number = 3155 Then
                 MsgBox "Hay un error en el registro de los datos, verifique o presione el botón de CANCELAR para continuar con la factura.", vbInformation
              Else
                 MsgBox "Hay un error en el registro de los datos, verifique o presione el botón de CANCELAR para continuar con la factura.", vbInformation
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

Private Sub t_apel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_ced.SetFocus
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cod.SetFocus
End If

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_domi.SetFocus
End If

End Sub

Private Sub t_correo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_trab.SetFocus
End If

End Sub

Private Sub t_domi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_telef.SetFocus
End If

End Sub

Private Sub t_telef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_correo.SetFocus
End If

End Sub

Private Sub t_trab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
