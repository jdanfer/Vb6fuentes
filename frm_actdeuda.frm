VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_actdeuda 
   BackColor       =   &H00FF8080&
   Caption         =   "Actualizar deuda por socio"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_actdeuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Recuperar deuda"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Agregar Factura a la deuda"
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Es deuda por cuota mensual"
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Limpiar ficha"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox t_imp 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3000
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mf 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox t_nrofact 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Data data_lindeu 
      Caption         =   "data_lindeu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Caption         =   "Importe (Opcional):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Fecha de Pago:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Nro. de factura:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6720
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "NUMERO DE MATRICULA:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   240
      Picture         =   "frm_actdeuda.frx":0442
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "frm_actdeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
   If Check3.Value = 1 Then
      Check3.Value = 0
   End If
   If Check2.Value = 1 Then
      Check2.Value = 0
   End If
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   If Check3.Value = 1 Then
      Check3.Value = 0
   End If
   If Check1.Value = 1 Then
      Check1.Value = 0
   End If
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   If Check1.Value = 1 Then
      Check1.Value = 0
   End If
   If Check2.Value = 1 Then
      Check2.Value = 0
   End If
End If

End Sub

Private Sub Command1_Click()
Dim Xmat, Xmes, Xano, Xbancuantos As Long
Dim Xsaldo As Double
On Error GoTo Nosepuede

frm_actdeuda.MousePointer = 11
Command1.Enabled = False
If Check1.Value = 1 Or Check4.Value = 1 Then
   If Check1.Value = 1 Then
      data_cli.Recordset.Edit
      data_cli.Recordset("saldo_cc") = 0
      data_cli.Recordset("cl_atrasoa") = 0
      data_cli.Recordset.Update
      data_cli.Refresh
   Else
      data1.RecordSource = "Select * from deudas where documento =" & t_nrofact.Text & " and cliente =" & Text1.Text & " and fecha_pago is not null"
      data1.Refresh
      If data1.Recordset.RecordCount > 0 Then
         data1.Recordset.MoveFirst
         data1.Recordset.Edit
         data1.Recordset("fecha_pago") = Null
         data1.Recordset.Update
      Else
         MsgBox "No hay registros"
      End If
   End If
Else
   If Check2.Value = 1 Then
      If Text1.Text = "" Then
         MsgBox "Debe ingresar matrícula"
      Else
         If t_nrofact.Text = "" Then
            MsgBox "Debe ingresar número de documento"
         Else
            data1.RecordSource = "Select * from deudas where documento =" & t_nrofact.Text & " and cliente =" & Text1.Text & " and fecha_pago is null"
            data1.Refresh
            If data1.Recordset.RecordCount > 0 Then
               data1.Recordset.MoveFirst
               data1.Recordset.Edit
               data1.Recordset("fecha_pago") = Format(mf.Text, "dd/mm/yyyy")
               data1.Recordset.Update
               If IsNull(data_cli.Recordset("saldo_cc")) = False Then
                  If data_cli.Recordset("saldo_cc") = 0 Then
                  Else
                     data_cli.Recordset.Edit
                     If data_cli.Recordset("saldo_cc") < 0 Then
                        data_cli.Recordset("saldo_cc") = 0
                        data_cli.Recordset("cl_atrasoa") = 0
                     Else
                        data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") - data1.Recordset("total")
                        data_cli.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa") - 1
                     End If
                     data_cli.Recordset.Update
                  End If
               End If
            End If
         End If
      End If
   Else
      If Check3.Value = 1 Then
         data1.RecordSource = "Select * from deudas where documento =" & t_nrofact.Text & " and cliente =" & Text1.Text & " and fecha_pago is not null"
         data1.Refresh
         If data1.Recordset.RecordCount > 0 Then
            data1.Recordset.MoveFirst
            data1.Recordset.Edit
            data1.Recordset("fecha_pago") = Null
            data1.Recordset.Update
            If Label6.Caption <> "" Then
               If IsNull(data_cli.Recordset("saldo_cc")) = False Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") + data1.Recordset("total")
                  data_cli.Recordset.Update
               End If
            End If
         Else
            MsgBox "No figura la factura en archivo deudas, VERIFIQUE!!", vbInformation
         End If
      Else
         If t_imp.Text <> "" Then
         Else
            t_imp.Text = 0
         End If
         data1.RecordSource = "Select * from deudas where documento =" & t_nrofact.Text & " and cliente =" & Text1.Text & " and fecha_pago is null"
         data1.Refresh
         If data1.Recordset.RecordCount > 0 Then
            data1.Recordset.MoveFirst
            data1.Recordset.Edit
            data1.Recordset("fecha_pago") = Format(mf.Text, "dd/mm/yyyy")
            data1.Recordset.Update
            If Label6.Caption <> "" Then
               If IsNull(data_cli.Recordset("saldo_cc")) = False Then
                  data_cli.Recordset.Edit
                  data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") - Val(Label6.Caption)
                  data_cli.Recordset.Update
               End If
            End If
         Else
            If IsNull(data_cli.Recordset("saldo_cc")) = False Then
               data_cli.Recordset.Edit
               data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") - Val(t_imp.Text)
               data_cli.Recordset.Update
            End If
         End If
      End If
   End If
End If
Command1.Enabled = True
frm_actdeuda.MousePointer = 0
MsgBox "Proceso terminado"

Exit Sub

Nosepuede:
         If Err.Number = 3150 Then
            MsgBox "Error al grabar", vbCritical
         Else
            MsgBox "Error al actualizar, verifique datos.", vbCritical
         End If
         
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data1.DatabaseName = ""
data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data1.RecordSource = "deudas"
'Data1.Refresh
'data_lindeu.DatabaseName = App.Path & "\sapp.mdb"
data_lindeu.Connect = "odbc;dsn=" & Xconexrmt & ";"

If WElusuario = "JFERNAN" Then
   Check1.Visible = True
Else
   Check1.Visible = False
End If

End Sub

Private Sub Form_Resize()
With Image1
     .Top = 0
     .Left = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub

Private Sub t_nrofact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mf.SetFocus
End If

End Sub

Private Sub t_nrofact_LostFocus()
If t_nrofact.Text <> "" Then
   If Text1.Text <> "" Then
      If Check3.Value = 1 Then
         data1.RecordSource = "Select * from deudas where documento =" & t_nrofact.Text & " and cliente =" & Text1.Text & " and fecha_pago is not null"
      Else
         data1.RecordSource = "Select * from deudas where documento =" & t_nrofact.Text & " and cliente =" & Text1.Text & " and fecha_pago is null"
      End If
      data1.Refresh
      If data1.Recordset.RecordCount > 0 Then
         Label6.Caption = Format(data1.Recordset("total"), "Standard")
      Else
         Label6.Caption = 0
         t_imp.SetFocus
      End If
   Else
      Label6.Caption = 0
      t_imp.SetFocus
   End If
Else
   Label6.Caption = 0
   t_imp.SetFocus
End If
         
   
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nrofact.SetFocus
End If

End Sub

Private Sub Text1_LostFocus()
If Text1.Text <> "" Then
'   data_cli.DatabaseName = App.Path & "\sapp.mdb"
   data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Text1.Text
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      Label2.Caption = data_cli.Recordset("cl_apellid")
   Else
      Label2.Caption = ""
   End If
End If

End Sub
