VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_valida_datos_socio 
   BorderStyle     =   0  'None
   Caption         =   "Verificar datos cliente"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6465
   Icon            =   "frm_valida_datos_socio.frx":0000
   LinkTopic       =   "Verificar datos cliente"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextMail 
      Height          =   375
      Left            =   1680
      MaxLength       =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox TextCelular 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox TextTelefono 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton confirmar 
      Caption         =   "confirmar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Verificar datos del cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7575
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12015
      Begin VB.Data data_cli 
         Caption         =   "data_cli"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6000
         Picture         =   "frm_valida_datos_socio.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cerrar sin grabar"
         Top             =   120
         Width           =   495
      End
      Begin VB.Data data_parsec 
         Caption         =   "data_parsec"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data data_abm 
         Caption         =   "data_abm"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4320
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Data data_history 
         Caption         =   "data_history"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4560
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Data data_mutual 
         Caption         =   "data_mutual"
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
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSDBCtls.DBCombo cbomutual 
         Bindings        =   "frm_valida_datos_socio.frx":0B14
         Height          =   330
         Left            =   1680
         TabIndex        =   10
         Top             =   2760
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "ca_nom"
         BoundColumn     =   "ca_nom"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label labCodConv 
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label labXmat 
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LError1 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   480
         TabIndex        =   9
         Top             =   4440
         Width           =   5175
      End
      Begin VB.Label LMutualista 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Mutualista:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label LMail 
         BackColor       =   &H00C0FFC0&
         Caption         =   "E-Mail:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label LCelular 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Celular:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label LTelefono 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_valida_datos_socio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private num_socio As String

'WRITTEN BY ROBDOG888
Private Declare Function RemoveMenu Lib "user32" ( _
                    ByVal hMenu As Long, _
                    ByVal nPosition As Long, _
                    ByVal wFlags As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemCount Lib "user32.dll" ( _
                    ByVal hMenu As Long) As Long

Private Const MF_BYPOSITION = &H400&

Public Property Let SetSocio(ByVal NewValue As String)
    num_socio = NewValue
End Property

Private Sub Command1_Click()
MsgBox "No ha confirmado datos!!!", vbCritical
DatosVerificadosOk = 1
XorigenDatos = 0

Unload Me

End Sub

Private Sub confirmar_Click()
    Dim Verifica_losdatos, XX As Integer
    Dim textocorreo, XcodConv As String
    Dim XmatDatos As Long
    Dim ConfirmarCorreoSi As String
    ConfirmarCorreoSi = ""
    textocorreo = ""
    Verifica_losdatos = 0

    On Error GoTo errorws
    
    If TextMail.Text = "NO APLICA" Then
       ConfirmarCorreoSi = MsgBox("Confirma que el paciente NO TIENE CORREO ELECTRÓNICO?", vbCritical + vbYesNo, "Verificación Datos")
    Else
       ConfirmarCorreoSi = vbYes
    End If
    If XorigenDatos = 1 Or XorigenDatos = 4 Then
       labXmat.Caption = frmabm.txt_mat.Caption
       labcodconv.Caption = frmabm.txt_codcnv.Text
    Else
       If XorigenDatos = 2 Then
          labXmat.Caption = frm_largador.txt_mat.Text
          labcodconv.Caption = frm_largador.txt_cat.Text
       Else
          If XorigenDatos = 3 Then
             labXmat.Caption = frm_especialistas.t_mat.Text
             labcodconv.Caption = frm_especialistas.t_conv.Text
          Else
             labXmat.Caption = 0
             labcodconv.Caption = "AA"
          End If
       End If
    End If
       
    If Trim(TextTelefono.Text) = "no aplica" Then
       TextTelefono.Text = "NO APLICA"
    End If
    If Trim(TextCelular.Text) = "no aplica" Then
       TextCelular.Text = "NO APLICA"
    End If
    If Trim(TextMail.Text) = "no aplica" Then
       TextMail.Text = "NO APLICA"
    End If
    If Trim(cbomutual.Text) = "no aplica" Then
       cbomutual.Text = "NO APLICA"
    End If
    If Trim(TextTelefono.Text) <> "" Then
       If IsNumeric(TextTelefono.Text) = True Then
          If Len(TextTelefono.Text) < 7 Then
             Verifica_losdatos = 1
          End If
       Else
          If Trim(TextTelefono.Text) <> "NO APLICA" Then
             Verifica_losdatos = 1
          End If
       End If
    Else
       Verifica_losdatos = 1
    End If
    If Trim(TextCelular.Text) <> "" Then
       If IsNumeric(TextCelular.Text) = True Then
          If Len(TextCelular.Text) < 9 Then '096317534
             Verifica_losdatos = 1
          End If
       Else
          If TextCelular.Text <> "NO APLICA" Then
             Verifica_losdatos = 1
          End If
       End If
    Else
       Verifica_losdatos = 1
    End If
    If Trim(TextMail.Text) <> "" Then
       If TextMail.Text <> "NO APLICA" Then
          For XX = 1 To Len(TextMail.Text)
              If Mid(TextMail.Text, XX, 1) = "@" Then
                 textocorreo = "@"
              Else
                 If Mid(TextMail.Text, XX, 1) = "." Then
                    If textocorreo = "@" Then
                       textocorreo = textocorreo + "."
                    End If
                 End If
              End If
          Next
          If textocorreo = "@." Then
          Else
             Verifica_losdatos = 1
          End If
       End If
    Else
       Verifica_losdatos = 1
    End If
    If Trim(cbomutual.Text) <> "" Then
    Else
       Verifica_losdatos = 1
    End If
    If Verifica_losdatos = 1 Then
       MsgBox "Hay error en los datos, verifique!", vbCritical
    Else
       If ConfirmarCorreoSi = vbYes Then
            altaValidacionDatos
            altaValidacionDatosabm
            If XorigenDatos = 1 Or XorigenDatos = 4 Then
               If XorigenDatos = 4 Then
                  frmabm.txt_telef.Text = TextTelefono.Text
                  frmabm.t_cel.Text = TextCelular.Text
                  frmabm.t_correo.Text = TextMail.Text
                  frmabm.cbomutual.Text = cbomutual.Text
                  data_cli.RecordSource = "select * from clientes where cl_codigo =" & Val(frmabm.txt_mat.Caption)
                  data_cli.Refresh
                  If data_cli.Recordset.RecordCount > 0 Then
                     If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                        If Trim(data_cli.Recordset("cl_telefon")) <> Trim(TextTelefono.Text) Then
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_telefon") = Trim(TextTelefono.Text)
                           data_cli.Recordset.Update
                        End If
                     Else
                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_telefon") = Trim(TextTelefono.Text)
                        data_cli.Recordset.Update
                     End If
                     If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                        If Trim(data_cli.Recordset("cl_dpto")) <> Trim(TextCelular.Text) Then
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_dpto") = Trim(TextCelular.Text)
                           data_cli.Recordset.Update
                        End If
                     Else
                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_dpto") = Trim(TextCelular.Text)
                        data_cli.Recordset.Update
                     End If
                     If IsNull(data_cli.Recordset("cl_referen")) = False Then
                        If Trim(data_cli.Recordset("cl_referen")) <> Trim(TextMail.Text) Then
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_referen") = Mid(Trim(TextMail.Text), 1, 74)
                           data_cli.Recordset.Update
                        End If
                     Else
                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_referen") = Mid(Trim(TextCelular.Text), 1, 74)
                        data_cli.Recordset.Update
                     End If
                     If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
                        If Trim(data_cli.Recordset("cl_socmnom")) <> Trim(cbomutual.Text) Then
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_socmnom") = Trim(cbomutual.Text)
                           data_cli.Recordset.Update
                        End If
                     Else
                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_socmnom") = Trim(cbomutual.Text)
                        data_cli.Recordset.Update
                     End If
                  End If
               Else
                  frmabm.txt_telef.Text = TextTelefono.Text
                  frmabm.t_cel.Text = TextCelular.Text
                  frmabm.t_correo.Text = TextMail.Text
                  frmabm.cbomutual.Text = cbomutual.Text
               End If
            Else
               If XorigenDatos = 2 Then
                  If data_cli.Recordset.RecordCount > 0 Then
                     If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                        If Trim(data_cli.Recordset("cl_telefon")) <> Trim(TextTelefono.Text) Then
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_telefon") = Trim(TextTelefono.Text)
                           data_cli.Recordset.Update
                        End If
                     Else
                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_telefon") = Trim(TextTelefono.Text)
                        data_cli.Recordset.Update
                     End If
                     If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                        If Trim(data_cli.Recordset("cl_dpto")) <> Trim(TextCelular.Text) Then
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_dpto") = Trim(TextCelular.Text)
                           data_cli.Recordset.Update
                        End If
                     Else
                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_dpto") = Trim(TextCelular.Text)
                        data_cli.Recordset.Update
                     End If
                     If IsNull(data_cli.Recordset("cl_referen")) = False Then
                        If Trim(data_cli.Recordset("cl_referen")) <> Trim(TextMail.Text) Then
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_referen") = Mid(Trim(TextMail.Text), 1, 74)
                           data_cli.Recordset.Update
                        End If
                     Else
                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_referen") = Mid(Trim(TextCelular.Text), 1, 74)
                        data_cli.Recordset.Update
                     End If
                     If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
                        If Trim(data_cli.Recordset("cl_socmnom")) <> Trim(cbomutual.Text) Then
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_socmnom") = Trim(cbomutual.Text)
                           data_cli.Recordset.Update
                        End If
                     Else
                        data_cli.Recordset.Edit
                        data_cli.Recordset("cl_socmnom") = Trim(cbomutual.Text)
                        data_cli.Recordset.Update
                     End If
                  End If
               Else
                  If XorigenDatos = 3 Then
                     If data_cli.Recordset.RecordCount > 0 Then
                        If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                           If Trim(data_cli.Recordset("cl_telefon")) <> Trim(TextTelefono.Text) Then
                              data_cli.Recordset.Edit
                              data_cli.Recordset("cl_telefon") = Trim(TextTelefono.Text)
                              data_cli.Recordset.Update
                           End If
                        Else
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_telefon") = Trim(TextTelefono.Text)
                           data_cli.Recordset.Update
                        End If
                        If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                           If Trim(data_cli.Recordset("cl_dpto")) <> Trim(TextCelular.Text) Then
                              data_cli.Recordset.Edit
                              data_cli.Recordset("cl_dpto") = Trim(TextCelular.Text)
                              data_cli.Recordset.Update
                           End If
                        Else
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_dpto") = Trim(TextCelular.Text)
                           data_cli.Recordset.Update
                        End If
                        If IsNull(data_cli.Recordset("cl_referen")) = False Then
                           If Trim(data_cli.Recordset("cl_referen")) <> Trim(TextMail.Text) Then
                              data_cli.Recordset.Edit
                              data_cli.Recordset("cl_referen") = Mid(Trim(TextMail.Text), 1, 74)
                              data_cli.Recordset.Update
                           End If
                        Else
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_referen") = Mid(Trim(TextCelular.Text), 1, 74)
                           data_cli.Recordset.Update
                        End If
                        If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
                           If Trim(data_cli.Recordset("cl_socmnom")) <> Trim(cbomutual.Text) Then
                              data_cli.Recordset.Edit
                              data_cli.Recordset("cl_socmnom") = Trim(cbomutual.Text)
                              data_cli.Recordset.Update
                           End If
                        Else
                           data_cli.Recordset.Edit
                           data_cli.Recordset("cl_socmnom") = Trim(cbomutual.Text)
                           data_cli.Recordset.Update
                        End If
                     End If
                  End If
               End If
            End If
            MsgBox "Datos validados con éxito!", vbInformation
            DatosVerificadosOk = 0
            XorigenDatos = 0
            Unload Me
       
       Else
          MsgBox "Verifique datos!", vbCritical
       End If
              
    End If
              
    
    Exit Sub

errorws:
        MsgBox "ERROR:" & Err.Description
        

End Sub



Private Sub Form_Load()
    '''''REMOVE THE SYSTEM MENU ITEM - CLOSE
    'RemoveMenu GetSystemMenu(Me.hwnd, 0), GetMenuItemCount(GetSystemMenu(Me.hwnd, 0)) - 1, MF_BYPOSITION
''''    'REMOVE THE MENU SEPARATOR
    'RemoveMenu GetSystemMenu(Me.hwnd, 0), GetMenuItemCount(GetSystemMenu(Me.hwnd, 0)) - 1, MF_BYPOSITION
    'Dim responseText As String
    'Dim telefono As String
    'Dim celular As String
    'Dim mail As String
    'Dim mutualista As String
    'Dim url_ws_rest_clientes
    'Dim Status As Integer
    On Error GoTo errorws
    data_mutual.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_mutual.RecordSource = "select * from ca_adm order by ca_nom"
    data_mutual.Refresh
    data_history.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_abm.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_parsec.DatabaseName = App.path & "\parse.mdb"
    data_parsec.RecordSource = "parsec0"
    data_parsec.Refresh

'    url_ws_rest_clientes = GetParametroBD.getValor(2)
'    Set objDatosCliente = consumirServicio2("GET", url_ws_rest_clientes & "/clientes/" & num_socio, "", 2)
'    responseText = objDatosCliente.responseText
'    Status = objDatosCliente.Status
'    Set responseJson = JSON.parse(responseText)
'    If Status <> 200 Then
'        LogError 100, responseText, "Form_Load - frm_valida_datos_socios, status<>200: " & Status, Erl
'        GoTo errorws
'    End If
    '
    'traigo los datos actuales del cliente
    '
'    If IsNull(Trim(responseJson.Item("clTelefon"))) Then
'        telefono = ""
'    Else
'        telefono = responseJson.Item("clTelefon")
'    End If
    'celular
    'TextMutualista.Text = mutualista
    If XorigenDatos = 1 Or XorigenDatos = 4 Then ' Ficha
       TextTelefono.Text = frmabm.txt_telef.Text
       TextCelular.Text = frmabm.t_cel.Text
       TextMail.Text = frmabm.t_correo.Text
       cbomutual.Text = frmabm.cbomutual.Text
    Else
       If XorigenDatos = 2 Then 'Despacho
          data_cli.RecordSource = "select * from clientes where cl_codigo =" & frm_largador.txt_mat.Text
          data_cli.Refresh
          If data_cli.Recordset.RecordCount > 0 Then
             If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                TextTelefono.Text = data_cli.Recordset("cl_telefon")
             Else
                TextTelefono.Text = frm_largador.txt_tel.Text
             End If
             If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                TextCelular.Text = data_cli.Recordset("cl_dpto")
             Else
                TextCelular.Text = ""
             End If
             'Mid(t_correo.Text, 1, 74)
             If IsNull(data_cli.Recordset("cl_referen")) = False Then
                TextMail.Text = data_cli.Recordset("cl_referen")
             Else
                TextMail.Text = ""
             End If
             'data_clientes.Recordset("cl_socmnom")
             If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
                cbomutual.Text = data_cli.Recordset("cl_socmnom")
             Else
                cbomutual.Text = ""
             End If
          Else
             TextTelefono.Text = frm_largador.txt_tel.Text
             TextCelular.Text = ""
             TextMail.Text = ""
             cbomutual.Text = ""
          End If
       Else
          If XorigenDatos = 3 Then
             data_cli.RecordSource = "select * from clientes where cl_codigo =" & frm_especialistas.t_mat.Text
             data_cli.Refresh
             If data_cli.Recordset.RecordCount > 0 Then
                If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                   TextTelefono.Text = data_cli.Recordset("cl_telefon")
                Else
                   TextTelefono.Text = frm_largador.txt_tel.Text
                End If
                If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                   TextCelular.Text = data_cli.Recordset("cl_dpto")
                Else
                   TextCelular.Text = ""
                End If
                'Mid(t_correo.Text, 1, 74)
                If IsNull(data_cli.Recordset("cl_referen")) = False Then
                   TextMail.Text = data_cli.Recordset("cl_referen")
                Else
                   TextMail.Text = ""
                End If
                'data_clientes.Recordset("cl_socmnom")
                If IsNull(data_cli.Recordset("cl_socmnom")) = False Then
                   cbomutual.Text = data_cli.Recordset("cl_socmnom")
                Else
                   cbomutual.Text = ""
                End If
             Else
                TextTelefono.Text = frm_largador.txt_tel.Text
                TextCelular.Text = ""
                TextMail.Text = ""
                cbomutual.Text = ""
             End If
          Else
             TextTelefono.Text = frmabm.txt_telef.Text
             TextCelular.Text = frmabm.t_cel.Text
             TextMail.Text = frmabm.t_correo.Text
             cbomutual.Text = frmabm.cbomutual.Text
          End If
       End If
    End If
'    frmabm.txt_telef.Text = ""
'    frmabm.t_cel.Text = ""
'    frmabm.t_correo.Text = ""
'    frmabm.cbomutual.Text = ""
    
Exit Sub
errorws:
        MsgBox "Error:" & Err.Description
        
End Sub

Private Sub cbomutual_LostFocus()
If cbomutual.Text <> "" Then
   data_mutual.Recordset.FindFirst "ca_nom ='" & cbomutual.Text & "'"
   If Not data_mutual.Recordset.NoMatch Then
   Else
      MsgBox "Mutualista no encontrada, Verifique!!"
      cbomutual.Text = ""
   End If
End If
End Sub


Public Function armarJsonBodyCliente(num_socio As String, telefono As String, celular As String, mail As String, mutualista As String) As String
     Dim JSON As String
     JSON = "{ ""clCodigo"": " & num_socio & ", ""clDpto"": """ & celular & """, ""clTelefon"": """ & telefono & """,""clReferen"": """ & mail & """, ""clSocmnom"": """ & mutualista & """}"
    armarJsonBodyCliente = JSON
End Function

Public Sub altaValidacionDatos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from clientes_history where cl_codigo =" & Val(labXmat.Caption)
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xrecclii.AddNew
Xrecclii("cl_dpto") = Trim(TextCelular.Text)
Xrecclii("cl_referen") = Trim(TextMail.Text)
Xrecclii("cl_socmnom") = cbomutual.Text
Xrecclii("cl_telefon") = Trim(TextTelefono.Text)
Xrecclii("fecha_modif") = Date
If XorigenDatos = 1 Then
   Xrecclii("origen") = "FICHA B." & data_parsec.Recordset("base")
Else
   If XorigenDatos = 2 Then
      Xrecclii("origen") = "DESPACHO"
   Else
      If XorigenDatos = 3 Then
         Xrecclii("origen") = "AGENDA B." & data_parsec.Recordset("base")
      Else
         Xrecclii("origen") = "FACTURACION"
      End If
   End If
End If
Xrecclii("usuario") = WElusuario
Xrecclii("cl_codigo") = Val(labXmat.Caption)
Xrecclii.Update

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub altaValidacionDatosabm()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from abmsocio where cl_codigo =" & Val(labXmat.Caption)
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xrecclii.AddNew
Xrecclii("usuario") = WElusuario
Xrecclii("fecha") = Date
Xrecclii("hora") = Format(Time, "HH:mm")
Xrecclii("cl_codigo") = Val(labXmat.Caption)
Xrecclii("desc") = "MODIF"
Xrecclii("cl_motivo") = "VALIDACION DATOS"
Xrecclii("convenio") = labcodconv.Caption
Xrecclii("base") = data_parsec.Recordset("base")
Xrecclii.Update


Xrecclii.Close
ConbdSapp.Close

End Sub

