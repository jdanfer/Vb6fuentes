VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_servapacc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acciones de Servicios AP"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11295
   Icon            =   "frm_servapacc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_servapacc.frx":058A
      Height          =   2055
      Left            =   240
      OleObjectBlob   =   "frm_servapacc.frx":059E
      TabIndex        =   10
      ToolTipText     =   "Doble click para ver el registro seleccionado."
      Top             =   4920
      Width           =   10815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registro de acciones"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10815
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   7560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   10320
         Picture         =   "frm_servapacc.frx":12C9
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Eliminar el registro seleccionado"
         Top             =   4080
         Width           =   375
      End
      Begin VB.Data data_usuar 
         Caption         =   "data_usuar"
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
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Enabled         =   0   'False
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         Picture         =   "frm_servapacc.frx":1853
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4080
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         Picture         =   "frm_servapacc.frx":1DDD
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Registrar una nueva acción"
         Top             =   4080
         Width           =   375
      End
      Begin VB.TextBox t_det 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2040
         Width           =   8535
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_servapacc.frx":2367
         Left            =   2160
         List            =   "frm_servapacc.frx":2377
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox t_nro 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label labnroacc 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label labusuar 
         BackColor       =   &H00404040&
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
         Left            =   8040
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Detalle:"
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
         TabIndex        =   7
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Opción de la acción:"
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
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label labtitulo 
         BackColor       =   &H00404040&
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
         Left            =   2160
         TabIndex        =   4
         Top             =   840
         Width           =   8415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Título:"
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
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Nro. de Servicio:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frm_servapacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_det.SetFocus
End If

End Sub

Private Sub Command1_Click()
If frm_servap.labsi.Caption = "N" Then
   MsgBox "No figura como destinatario.", vbCritical
   Command3.Enabled = False
   Command2.Enabled = False
Else
    If frm_servap.mfh.Text = "__/__/____" Then
        t_nro.Text = frm_servap.txt_nro.Text
        labusuar.Caption = WElusuario
        labtitulo.Caption = frm_servap.Combo1.Text
        Combo1.ListIndex = -1
        t_det.Text = ""
        Combo1.SetFocus
        Command1.Enabled = False
        Command2.Enabled = True
        DBGrid1.Enabled = False
        labnroacc.Caption = ""
        Command3.Enabled = False
    Else
        MsgBox "El registro se encuentra cerrado.", vbCritical
        Command3.Enabled = False
        Command2.Enabled = False
    End If
End If

End Sub

Private Sub Command2_Click()
Dim EnviarCorreo, CorreoAP, textocorreo As String
Dim Noenvia As Integer
On Error GoTo Nohaydatos

Noenvia = 0
EnviarCorreo = ""
textocorreo = ""
CorreoAP = ""

If frm_servap.labsi.Caption = "N" Then
   MsgBox "No es destinatario del registro.", vbCritical
Else
    If Combo1.ListIndex >= 0 Then
       If Trim(t_det.Text) <> "" Then
          If Trim(labnroacc.Caption) <> "" Then
            If Data1.Recordset("usuario") = WElusuario Then
                Data1.Recordset.Edit
                Data1.Recordset("fecha") = Date
                Data1.Recordset("hora") = Format(Time, "HH:mm")
                Data1.Recordset("detalle") = t_det.Text
                Data1.Recordset("usuario") = WElusuario
                Data1.Recordset("op_ind") = Combo1.ListIndex
                Data1.Recordset("op_desc") = Combo1.Text
                Data1.Recordset.Update
                Data3.RecordSource = "select * from env_soc where cl_codigo =" & t_nro.Text
                Data3.Refresh
                If Data3.Recordset.RecordCount > 0 Then
                   If Combo1.Text = "TESORERIA" Then
                      If IsNull(Data3.Recordset("ci_tarj")) = False Then
                         If Data3.Recordset("ci_tarj") = 0 Then
                            Data3.Recordset.Edit
                            Data3.Recordset("ci_tarj") = 1
                            Data3.Recordset.Update
                         End If
                      End If
                   End If
                   If Combo1.Text = "HABILITACION DT" Then
                      If IsNull(Data3.Recordset("ultmespmut")) = False Then
                         If Data3.Recordset("ultmespmut") = 0 Then
                            Data3.Recordset.Edit
                            Data3.Recordset("ultmespmut") = 1
                            Data3.Recordset.Update
                         End If
                      End If
                   End If
                   If Combo1.Text = "PRESUPUESTO" Then
                      If IsNull(Data3.Recordset("ultanopmut")) = False Then
                         If Data3.Recordset("ultanopmut") = 0 Then
                            Data3.Recordset.Edit
                            Data3.Recordset("ultanopmut") = 1
                            Data3.Recordset.Update
                         End If
                      End If
                   End If
                   If Combo1.Text = "OTROS" Then
                      If IsNull(Data3.Recordset("cl_mes_ant")) = False Then
                         If Data3.Recordset("cl_mes_ant") = 0 Then
                            Data3.Recordset.Edit
                            Data3.Recordset("cl_mes_ant") = 1
                            Data3.Recordset.Update
                         End If
                      End If
                   End If
                
                End If
                textocorreo = "MODIFICACION REGISTRADA POR: " & WElusuario & vbCrLf
                textocorreo = textocorreo & "OPCIÓN: " & Combo1.Text & vbCrLf
                textocorreo = textocorreo & "DETALLE: " & t_det.Text
                
                EnviarCorreo = MsgBox("Desea enviar correo a los otros destinatarios?", vbInformation + vbYesNo)
                If EnviarCorreo = vbYes Then
                   If Trim(textocorreo) <> "" Then
                      frm_servapacc.MousePointer = 11
                      Dim MenCorreo2 As String
                      Dim oMail2 As Class1
                      frm_servap.List1.ListIndex = 0
                      For X = 1 To frm_servap.List1.ListCount
                          data_usuar.RecordSource = "Select * from usuarios where nombre ='" & frm_servap.List1.List(frm_servap.List1.ListIndex) & "' and serv_ap in ('S')"
                          data_usuar.Refresh
                          If data_usuar.Recordset.RecordCount > 0 Then
                             If data_usuar.Recordset("usuario") = WElusuario Then
                                Noenvia = 1
                             Else
                                Noenvia = 0
                             End If
                             If IsNull(data_usuar.Recordset("correo_ap")) = False Then
                                CorreoAP = data_usuar.Recordset("correo_ap")
                             Else
                                CorreoAP = "jdanfer@gmail.com"
                             End If
                          Else
                             CorreoAP = "jdanfer@gmail.com"
                          End If
                          If Noenvia = 1 Then
                             If frm_servap.List1.ListCount - 1 = frm_servap.List1.ListIndex Then
                             Else
                                frm_servap.List1.ListIndex = frm_servap.List1.ListIndex + 1
                             End If
                          Else
                            Set oMail2 = New Class1
                                With oMail2
                                  .servidor = "smtp.office365.com"
                                  .puerto = 25
                                  .UseAuntentificacion = True
                                  .ssl = True
                                  .Usuario = "jefedepartamentoti@sapp.com.uy"
                                  .PassWord = "DptotiJunio2021"
                                  .Asunto = frm_servap.List1.List(frm_servap.List1.ListIndex) & " Se ha MODIFICADO una acción de: " & Combo1.Text & " Servicio AP Nro:" & t_nro.Text
                                  .de = "jefedepartamentoti@sapp.com.uy"
                                  .para = CorreoAP
                     '             .Adjunto = Xarchtex
                                  .Mensaje = textocorreo
                                  .Enviar_Backup ' manda el mail
                                End With
                               Set oMail2 = Nothing
                               If frm_servap.List1.ListCount - 1 = frm_servap.List1.ListIndex Then
                               Else
                                  frm_servap.List1.ListIndex = frm_servap.List1.ListIndex + 1
                               End If
                          End If
                      Next
                      frm_servapacc.MousePointer = 0
                      MsgBox "Correos enviados!", vbInformation
                   Else
                      MsgBox "No hay texto para enviar correo, verifique si hay datos ingresados.", vbCritical
                   End If
                End If
                frm_servapacc.MousePointer = 0
                Data1.Refresh
                Command2.Enabled = False
                Command3.Enabled = False
                Command1.Enabled = True
                DBGrid1.Enabled = True
                labnroacc.Caption = ""
             Else
                MsgBox "No es el usuario creador de la acción.", vbCritical
             End If
          Else
            Data1.Recordset.AddNew
            Data1.Recordset("fecha") = Date
            Data1.Recordset("hora") = Format(Time, "HH:mm")
            Data1.Recordset("nro_acc") = t_nro.Text
            Data1.Recordset("detalle") = t_det.Text
            Data1.Recordset("usuario") = WElusuario
            Data1.Recordset("op_ind") = Combo1.ListIndex
            Data1.Recordset("op_desc") = Combo1.Text
            Data1.Recordset.Update
            Data3.RecordSource = "select * from env_soc where cl_codigo =" & t_nro.Text
            Data3.Refresh
            If Data3.Recordset.RecordCount > 0 Then
               If Combo1.Text = "TESORERIA" Then
                  If IsNull(Data3.Recordset("ci_tarj")) = False Then
                     If Data3.Recordset("ci_tarj") = 0 Then
                        Data3.Recordset.Edit
                        Data3.Recordset("ci_tarj") = 1
                        Data3.Recordset.Update
                     End If
                  End If
               End If
               If Combo1.Text = "HABILITACION DT" Then
                  If IsNull(Data3.Recordset("ultmespmut")) = False Then
                     If Data3.Recordset("ultmespmut") = 0 Then
                        Data3.Recordset.Edit
                        Data3.Recordset("ultmespmut") = 1
                        Data3.Recordset.Update
                     End If
                  End If
               End If
               If Combo1.Text = "PRESUPUESTO" Then
                  If IsNull(Data3.Recordset("ultanopmut")) = False Then
                     If Data3.Recordset("ultanopmut") = 0 Then
                        Data3.Recordset.Edit
                        Data3.Recordset("ultanopmut") = 1
                        Data3.Recordset.Update
                     End If
                  End If
               End If
               If Combo1.Text = "OTROS" Then
                  If IsNull(Data3.Recordset("cl_mes_ant")) = False Then
                     If Data3.Recordset("cl_mes_ant") = 0 Then
                        Data3.Recordset.Edit
                        Data3.Recordset("cl_mes_ant") = 1
                        Data3.Recordset.Update
                     End If
                  End If
               End If
            End If
            textocorreo = "ACCIÓN REGISTRADA POR: " & WElusuario & vbCrLf
            textocorreo = textocorreo & "OPCIÓN: " & Combo1.Text & vbCrLf
            textocorreo = textocorreo & "DETALLE: " & t_det.Text
            
            EnviarCorreo = MsgBox("Desea enviar correo a los otros destinatarios?", vbInformation + vbYesNo)
            If EnviarCorreo = vbYes Then
               frm_servapacc.MousePointer = 11
               Dim MenCorreo As String
               Dim oMail As Class1
               frm_servap.List1.ListIndex = 0
               For X = 1 To frm_servap.List1.ListCount
                   data_usuar.RecordSource = "Select * from usuarios where nombre ='" & frm_servap.List1.List(frm_servap.List1.ListIndex) & "' and serv_ap in ('S')"
                   data_usuar.Refresh
                   If data_usuar.Recordset.RecordCount > 0 Then
                      If data_usuar.Recordset("usuario") = WElusuario Then
                         Noenvia = 1
                      Else
                         Noenvia = 0
                      End If
                      If IsNull(data_usuar.Recordset("correo_ap")) = False Then
                         CorreoAP = data_usuar.Recordset("correo_ap")
                      Else
                         CorreoAP = "jdanfer@gmail.com"
                      End If
                   Else
                      CorreoAP = "jdanfer@gmail.com"
                   End If
                   If Noenvia = 1 Then
                      If frm_servap.List1.ListCount - 1 = frm_servap.List1.ListIndex Then
                      Else
                         frm_servap.List1.ListIndex = frm_servap.List1.ListIndex + 1
                      End If
                   Else
                     Set oMail = New Class1
                         With oMail
                           .servidor = "smtp.office365.com"
                           .puerto = 25
                           .UseAuntentificacion = True
                           .ssl = True
                           .Usuario = "jefedepartamentoti@sapp.com.uy"
                           .PassWord = "DptotiJunio2021"
                           .Asunto = frm_servap.List1.List(frm_servap.List1.ListIndex) & " Se ha registrado una acción de: " & Combo1.Text & " Servicio AP Nro:" & t_nro.Text
                           .de = "jefedepartamentoti@sapp.com.uy"
                           .para = CorreoAP
              '             .Adjunto = Xarchtex
                           .Mensaje = textocorreo
                           .Enviar_Backup ' manda el mail
                         End With
                        Set oMail = Nothing
                        If frm_servap.List1.ListCount - 1 = frm_servap.List1.ListIndex Then
                        Else
                           frm_servap.List1.ListIndex = frm_servap.List1.ListIndex + 1
                        End If
                   End If
               Next
               frm_servapacc.MousePointer = 0
               MsgBox "Correos enviados!", vbInformation
            End If
            frm_servapacc.MousePointer = 0
            Data1.Refresh
            Command2.Enabled = False
            Command1.Enabled = True
            Command3.Enabled = False
            DBGrid1.Enabled = True
            labnroacc.Caption = ""
          End If
       Else
          MsgBox "Falta detalles"
       End If
    Else
       MsgBox "Falta opción."
    End If
End If

Exit Sub

Nohaydatos:
            If Err.Number = 3150 Then
               MsgBox "ERROR:" & Err.Description
            Else
               MsgBox "ERROR:" & Err.Description
            End If
            
End Sub

Private Sub Command3_Click()
Dim Deseaborrar As String

If Data1.Recordset("usuario") = WElusuario Then
   Deseaborrar = MsgBox("Desea borrar el registro seleccionado?", vbInformation + vbYesNo, "Borrar")
   If Deseaborrar = vbYes Then
      Data1.Recordset.Delete
      Data1.Refresh
      MsgBox "Registro eliminado."
   End If
Else
   MsgBox "No es el usuario creador de la acción.", vbCritical
End If
Command3.Enabled = False
Command2.Enabled = False

      
End Sub

Private Sub DBGrid1_DblClick()

t_nro.Text = Data1.Recordset("nro_acc")
labusuar.Caption = Data1.Recordset("usuario")
labtitulo.Caption = frm_servap.Combo1.Text
Combo1.ListIndex = Data1.Recordset("op_ind")
t_det.Text = Data1.Recordset("detalle")
labnroacc.Caption = Data1.Recordset("id")
If frm_servap.mfh.Text = "__/__/____" Then
   Command2.Enabled = True
   Command3.Enabled = True
Else
   MsgBox "El registro se encuentra cerrado.", vbCritical
   Command2.Enabled = False
   Command3.Enabled = False
End If

End Sub

Private Sub Form_Load()

Data1.Connect = "odbc;dsn=sappespecial;"
Data1.RecordSource = "select * from servap_acc where nro_acc =" & frm_servap.txt_nro.Text & " order by fecha"
Data1.Refresh

Data3.Connect = "odbc;dsn=sappespecial;"

Data2.Connect = "odbc;dsn=sappespecial;"
data_usuar.Connect = "odbc;dsn=sappnew;"


End Sub
