VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_autoriza 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizaciones"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8865
   Icon            =   "frm_autoriza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8865
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar autorización"
      Height          =   495
      Left            =   240
      Picture         =   "frm_autoriza.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cuestionario para autorización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   8415
      Begin VB.ComboBox cboanio 
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
         Left            =   3600
         Style           =   1  'Simple Combo
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   5520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox t_obs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         MaxLength       =   150
         TabIndex        =   14
         Top             =   3360
         Width           =   5295
      End
      Begin VB.TextBox t_contact 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2640
         Width           =   5295
      End
      Begin VB.ComboBox cbocuando 
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
         ItemData        =   "frm_autoriza.frx":0B14
         Left            =   2880
         List            =   "frm_autoriza.frx":0B2A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   2535
      End
      Begin VB.ComboBox cbomes 
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
         ItemData        =   "frm_autoriza.frx":0B70
         Left            =   2880
         List            =   "frm_autoriza.frx":0B72
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cbotipodeuda 
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
         ItemData        =   "frm_autoriza.frx":0B74
         Left            =   2880
         List            =   "frm_autoriza.frx":0B81
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox cbomodulo 
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
         ItemData        =   "frm_autoriza.frx":0BAF
         Left            =   2880
         List            =   "frm_autoriza.frx":0BC2
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ingrese datos de contacto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cuando va a pagar?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Último recibo pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Deuda generada por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Módulo del sistema:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Label labmat 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aguarde a que termine el proceso....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6840
      TabIndex        =   17
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "El sistema está siendo operado por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label labfecha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7320
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema automático de autorizaciones"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frm_autoriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbotipodeuda_Click()
If cbotipodeuda.ListIndex = 0 Then
   Label4.Visible = True
   cbomes.Visible = True
   cboanio.Visible = True
Else
   Label4.Visible = False
   cbomes.Visible = False
   cboanio.Visible = False
End If


End Sub

Private Sub Command1_Click()
Dim XdeudasRef, XdeudasMes, XserviCant, i, Xerr, Xmutual, SidarCod As Integer
Dim XsumaSer As Double

Dim FechaVence As Date
XdeudasRef = 0
XdeudasMes = 0
XsumaSer = 0
Xerr = 0
Xmutual = 0
SidarCod = 0
'31978
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
If labmat.Caption <> "" And t_contact.Text <> "" Then
   frm_autoriza.MousePointer = 11
   If cbomodulo.Text = "Despacho" Then
      pb.Max = 10000
   Else
      pb.Max = 1000000
   End If
   pb.Visible = True
   Label10.Visible = True
   If Data3.Recordset.RecordCount > 0 Then
      Data3.Recordset.MoveLast
      pb.Max = pb.Max + Data3.Recordset.RecordCount
      Data3.Recordset.MoveFirst
      Do While Not Data3.Recordset.EOF
         Data3.Recordset.Delete
         pb.Value = pb.Value + 1
         Data3.Recordset.MoveNext
      Loop
   End If
   DoEvents
   If cbomodulo.Text = "Despacho" Then
      If frm_largador.txt_cat.Text = "CCNOS" Or frm_largador.txt_cat.Text = "SMIN" Or _
         frm_largador.txt_cat.Text = "UNIVS" Or frm_largador.txt_cat.Text = "HEVANO" Or frm_largador.txt_cat.Text = "GANOS" Or frm_largador.txt_cat.Text = "CASANO" Then
      Else
         data_conv.RecordSource = "select * from convenio where cnv_codigo ='" & frm_largador.txt_cat.Text & "' and cnv_grupo in ('CCOU','UNIVERSAL','CASA DE GALICIA','SMI','H.EVANGELICO')"
         data_conv.Refresh
         If data_conv.Recordset.RecordCount > 0 Then
            Xmutual = 9
         End If
      End If
   End If
   
   Data1.RecordSource = "Select * from deudas where cliente =" & Val(labmat.Caption) & " and origen >='" & "Refinan" & "' and mes =" & 0 & " and fecha_pago is null order by fecha"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveLast
      pb.Max = pb.Max + Data1.Recordset.RecordCount
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         FechaVence = Data1.Recordset("fecha") + Data1.Recordset("nro_superv") + 30
         If Format(FechaVence, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
            XdeudasRef = 1
         End If
         pb.Value = pb.Value + 1
         Data1.Recordset.MoveNext
      Loop
   End If
   Data1.RecordSource = "select * from deudas where cliente =" & Val(labmat.Caption) & " and mes >" & 0 & " and fecha_pago is null order by fecha"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveLast
      pb.Max = pb.Max + Data1.Recordset.RecordCount
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         pb.Value = pb.Value + 1
         Data1.Recordset.MoveNext
      Loop
      XdeudasMes = Data1.Recordset.RecordCount
   End If
   Data1.RecordSource = "select * from deudas where cliente =" & Val(labmat.Caption) & " and mes =" & 0 & " and origen <='" & "Refinan" & "' order by fecha"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveLast
      pb.Max = pb.Max + Data1.Recordset.RecordCount
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         XsumaSer = XsumaSer + Data1.Recordset("total")
         XserviCant = XserviCant + 1
         Data1.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
   End If
   If cbomodulo.Text = "Despacho" Then
      For i = 1 To 10000
          pb.Value = pb.Value + 1
      Next
   Else
      For i = 1 To 1000000
          pb.Value = pb.Value + 1
      Next
   End If
   frm_autoriza.MousePointer = 0
   If Xmutual = 9 Then
      Data1.RecordSource = "select * from Codigos_aut"
      Data1.Refresh
      Data1.Recordset.AddNew
      Data1.Recordset("fecha") = Date
      Data1.Recordset("usuario") = WElusuario
      Data1.Recordset("codaut") = "MUTUAL"
      Data1.Recordset("socio") = Val(labmat.Caption)
      Data1.Recordset("modulo") = cbomodulo.Text
      Data1.Recordset("usuario_caja") = WElusuario
      If cbotipodeuda.ListIndex >= 0 Then
         Data1.Recordset("tipo_deuda") = cbotipodeuda.Text
      End If
      If cbomes.Visible = True Then
         If cbomes.Text <> "" Then
            If cboanio.Text <> "" Then
               Data1.Recordset("mes_anio") = cbomes.Text & "/" & cboanio.Text
            End If
         End If
      End If
      If cbocuando.ListIndex >= 0 Then
         Data1.Recordset("cuando") = cbocuando.Text
      End If
      If t_contact.Text <> "" Then
         Data1.Recordset("contacto") = t_contact.Text
      End If
      If t_obs.Text <> "" Then
         Data1.Recordset("observa") = t_obs.Text
      End If
      Data1.Recordset.Update
      Data3.Recordset.AddNew
      Data3.Recordset("info_debit") = "AUTORIZADO COMO SOCIO MUTUAL!" & chr(13) & "RECUERDE! COBRO DE CONSULTA SI CORRESPONDE" & _
      chr(13) & "Avise al paciente que administración se comunicará."
      Data3.Recordset.Update
      Wopszond = "MUTUAL"
      frm_mensaauto.Show vbModal
   Else
        If XdeudasRef = 1 Then
'           If cbocuando.ListIndex = 5 Then
'              SidarCod = 9
'           Else
               Data3.Recordset.AddNew
               Data3.Recordset("info_debit") = "NO es posible generar código de autorización." & _
               chr(13) & " Avise al paciente se comunique con administración al 097215419 de 8a15Hs. de Lunes a Viernes"
               Data3.Recordset.Update
               frm_mensaauto.Show vbModal
               
    '           MsgBox "ATENCIÓN! NO ES POSIBLE GENERAR CÓDIGO DE AUTORIZACIÓN. DEBE COMUNICARSE CON ADMINISTRACIÓN AL 097215419 O AL SERVICIO 3030.", vbCritical, "ERROR EN AUTORIZACIÓN"
               Wopszond = ""
               Data1.RecordSource = "select * from Codigos_aut"
               Data1.Refresh
               Data1.Recordset.AddNew
               Data1.Recordset("fecha") = Date
               Data1.Recordset("usuario") = WElusuario
               Data1.Recordset("codaut") = "NO AUTORIZADO"
               Data1.Recordset("socio") = Val(labmat.Caption)
               Data1.Recordset("modulo") = cbomodulo.Text
               Data1.Recordset("usuario_caja") = WElusuario
               If cbotipodeuda.ListIndex >= 0 Then
                  Data1.Recordset("tipo_deuda") = cbotipodeuda.Text
               End If
               If cbomes.Text <> "" Then
                  If cboanio.Text <> "" Then
                     Data1.Recordset("mes_anio") = cbomes.Text & "/" & cboanio.Text
                  End If
               End If
               If cbocuando.ListIndex >= 0 Then
                  Data1.Recordset("cuando") = cbocuando.Text
               End If
               If t_contact.Text <> "" Then
                  Data1.Recordset("contacto") = t_contact.Text
               End If
               If t_obs.Text <> "" Then
                  Data1.Recordset("observa") = t_obs.Text
               End If
               Data1.Recordset.Update
'           End If
        Else
           If XdeudasMes > 3 Then
              If cbocuando.ListIndex = 4 Or cbocuando.ListIndex = 5 Or cbocuando.ListIndex = 0 Or cbocuando.ListIndex = 1 Or cbocuando.ListIndex = 2 Then
                 SidarCod = 9
              Else
                 Data3.Recordset.AddNew
                 Data3.Recordset("info_debit") = "NO es posible generar código de autorización." & _
                 chr(13) & " Avise al paciente se comunique con administración al 097215419 de 8a15Hs. de Lunes a Viernes"
                 Data3.Recordset.Update
                 frm_mensaauto.Show vbModal
'''                 MsgBox "ATENCIÓN! NO ES POSIBLE GENERAR CÓDIGO DE AUTORIZACIÓN. DEBE COMUNICARSE CON ADMINISTRACIÓN AL 097215419 O AL SERVICIO 3030.", vbCritical, "ERROR EN AUTORIZACIÓN"
                 Wopszond = ""
                 Data1.RecordSource = "select * from Codigos_aut"
                 Data1.Refresh
                 Data1.Recordset.AddNew
                 Data1.Recordset("fecha") = Date
                 Data1.Recordset("usuario") = WElusuario
                 Data1.Recordset("codaut") = "NO AUTORIZADO"
                 Data1.Recordset("socio") = Val(labmat.Caption)
                 Data1.Recordset("modulo") = cbomodulo.Text
                 Data1.Recordset("usuario_caja") = WElusuario
                 If cbotipodeuda.ListIndex >= 0 Then
                    Data1.Recordset("tipo_deuda") = cbotipodeuda.Text
                 End If
                 If cbomes.Text <> "" Then
                    If cboanio.Text <> "" Then
                       Data1.Recordset("mes_anio") = cbomes.Text & "/" & cboanio.Text
                    End If
                 End If
                 If cbocuando.ListIndex >= 0 Then
                    Data1.Recordset("cuando") = cbocuando.Text
                 End If
                 If t_contact.Text <> "" Then
                    Data1.Recordset("contacto") = t_contact.Text
                 End If
                 If t_obs.Text <> "" Then
                    Data1.Recordset("observa") = t_obs.Text
                 End If
                 Data1.Recordset.Update
              End If
           Else
              If XserviCant > 2 Then
                 If cbocuando.ListIndex = 4 Or cbocuando.ListIndex = 5 Or cbocuando.ListIndex = 0 Or cbocuando.ListIndex = 1 Or cbocuando.ListIndex = 2 Then
                    SidarCod = 9
                 Else
                    Data3.Recordset.AddNew
                    Data3.Recordset("info_debit") = "NO es posible generar código de autorización." & _
                    chr(13) & " Avise al paciente se comunique con administración al 097215419 de 8a15Hs. de Lunes a Viernes"
                    Data3.Recordset.Update
                    frm_mensaauto.Show vbModal
''                    MsgBox "ATENCIÓN! NO ES POSIBLE GENERAR CÓDIGO DE AUTORIZACIÓN. DEBE COMUNICARSE CON ADMINISTRACIÓN AL 097215419 O AL SERVICIO 3030.", vbCritical, "ERROR EN AUTORIZACIÓN"
                    Wopszond = ""
                    Data1.RecordSource = "select * from Codigos_aut"
                    Data1.Refresh
                    Data1.Recordset.AddNew
                    Data1.Recordset("fecha") = Date
                    Data1.Recordset("usuario") = WElusuario
                    Data1.Recordset("codaut") = "NO AUTORIZADO"
                    Data1.Recordset("socio") = Val(labmat.Caption)
                    Data1.Recordset("modulo") = cbomodulo.Text
                    Data1.Recordset("usuario_caja") = WElusuario
                    If cbotipodeuda.ListIndex >= 0 Then
                       Data1.Recordset("tipo_deuda") = cbotipodeuda.Text
                    End If
                    If cbomes.Text <> "" Then
                       If cboanio.Text <> "" Then
                          Data1.Recordset("mes_anio") = cbomes.Text & "/" & cboanio.Text
                       End If
                    End If
                    If cbocuando.ListIndex >= 0 Then
                       Data1.Recordset("cuando") = cbocuando.Text
                    End If
                    If t_contact.Text <> "" Then
                       Data1.Recordset("contacto") = t_contact.Text
                    End If
                    If t_obs.Text <> "" Then
                       Data1.Recordset("observa") = t_obs.Text
                    End If
                    Data1.Recordset.Update
                 End If
              Else
                 If XsumaSer > 3000 Then
                    If cbocuando.ListIndex = 4 Or cbocuando.ListIndex = 5 Or cbocuando.ListIndex = 0 Or cbocuando.ListIndex = 1 Or cbocuando.ListIndex = 2 Then
                       SidarCod = 9
                    Else
                        Data3.Recordset.AddNew
                        Data3.Recordset("info_debit") = "NO es posible generar código de autorización." & _
                        chr(13) & " Avise al paciente se comunique con administración al 097215419 de 8a15Hs. de Lunes a Viernes"
                        Data3.Recordset.Update
                        frm_mensaauto.Show vbModal
'                        MsgBox "ATENCIÓN! NO ES POSIBLE GENERAR CÓDIGO DE AUTORIZACIÓN. DEBE COMUNICARSE CON ADMINISTRACIÓN AL 097215419 O AL SERVICIO 3030.", vbCritical, "ERROR EN AUTORIZACIÓN"
                        Wopszond = ""
                        Data1.RecordSource = "select * from Codigos_aut"
                        Data1.Refresh
                        Data1.Recordset.AddNew
                        Data1.Recordset("fecha") = Date
                        Data1.Recordset("usuario") = WElusuario
                        Data1.Recordset("codaut") = "NO AUTORIZADO"
                        Data1.Recordset("socio") = Val(labmat.Caption)
                        Data1.Recordset("modulo") = cbomodulo.Text
                        Data1.Recordset("usuario_caja") = WElusuario
                        If cbotipodeuda.ListIndex >= 0 Then
                           Data1.Recordset("tipo_deuda") = cbotipodeuda.Text
                        End If
                        If cbomes.Text <> "" Then
                           If cboanio.Text <> "" Then
                              Data1.Recordset("mes_anio") = cbomes.Text & "/" & cboanio.Text
                           End If
                        End If
                        If cbocuando.ListIndex >= 0 Then
                           Data1.Recordset("cuando") = cbocuando.Text
                        End If
                        If t_contact.Text <> "" Then
                           Data1.Recordset("contacto") = t_contact.Text
                        End If
                        If t_obs.Text <> "" Then
                           Data1.Recordset("observa") = t_obs.Text
                        End If
                        Data1.Recordset.Update
                    End If
                 Else
                    Data2.Recordset.Edit
                    Data2.Recordset("codigo") = Data2.Recordset("codigo") + 1
                    Data2.Recordset.Update
                    Data1.RecordSource = "select * from Codigos_aut"
                    Data1.Refresh
                    Data1.Recordset.AddNew
                    Data1.Recordset("fecha") = Date
                    Data1.Recordset("usuario") = WElusuario
                    Data1.Recordset("codaut") = Trim(str(Data2.Recordset("codigo")))
                    Data1.Recordset("socio") = Val(labmat.Caption)
                    Data1.Recordset("modulo") = cbomodulo.Text
                    Data1.Recordset("usuario_caja") = WElusuario
                    If cbotipodeuda.ListIndex >= 0 Then
                       Data1.Recordset("tipo_deuda") = cbotipodeuda.Text
                    Else
                       Xerr = 1
                    End If
                    If cbomes.Visible = True Then
                       If cbomes.Text <> "" Then
                          If cboanio.Text <> "" Then
                             Data1.Recordset("mes_anio") = cbomes.Text & "/" & cboanio.Text
                          Else
                             Xerr = 1
                          End If
                       Else
                          Xerr = 1
                       End If
                    End If
                    If cbocuando.ListIndex >= 0 Then
                       Data1.Recordset("cuando") = cbocuando.Text
                    Else
                       Xerr = 1
                    End If
                    If t_contact.Text <> "" Then
                       Data1.Recordset("contacto") = t_contact.Text
                    Else
                       Xerr = 1
                    End If
                    If t_obs.Text <> "" Then
                       Data1.Recordset("observa") = t_obs.Text
                    End If
                    If Xerr = 1 Then
                       Data1.Recordset.CancelUpdate
                       MsgBox "FALTAN DATOS", vbCritical
                       Wopszond = ""
                    Else
                       Data1.Recordset.Update
                       Data3.Recordset.AddNew
                       Data3.Recordset("info_debit") = "Autorización realizada con éxito. " & chr(13) & " NÚMERO DE AUTORIZACIÓN:" & chr(13) _
                       & Data2.Recordset("codigo")
                       Data3.Recordset.Update
                       Wopszond = Trim(str(Data2.Recordset("codigo")))
                       frm_mensaauto.Show vbModal
                    End If
                 End If
              End If
           End If
           If SidarCod = 9 Then
                Data2.Recordset.Edit
                Data2.Recordset("codigo") = Data2.Recordset("codigo") + 1
                Data2.Recordset.Update
                Data1.RecordSource = "select * from Codigos_aut"
                Data1.Refresh
                Data1.Recordset.AddNew
                Data1.Recordset("fecha") = Date
                Data1.Recordset("usuario") = WElusuario
                Data1.Recordset("codaut") = Trim(str(Data2.Recordset("codigo")))
                Data1.Recordset("socio") = Val(labmat.Caption)
                Data1.Recordset("modulo") = cbomodulo.Text
                Data1.Recordset("usuario_caja") = WElusuario
                If cbotipodeuda.ListIndex >= 0 Then
                   Data1.Recordset("tipo_deuda") = cbotipodeuda.Text
                Else
                   Xerr = 1
                End If
                If cbomes.Visible = True Then
                   If cbomes.Text <> "" Then
                      If cboanio.Text <> "" Then
                         Data1.Recordset("mes_anio") = cbomes.Text & "/" & cboanio.Text
                      Else
                         Xerr = 1
                      End If
                   Else
                      Xerr = 1
                   End If
                End If
                If cbocuando.ListIndex >= 0 Then
                   Data1.Recordset("cuando") = cbocuando.Text
                Else
                   Xerr = 1
                End If
                If t_contact.Text <> "" Then
                   Data1.Recordset("contacto") = t_contact.Text
                Else
                   Xerr = 1
                End If
                If t_obs.Text <> "" Then
                   Data1.Recordset("observa") = t_obs.Text
                End If
                If Xerr = 1 Then
                   Data1.Recordset.CancelUpdate
                   MsgBox "FALTAN DATOS", vbCritical
                   Wopszond = ""
                Else
                   Data1.Recordset.Update
                   Data3.Recordset.AddNew
                   Data3.Recordset("info_debit") = "Autorización realizada con éxito. " & chr(13) & " NÚMERO DE AUTORIZACIÓN:" & chr(13) _
                   & Data2.Recordset("codigo")
                   Data3.Recordset.Update
                   Wopszond = Trim(str(Data2.Recordset("codigo")))
                   frm_mensaauto.Show vbModal
                End If
           
           End If
        End If
   End If
Else
   If Trim(t_contact.Text) = "" Then
      MsgBox "No ingresó contacto, no se puede generar autorización", vbCritical
   End If
End If
SidarCod = 0

Label10.Visible = False
Unload Me

'qué hacer cuando no se genera codigo de autorización y luego se autoriza


End Sub

Private Sub Form_Load()
'Socio mutual se autoriza y los complementos con deuda salen a costo parcial

labfecha.Caption = Format(Date, "dd/mm/yyyy")
Label9.Caption = WElusuario
If Xdeb = 1 Then
   cbomodulo.ListIndex = 0
   If frm_largador.txt_tel.Text <> "" Then
      t_contact.Text = frm_largador.txt_tel.Text
   End If
Else
   If Xdeb = 3 Then
      cbomodulo.ListIndex = 1
   Else
      If Xdeb = 4 Then
         cbomodulo.ListIndex = 3
      Else
         If Xdeb = 5 Then
            cbomodulo.ListIndex = 4
         Else
            cbomodulo.ListIndex = 1
         End If
      End If
   End If
End If

labmat.Caption = Xhab

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data2.RecordSource = "codigos"
Data2.Refresh
Data3.DatabaseName = App.path & "\informes.mdb"
Data3.RecordSource = "infcli"
Data3.Refresh
cbomes.AddItem "1"
cbomes.AddItem "2"
cbomes.AddItem "3"
cbomes.AddItem "4"
cbomes.AddItem "5"
cbomes.AddItem "6"
cbomes.AddItem "7"
cbomes.AddItem "8"
cbomes.AddItem "9"
cbomes.AddItem "10"
cbomes.AddItem "11"
cbomes.AddItem "12"
cboanio.AddItem "2018"
cboanio.AddItem "2019"
cboanio.AddItem "2020"
If Wxquepreg = 2 Then
   cbotipodeuda.ListIndex = 0
   cbomes.Text = Xop4
   cboanio.Text = Xop5
Else
   If Wxquepreg = 3 Then
      cbotipodeuda.ListIndex = 2
   Else
      cbotipodeuda.ListIndex = 1
   End If
End If
cbotipodeuda.Enabled = False

   

End Sub

Private Sub Timer1_Timer()

End Sub

