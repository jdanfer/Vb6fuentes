VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_buslla 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   6330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
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
      Top             =   5040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_aut 
      Caption         =   "data_aut"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc data_convbus 
      Height          =   495
      Left            =   1080
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "data_convbus"
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
   Begin VB.Data data_deuda 
      Caption         =   "data_deuda"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cbobus 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frm_buslla.frx":0000
      Left            =   2160
      List            =   "frm_buslla.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton b_cierra 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H0000FF00&
      Picture         =   "frm_buslla.frx":0045
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox txt_buscacli 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buslla.frx":05CF
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frm_buslla.frx":05E3
      TabIndex        =   2
      Top             =   840
      Width           =   11655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "BUSCAR POR..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2160
      Picture         =   "frm_buslla.frx":19BE
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "frm_buslla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_cierra_Click()
frm_buslla.Hide

End Sub

Private Sub cbobus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_buscacli.SetFocus
End If

End Sub

Private Sub DBGrid1_DblClick()
Dim Xelcodigoaut, Xlapersona As String
Dim Xrecconve As New ADODB.Recordset
Dim Xsqlstr As String
Dim Xcodzoning As Integer
Dim Xaltaanterior As Integer
Dim MensajeClave3 As String

Xaltaanterior = XAlta
XwYalomostro = 99
frm_largador.txt_mat.Text = Data1.Recordset("cl_codigo")

If IsNull(Data1.Recordset("estado")) = False Then
   If Data1.Recordset("estado") = 2 Or Data1.Recordset("estado") = 3 Then
      If IsNull(Data1.Recordset("fecha_baja")) = False Then
         MsgBox "ATENCION!! SOCIO FIGURA DE BAJA con FECHA: " & Format(Data1.Recordset("fecha_baja"), "dd/mm/yyyy") & " Comuníquese con administración al 097215419", vbCritical, "Mensaje"
      Else
         MsgBox "ATENCION!! SOCIO FIGURA DE BAJA. Comuníquese con administración al 097215419", vbCritical, "Mensaje"
      End If
   End If
End If


Dim XXlafecdecons As Date
XXlafecdecons = Date - 2
If IsNull(Data1.Recordset("cl_cedula")) = False Then
   If Data1.Recordset("cl_cedula") > 0 Then
      frm_largador.data_cons.RecordSource = "Select * from consmas where mat =" & Data1.Recordset("cl_cedula") & " and fecha >=#" & Format(XXlafecdecons, "yyyy/mm/dd") & "#"
      frm_largador.data_cons.Refresh
      If frm_largador.data_cons.Recordset.RecordCount > 0 Then
         frm_largador.data_cons.Recordset.MoveLast
         MsgBox "CONSULTAS EN LAS 48HS -ULTIMA CONSULTA EL: " & Format(frm_largador.data_cons.Recordset("fecha"), "dd/mm/yyyy") & " POR: " & frm_largador.data_cons.Recordset("motivo"), vbInformation, "Mensaje"
      End If
   End If
End If

frm_largador.txt_nomb.Text = Data1.Recordset("cl_apellid")
'frm_largador.data_ref.Recordset.FindFirst "mat =" & Data1.Recordset("cl_codigo")
frm_largador.data_ref.RecordSource = "Select * from referen where mat =" & Data1.Recordset("cl_codigo")
frm_largador.data_ref.Refresh
'If Not frm_largador.data_ref.Recordset.NoMatch Then
If IsNull(Data1.Recordset("cl_direcci")) = False Then
   If IsNull(Data1.Recordset("cl_zona")) = False Then
      If Data1.Recordset("cl_zona") <> "*TODOS" Then
         frm_largador.txt_direc.Text = Data1.Recordset("cl_direcci") & "//" & Data1.Recordset("cl_zona")
      Else
         frm_largador.txt_direc.Text = Data1.Recordset("cl_direcci")
      End If
   Else
      frm_largador.txt_direc.Text = Data1.Recordset("cl_direcci")
   End If
Else

End If

If frm_largador.data_ref.Recordset.RecordCount > 0 Then
   If IsNull(frm_largador.data_ref.Recordset("refmat")) = False Then
      frm_largador.txt_direc.Text = frm_largador.data_ref.Recordset("refmat")
   End If
Else
End If
If IsNull(Data1.Recordset("cl_edad")) = False Then
   frm_largador.txt_mat.Text = Data1.Recordset("cl_codigo")
   If IsNull(Data1.Recordset("cl_edad")) = False Then
      frm_largador.txt_edad.Text = Data1.Recordset("cl_edad")
   Else
      frm_largador.txt_edad.Text = 0
   End If
   If IsNull(Data1.Recordset("cl_uniedad")) = False Then
      If Data1.Recordset("cl_uniedad") = "A" Then
         frm_largador.cboed.ListIndex = 0
      Else
         If Data1.Recordset("cl_uniedad") = "M" Then
            frm_largador.cboed.ListIndex = 1
         Else
            If Data1.Recordset("cl_uniedad") = "D" Then
               frm_largador.cboed.ListIndex = 2
            Else
               frm_largador.cboed.ListIndex = 0
            End If
         End If
      End If
   Else
      frm_largador.cboed.ListIndex = 0
   End If
Else
   frm_largador.txt_edad.Text = 0
   frm_largador.cboed.ListIndex = 0
End If
If IsNull(Data1.Recordset("cl_codconv")) = False Then
   frm_largador.txt_cat.Text = Data1.Recordset("cl_codconv")
   data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & UCase(Data1.Recordset("cl_codconv")) & "' and cnv_umpago not in (1)"
   data_convbus.Refresh
   If data_convbus.Recordset.RecordCount > 0 Then
      If IsNull(data_convbus.Recordset("cnv_fbaja")) = False Then
         MsgBox "ATENCION!! El convenio figura de BAJA, VERIFIQUE con Atención al Socio!!!", vbCritical
         MsgBox "Se ingresará cómo categoría PARTICULAR"
         frm_largador.txt_cat.Text = "PART"
         frm_largador.txt_nomcat.Text = "PARTICULARES"
      End If
   Else
      frm_largador.txt_cat.Text = "PART"
      frm_largador.txt_nomcat.Text = "PARTICULARES"
   End If
End If

If IsNull(Data1.Recordset("cl_nomconv")) = False Then
   frm_largador.txt_nomcat.Text = Data1.Recordset("cl_nomconv")
End If
If IsNull(Data1.Recordset("cl_zona")) = False Then
   If Data1.Recordset("cl_zona") = "*TODOS" Then
      frm_largador.txt_locali.Text = ""
   Else
      frm_largador.txt_locali.Text = Data1.Recordset("cl_zona")
   End If
Else
   frm_largador.txt_locali.Text = ""
End If
If IsNull(Data1.Recordset("cl_cedula")) = False Then
   frm_largador.txt_ced.Text = Int(Data1.Recordset("cl_cedula"))
Else
   frm_largador.txt_ced.Text = 0
End If
If IsNull(Data1.Recordset("cl_codced")) = False Then
   frm_largador.t_codced.Text = Int(Data1.Recordset("cl_codced"))
Else
   frm_largador.t_codced.Text = 0
End If
If IsNull(Data1.Recordset("cl_dpto")) = False Then
   If IsNull(Data1.Recordset("cl_telefon")) = False Then
      frm_largador.txt_tel.Text = Data1.Recordset("cl_dpto") & "//" & Data1.Recordset("cl_telefon")
   Else
      frm_largador.txt_tel.Text = Data1.Recordset("cl_dpto")
   End If
Else
   If IsNull(Data1.Recordset("cl_telefon")) = False Then
      frm_largador.txt_tel.Text = Data1.Recordset("cl_telefon")
   Else
   
   End If
End If
If IsNull(Data1.Recordset("cl_sexo")) = False Then
   If Data1.Recordset("cl_sexo") = 1 Then
      frm_largador.Combo3.ListIndex = 0
   Else
      frm_largador.Combo3.ListIndex = 1
   End If
Else
   frm_largador.Combo3.ListIndex = 0
End If
If IsNull(Data1.Recordset("cl_grupo")) = False Then
   If Data1.Recordset("cl_grupo") >= 100 And Data1.Recordset("cl_grupo") <= 530 Then
      frm_largador.cbozona.Text = "1"
   Else
      If Data1.Recordset("cl_grupo") >= 600 And Data1.Recordset("cl_grupo") <= 689 Then
         frm_largador.cbozona.Text = "2"
      Else
         If Data1.Recordset("cl_grupo") >= 700 And Data1.Recordset("cl_grupo") <= 788 Then
            frm_largador.cbozona.Text = "1"
         Else
            frm_largador.cbozona.Text = "2"
         End If
      End If
   End If
Else
   frm_largador.cbozona.Text = "1"
End If
If Data1.Recordset.RecordCount > 0 Then
   If IsNull(Data1.Recordset("cl_grupo")) = False Then
      Xcodzoning = Data1.Recordset("cl_grupo")
   Else
      Xcodzoning = 0
   End If
Else
   Xcodzoning = 0
End If
          
          
          
    Wxquepreg = 0
    Wopszond = ""
    Xop4 = 0
    Xop5 = 0
    Xhab = Data1.Recordset("cl_codigo")
    Dim Xq As Integer
    If frm_largador.Check1.Value = 1 Then
        data_deuda.RecordSource = "Select * from deudas where cliente =" & Data1.Recordset("cl_codigo") & " and mes >" & 0 & " and fecha_pago is null order by ano,mes"
        data_deuda.Refresh
        If data_deuda.Recordset.RecordCount > 0 Then
           data_deuda.Recordset.MoveLast
           If data_deuda.Recordset.RecordCount > 2 Then
              Xop4 = data_deuda.Recordset("mes")
              Xop5 = data_deuda.Recordset("ano")
              Xq = 9
              Wxquepreg = 2 'Deuda por cuota
           End If
        End If
    Else
        If Trim(Data1.Recordset("cl_codconv")) = "" Then
           data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & "AABB" & "'"
           data_convbus.Refresh
        Else
           data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & Data1.Recordset("cl_codconv") & "' and cnv_sindeuda in (1) and cnv_fbaja is null"
           data_convbus.Refresh
        End If
        If data_convbus.Recordset.RecordCount > 0 Then
           Xq = 0
        Else
             If Data1.Recordset.RecordCount > 0 Then
                Dim Xladat, Xhoy As Date
                Xhoy = Date
                Xq = 0
               data_deuda.RecordSource = "Select * from deudas where cliente =" & Data1.Recordset("cl_codigo") & " and tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanciacion" & "' and mes =" & 0
               data_deuda.Refresh
               If data_deuda.Recordset.RecordCount > 0 Then
                  data_deuda.Recordset.MoveFirst
                  Do While Not data_deuda.Recordset.EOF
                     If IsNull(data_deuda.Recordset("nro_superv")) = False Then
                        Xladat = data_deuda.Recordset("fecha") + data_deuda.Recordset("nro_superv")
                     Else
                        Xladat = data_deuda.Recordset("fecha") + 15
                     End If
                     If Format(Xladat, "yyyy/mm/dd") < Format(Xhoy, "yyyy/mm/dd") Then
                        Xq = 9
                        Wxquepreg = 1 'es deuda por servicio
                     End If
                     data_deuda.Recordset.MoveNext
                  Loop
               End If
               data_deuda.RecordSource = "Select * from deudas where cliente =" & Data1.Recordset("cl_codigo") & " and mes >" & 0 & " and fecha_pago is null and origen <='" & "Refinan" & "' order by ano,mes"
               data_deuda.Refresh
               If data_deuda.Recordset.RecordCount > 0 Then
                  data_deuda.Recordset.MoveLast
                  If data_deuda.Recordset.RecordCount > 2 Then
                     Xop4 = data_deuda.Recordset("mes")
                     Xop5 = data_deuda.Recordset("ano")
                     Xq = 9
                     If Wxquepreg = 0 Then
                        Wxquepreg = 2 'es por cuota
                     End If
                  End If
               End If
               data_deuda.RecordSource = "Select * from deudas where cliente =" & Data1.Recordset("cl_codigo") & " and fecha_pago is null and origen >='" & "Refinan" & "'"
               data_deuda.Refresh
               If data_deuda.Recordset.RecordCount > 0 Then
                  data_deuda.Recordset.MoveFirst
                  Do While Not data_deuda.Recordset.EOF
                     If IsNull(data_deuda.Recordset("nro_superv")) = False Then
                        Xladat = data_deuda.Recordset("fecha") + data_deuda.Recordset("nro_superv")
                     Else
                        Xladat = data_deuda.Recordset("fecha") + 30
                     End If
                     If Format(Xladat, "yyyy/mm/dd") < Format(Xhoy, "yyyy/mm/dd") Then
                        Xq = 9
                        Wxquepreg = 3 'es por refinanc
                     End If
                     data_deuda.Recordset.MoveNext
                  Loop
               End If
               
               If Xq = 9 Then
                  XAlta = 599
                  Xtot = Data1.Recordset("cl_codigo")
                  Xhab = Data1.Recordset("cl_codigo")
                  frm_veodeuda.Show vbModal
                                
                  Xdeb = 1
                  MensajeClave3 = MsgBox("PACIENTE CON DEUDA! ES UN LLAMADO DE URGENCIA?", vbExclamation + vbYesNo + vbDefaultButton2)
                  
                  If MensajeClave3 = vbYes Then
                     Xelcodigoaut = "URGENCIA"
                     Xq = 0
                     Xdeudasi = 0
                     data_aut.RecordSource = "select * from Codigos_aut"
                     data_aut.Refresh
                     data_aut.Recordset.AddNew
                     data_aut.Recordset("fecha") = Date
                     data_aut.Recordset("usuario") = Mid(Data1.Recordset("cl_apellid"), 1, 50)
                     data_aut.Recordset("codaut") = "URGENCIA"
                     data_aut.Recordset("socio") = Data1.Recordset("cl_codigo")
                     data_aut.Recordset("modulo") = "DESPACHO"
                     data_aut.Recordset("usuario_caja") = WElusuario
                     data_aut.Recordset.Update
                  Else
                     Xhab = Data1.Recordset("cl_codigo")
                     frm_autoriza.Show vbModal
                      '14063
                      '117670
                      '5112
                      Xelcodigoaut = InputBox("SOCIO CON CRÉDITOS PENDIENTES O CUOTAS, INGRESE CODIGO DE AUTORIZACIÓN SI ES CLAVE 3", "SOCIO CON CRÉDITOS PENDIENTES", Wopszond)
                      If Trim(Xelcodigoaut) <> "" Then
                         data_aut.RecordSource = "select * from Codigos_aut where codaut ='" & Trim(Xelcodigoaut) & "' and socio =" & Data1.Recordset("cl_codigo")
                         data_aut.Refresh
                         If data_aut.Recordset.RecordCount > 0 Then
                            Xq = 0
                            Xdeudasi = 0
                         Else
                            MsgBox "ATENCION! No se encuentra código de autorización, realice nuevamente la autorización o comunique a Administración", vbCritical
                            Xq = 9
                            Xdeudasi = 9
                         End If
                      Else
                         MsgBox "Socio con créditos o cuotas(>=3) pendientes, NO SE PODRÁ GRABAR LLAMADO CLAVE 3. ", vbCritical
                         Xq = 9
                         Xdeudasi = 9
                      End If
                  End If
               Else
                  Xdeudasi = 0
               End If
               If XAlta = 599 Then
                  XAlta = Xaltaanterior
               End If
               If IsNull(Data1.Recordset("saldo_chc2")) = False Then
                  If Data1.Recordset("saldo_chc2") = 1 Then
                     Xq = 11
                  End If
                  If Xq = 11 Then
                     MsgBox "ATENCION!! Socio con servicios RESTRINGIDOS! Estimado Funcionario NO dar servicio." & chr(13) _
                     & "El hacerlo estará bajo su exclusiva responsabilidad." & chr(13) & "El sistema no permitirá la continuidad de dicho servicio.", vbCritical, "SOCIOS"
                     MsgBox "SI ES UN LLAMADO CLAVE 3, DEBERA SOLICITAR AUTORIZACION al 097215419 PARA PODER GRABAR DATOS", vbInformation, "LLAMADO"
                     Xdeudasi = 9
                  End If
               End If
             End If
        End If
    End If
          
data_parsec.DatabaseName = App.path & "\mensa.mdb"
data_parsec.RecordSource = "mensaje"
data_parsec.Refresh

Wopspro = 99
If frm_largador.Check1.Value <> 1 Then
   If (Xcodzoning = 400 Or Xcodzoning = 401 Or Xcodzoning = 402 Or Xcodzoning = 403 Or Xcodzoning = 670 Or Xcodzoning = 671) And _
      (Data1.Recordset("cl_codconv") = "CCNOS" Or Data1.Recordset("cl_codconv") = "CCNSAM") Then
      Wopscob = 0
   Else
        If Val(frm_largador.cbozona.Text) = 1 Or Val(frm_largador.cbozona.Text) = 2 Or Val(frm_largador.cbozona.Text) = 3 Or Val(frm_largador.cbozona.Text) = 5 Or Val(frm_largador.cbozona.Text) = 6 Then
           If frm_largador.txt_cat.Text <> "" Then
              If frm_largador.txt_cat.Text = "SMIN" Or frm_largador.txt_cat.Text = "SMINA" Or frm_largador.txt_cat.Text = "UNIVS" Or _
                 frm_largador.txt_cat.Text = "UNNSAM" Or frm_largador.txt_cat.Text = "HEVANO" Or frm_largador.txt_cat.Text = "EVNSAM" Or _
                 frm_largador.txt_cat.Text = "CCNOS" Or frm_largador.txt_cat.Text = "CCNSAM" Or frm_largador.txt_cat.Text = "GANOS" Or _
                 frm_largador.txt_cat.Text = "CASANO" Or frm_largador.txt_cat.Text = "CASNSA" Then
                 ConectarBD
                 ConbdSapp.Open
                 Xsqlstr = "Select * from linmmdd where cod_cli =" & Data1.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806)"
                 With Xrecconve
                     .CursorLocation = adUseClient
                     .CursorType = adOpenKeyset
                     .LockType = adLockOptimistic
                     .Open Xsqlstr, ConbdSapp, , , adCmdText
                 End With
                 If Xrecconve.RecordCount > 0 Then
                    ConbdSapp.Close
                    Wopspro = 0
                    Wopscob = 0
                     data_parsec.DatabaseName = App.path & "\mensa.mdb"
                     data_parsec.RecordSource = "mensaje"
                     data_parsec.Refresh
                     data_parsec.Recordset.Edit
                     data_parsec.Recordset("text") = "RECUERDE! Confirmar socio con la mutualista."
                     data_parsec.Recordset.Update
                     XwYalomostro = 99
                     frm_mensajesvar.Show vbModal
                 Else
                     ConbdSapp.Close
                     If IsNull(Data1.Recordset("cl_decuota")) = False Then
                        If Data1.Recordset("cl_decuota") = 0 Or _
                           Data1.Recordset("cl_decuota") = 1 Or _
                           Data1.Recordset("cl_decuota") = 3 Or _
                           Data1.Recordset("cl_decuota") = 4 Then
                           data_parsec.DatabaseName = App.path & "\mensa.mdb"
                           data_parsec.RecordSource = "mensaje"
                           data_parsec.Refresh
                           data_parsec.Recordset.Edit
                           If frm_largador.txt_cat.Text = "SMIN" Or frm_largador.txt_cat.Text = "SMINA" Then
                              If Val(frm_largador.cbozona.Text) = 2 Or Val(frm_largador.cbozona.Text) = 3 Then
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                 & "Fotocopia de Cédula de identidad vigente. " & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                 chr(13) & "para realizar la misma."
                              Else
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Requerimientos:" & chr(13) _
                                 & "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                                 & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, por lo menos 6 meses de consumo," _
                                 & " a nombre del cliente que sea del mes corriente o anterior.)" & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                 chr(13) & "para realizar la misma."
                              End If
                           Else
                              data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                              & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                              chr(13) & "para realizar la misma."
                           End If
                           data_parsec.Recordset.Update
                           data_parsec.Refresh
                           XwYalomostro = 99
                           frm_mensajesvar.Show vbModal
                        Else
                           ConectarBD
                           ConbdSapp.Open
                           Xsqlstr = "Select * from linmmdd where cod_cli =" & Data1.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806)"
                           With Xrecconve
                               .CursorLocation = adUseClient
                               .CursorType = adOpenKeyset
                               .LockType = adLockOptimistic
                               .Open Xsqlstr, ConbdSapp, , , adCmdText
                           End With
                           If Xrecconve.RecordCount > 0 Then
                           Else
                              data_parsec.Recordset.Edit
                              If frm_largador.txt_cat.Text = "SMIN" Or frm_largador.txt_cat.Text = "SMINA" Then
                                 If Val(frm_largador.cbozona.Text) = 2 Or Val(frm_largador.cbozona.Text) = 3 Then
                                    data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                    & "Fotocopia de Cédula de identidad vigente. " & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                    chr(13) & "para realizar la misma."
                                 Else
                                    data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Requerimientos:" & chr(13) _
                                    & "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                                    & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, por lo menos 6 meses de consumo," _
                                    & " a nombre del cliente que sea del mes corriente o anterior.)" & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                    chr(13) & "para realizar la misma."
                                 End If
                              Else
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                 & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                 chr(13) & "para realizar la misma."
                              End If
                              data_parsec.Recordset.Update
                              data_parsec.Refresh
                              XwYalomostro = 99
                              frm_mensajesvar.Show vbModal
                           End If
                           ConbdSapp.Close
                        End If
                     Else
                        ConectarBD
                        ConbdSapp.Open
                        Xsqlstr = "Select * from linmmdd where cod_cli =" & Data1.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806)"
                        With Xrecconve
                            .CursorLocation = adUseClient
                            .CursorType = adOpenKeyset
                            .LockType = adLockOptimistic
                            .Open Xsqlstr, ConbdSapp, , , adCmdText
                        End With
                        If Xrecconve.RecordCount > 0 Then
                        Else
                           data_parsec.Recordset.Edit
                           If frm_largador.txt_cat.Text = "SMIN" Or frm_largador.txt_cat.Text = "SMINA" Then
                              If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                 & "Fotocopia de Cédula de identidad vigente. " & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                 chr(13) & "para realizar la misma."
                              Else
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Requerimientos:" & chr(13) _
                                 & "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                                 & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, por lo menos 6 meses de consumo," _
                                 & " a nombre del cliente que sea del mes corriente o anterior.)" & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                 chr(13) & "para realizar la misma."
                              End If
                           Else
                              data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                              & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                              chr(13) & "para realizar la misma."
                           End If
                           data_parsec.Recordset.Update
                           data_parsec.Refresh
                           XwYalomostro = 99
                           frm_mensajesvar.Show vbModal
                        End If
                        ConbdSapp.Close
                     End If
                 End If
              End If
           End If
        End If
    End If

End If
Wopspro = 0
          
If frm_largador.txt_ced.Text <> "" Then
   frm_largador.data_histant.RecordSource = "Select * from ante where ced =" & frm_largador.txt_ced.Text
   frm_largador.data_histant.Refresh
   If frm_largador.data_histant.Recordset.RecordCount > 0 Then
      frm_largador.txt_ante.Text = frm_largador.data_histant.Recordset("ante")
   Else
      frm_largador.txt_ante.Text = ""
   End If
End If


frm_buslla.Hide

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Deactivate()
'frm_busca.Hide

End Sub

Private Sub Form_Initialize()
'Data1.Recordset.MoveLast

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_aut.Connect = "odbc;dsn=" & Xconexrmt & ";"

'Data1.RecordSource = "Select top 500, * from clientes"
'Data1.Refresh
data_deuda.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_convbus.ConnectionString = "dsn=" & Xconexrmt

cbobus.ListIndex = 0

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_busca.Hide

End Sub


Private Sub txt_buscacli_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))
If KeyAscii = 13 Then
   If cbobus.ListIndex = 0 Then
      If txt_buscacli.Text <> "" Then
         Data1.RecordSource = "select top 80, * from clientes where cl_apellid >='" & txt_buscacli.Text & "' order by cl_apellid"
         Data1.Refresh
      Else
         MsgBox "No ingresó datos"
      End If
   Else
      If cbobus.ListIndex = 1 Then
         If txt_buscacli.Text <> "" Then
            Data1.RecordSource = "select * from clientes where cl_codigo =" & Val(txt_buscacli.Text) & " order by cl_codigo"
            Data1.Refresh
         Else
            MsgBox "No ingresó datos"
         End If
      Else
         If cbobus.ListIndex = 2 Then
            If txt_buscacli.Text <> "" Then
               Data1.RecordSource = "select * from clientes where cl_cedula =" & Val(txt_buscacli.Text) & " order by cl_cedula"
               Data1.Refresh
               If Data1.Recordset.RecordCount > 0 Then
               Else
                  MsgBox "No se encuentra CEDULA, busque por APELLIDOS"
                  txt_buscacli.SetFocus
               End If
            Else
               MsgBox "No ingresó datos"
            End If
         Else
            If cbobus.ListIndex = 3 Then
               If txt_buscacli.Text <> "" Then
                  Data1.RecordSource = "select top 70, * from clientes where cl_telefon >='" & txt_buscacli.Text & "' order by cl_telefon"
                  Data1.Refresh
               Else
                  MsgBox "No ingresó datos"
               End If
            Else
               If cbobus.ListIndex = 4 Then
                  If txt_buscacli.Text <> "" Then
                     Data1.RecordSource = "select top 70, * from clientes where cl_dpto >='" & txt_buscacli.Text & "' order by cl_dpto"
                     Data1.Refresh
                  Else
                     MsgBox "No ingresó datos"
                  End If
               Else
                  MsgBox "No seleccionó opción de búsqueda", vbInformation, "Mensaje"
                  cbobus.SetFocus
                  cbobus.ListIndex = 0
               End If
            End If
         End If
      End If
   End If
   DBGrid1.SetFocus
End If


End Sub
