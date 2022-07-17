VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pendlla 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendientes y En CURSO"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   17670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_pendlla.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   17670
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "frm_pendlla.frx":0442
      Height          =   5295
      Left            =   120
      OleObjectBlob   =   "frm_pendlla.frx":0456
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   16695
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   14280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frm_pendlla.frx":2399
      Height          =   4935
      Left            =   240
      OleObjectBlob   =   "frm_pendlla.frx":23AD
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   17055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc data_entrantes 
      Height          =   495
      Left            =   3720
      Top             =   7800
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_entrantes"
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
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   13800
      Top             =   7680
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_pendlla.frx":3FB4
      Height          =   6495
      Left            =   120
      OleObjectBlob   =   "frm_pendlla.frx":3FC8
      TabIndex        =   1
      Top             =   960
      Width           =   17175
   End
   Begin MSComctlLib.TabStrip tabver 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   17535
      _ExtentX        =   30930
      _ExtentY        =   12938
      MultiRow        =   -1  'True
      Style           =   2
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PENDIENTES"
            Key             =   "a"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "EN CURSO"
            Key             =   "b"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Traslados pendientes"
            Key             =   "c"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CMT"
            Key             =   "d"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seg. COVID19"
            Key             =   "e"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Covid19 HOY"
            Key             =   "f"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "COVID-19 +"
            Key             =   "g"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Resultados"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labcodus 
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labtoregi 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   15840
      TabIndex        =   16
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Registros:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   13920
      TabIndex        =   15
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FFFF&
      Caption         =   "RECLASIFICACIÓN"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   9600
      Width           =   10935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MOTIVO:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   9600
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   17640
      Y1              =   10200
      Y2              =   10200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   17640
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   9240
      Width           =   5895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Convenio:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   9240
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TELEFONO:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   9240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   8520
      Width           =   10935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DIRECCION:"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Más datos del llamado seleccionado:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   8160
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "F1 = Cambia de Pendientes a En Curso"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   11160
      Picture         =   "frm_pendlla.frx":5D5F
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   2655
   End
End
Attribute VB_Name = "frm_pendlla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tab_Click()


End Sub


Private Sub Command1_Click()

End Sub

Private Sub DBGrid1_DblClick()
Dim Voyallamar As String
Dim XfecCovid As Date
XfecCovid = Date - 5
Voyallamar = vbNo

On Error GoTo Quepasoalabrir
    If tabver.SelectedItem.index = 4 Or tabver.SelectedItem.index = 7 Then
       Voyallamar = MsgBox("Va a realizar la llamada a este paciente?", vbInformation + vbYesNo, "Despacho")
       If Voyallamar = vbYes Then
          If IsNull(Data1.Recordset("cmt_enproceso")) = False Then
             If Data1.Recordset("cmt_enproceso") = 1 Then
                MsgBox "ATENCION:! A este paciente lo está llamando el usuario: " & Data1.Recordset("cmt_usproc")
             End If
          End If
       End If
    End If
    Data1.RecordSource = "Select * from llamado where nrolla =" & DBGrid1.Columns(13)
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
       If Voyallamar = vbYes Then
          If IsNull(Data1.Recordset("cmt_enproceso")) = False Then
             If Data1.Recordset("cmt_enproceso") <> 1 Then
                Data1.Recordset.Edit
                Data1.Recordset("cmt_enproceso") = 1
                Data1.Recordset("cmt_usproc") = WElusuario
                Data1.Recordset.Update
             End If
          Else
             Data1.Recordset.Edit
             Data1.Recordset("cmt_enproceso") = 1
             Data1.Recordset("cmt_usproc") = WElusuario
             Data1.Recordset.Update
          End If
        End If
        frm_largador.txt_nro.Text = Data1.Recordset("nrolla")
        frm_largador.mfecha.Text = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
        frm_largador.txt_hora.Text = Format(Data1.Recordset("hora"), "HH:mm")
        frm_largador.txt_usua.Text = Data1.Recordset("usuario")
        If IsNull(Data1.Recordset("hora_anterior")) = False Then
           frm_largador.labanthor.Caption = Data1.Recordset("hora_anterior")
        Else
           frm_largador.labanthor.Caption = ""
        End If
        If IsNull(Data1.Recordset("matric")) = False Then
           frm_largador.txt_mat.Text = Data1.Recordset("matric")
        Else
           frm_largador.txt_mat.Text = 0
        End If
        If IsNull(Data1.Recordset("nomodif")) = False Then
           frm_largador.Check1.Value = Data1.Recordset("nomodif")
        Else
           frm_largador.Check1.Value = 0
        End If
        If IsNull(Data1.Recordset("segui_covid")) = False Then
           frm_largador.chcovid.Value = Data1.Recordset("segui_covid")
        Else
           frm_largador.chcovid.Value = 0
        End If
        If IsNull(Data1.Recordset("timbre")) = False Then
           frm_largador.cbotimbre.ListIndex = Data1.Recordset("timbre")
        Else
           frm_largador.cbotimbre.ListIndex = -1
        End If
        If IsNull(Data1.Recordset("valor_timbre")) = False Then
           frm_largador.t_timbre.Text = Data1.Recordset("valor_timbre")
        Else
           frm_largador.t_timbre.Text = ""
        End If
        If IsNull(Data1.Recordset("nombre")) = False Then
           frm_largador.txt_nomb.Text = Data1.Recordset("nombre")
        Else
           frm_largador.txt_nomb.Text = ""
        End If
        If IsNull(Data1.Recordset("edad")) = False Then
           frm_largador.txt_edad.Text = Data1.Recordset("edad")
        Else
           frm_largador.txt_edad.Text = ""
        End If
        If IsNull(Data1.Recordset("mes")) = False Then
           If Data1.Recordset("mes") > 10 Then
              frm_largador.txt_costo.Text = Data1.Recordset("mes")
           Else
              frm_largador.txt_costo.Text = 0
           End If
        Else
           frm_largador.txt_costo.Text = 0
        End If
        If Format(Data1.Recordset("fecha"), "yyyy/mm/dd") >= Format("2016/12/01", "yyyy/mm/dd") Then
           If IsNull(Data1.Recordset("realiza")) = False Then
              frm_largador.chtmut.Value = Data1.Recordset("realiza")
           Else
              frm_largador.chtmut.Value = 0
           End If
        Else
           frm_largador.chtmut.Value = 0
        End If
        If IsNull(Data1.Recordset("ano")) = False Then
           If Data1.Recordset("ano") > 10 Then
              frm_largador.txt_boleta.Text = Data1.Recordset("ano")
           Else
              frm_largador.txt_boleta.Text = 0
           End If
        Else
           frm_largador.txt_boleta.Text = 0
        End If
        If Data1.Recordset("pend") = 4 Then
           frm_largador.Command5.Visible = True
           frm_largador.Frame2.Visible = False
        Else
           frm_largador.Command5.Visible = False
           frm_largador.Frame2.Visible = True
        End If
        If IsNull(Data1.Recordset("unied")) = False Then
           If Data1.Recordset("unied") = 3 Then
              frm_largador.cboed.ListIndex = 0
           Else
              If Data1.Recordset("unied") = 2 Then
                 frm_largador.cboed.ListIndex = 1
              Else
                 If Data1.Recordset("unied") = 1 Then
                    frm_largador.cboed.ListIndex = 2
                 Else
                    frm_largador.cboed.ListIndex = 0
                 End If
              End If
           End If
        Else
           frm_largador.cboed.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("categ")) = False Then
           frm_largador.txt_cat.Text = Data1.Recordset("categ")
        Else
           frm_largador.txt_cat.Text = ""
        End If
        If IsNull(Data1.Recordset("nomcat")) = False Then
           frm_largador.txt_nomcat.Text = Data1.Recordset("nomcat")
        Else
           frm_largador.txt_nomcat.Text = ""
        End If
        If IsNull(Data1.Recordset("aft")) = False Then
           frm_largador.Label40.Caption = "AFT:" & Data1.Recordset("aft")
        Else
           frm_largador.Label40.Caption = ""
        End If
        If IsNull(Data1.Recordset("ci")) = False Then
           frm_largador.txt_ced.Text = Int(Data1.Recordset("ci"))
        Else
           frm_largador.txt_ced.Text = 0
        End If
        If IsNull(Data1.Recordset("telef")) = False Then
           frm_largador.txt_tel.Text = Data1.Recordset("telef")
        Else
           frm_largador.txt_tel.Text = ""
        End If
        If IsNull(Data1.Recordset("codzon")) = False Then
           If Data1.Recordset("codzon") = 2 Then
              frm_largador.cbozona.ListIndex = 1
           Else
              If Data1.Recordset("codzon") = 3 Then
                 frm_largador.cbozona.ListIndex = 2
              Else
                 If Data1.Recordset("codzon") = 4 Then
                    frm_largador.cbozona.ListIndex = 3
                 Else
                    If Data1.Recordset("codzon") = 5 Then
                       frm_largador.cbozona.ListIndex = 4
                    Else
                       If Data1.Recordset("codzon") = 6 Then
                          frm_largador.cbozona.ListIndex = 5
                       Else
                          If Data1.Recordset("codzon") = 7 Then
                             frm_largador.cbozona.ListIndex = 6
                          Else
                             frm_largador.cbozona.ListIndex = 0
                          End If
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbozona.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("base")) = False Then
           frm_largador.cbobase.Text = Data1.Recordset("base")
        Else
           frm_largador.cbobase.Text = 0
        End If
        If IsNull(Data1.Recordset("referen")) = False Then
           frm_largador.txt_direc.Text = Data1.Recordset("referen")
        Else
           frm_largador.txt_direc.Text = ""
        End If
        If IsNull(Data1.Recordset("motcon")) = False Then
           frm_largador.txt_ante.Text = Data1.Recordset("motcon")
        Else
           frm_largador.txt_ante.Text = ""
        End If
        If IsNull(Data1.Recordset("obs")) = False Then
           frm_largador.txt_obs.Text = Data1.Recordset("obs")
        Else
           frm_largador.txt_obs.Text = ""
        End If
        If IsNull(Data1.Recordset("obsmot")) = False Then
           frm_largador.txt_mot.Text = Data1.Recordset("obsmot")
        Else
           frm_largador.txt_mot.Text = ""
        End If
        If IsNull(Data1.Recordset("codmot")) = False Then
           If Data1.Recordset("codmot") = "R" Then
              frm_largador.cbocolor.ListIndex = 2
           Else
              If Data1.Recordset("codmot") = "A" Then
                 frm_largador.cbocolor.ListIndex = 1
              Else
                 If Data1.Recordset("codmot") = "C" Then
                    frm_largador.cbocolor.ListIndex = 3
                 Else
                    If Data1.Recordset("codmot") = "Z" Then
                       frm_largador.cbocolor.ListIndex = 4
                    Else
                       If Data1.Recordset("codmot") = "N" Then
                          frm_largador.cbocolor.ListIndex = 5
                       Else
                          frm_largador.cbocolor.ListIndex = 0
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolor.ListIndex = 0
        End If
        If frm_largador.cbocolor.Text = "VERDE" Then
           frm_largador.cbocolor.BackColor = &HC000&
        Else
           If frm_largador.cbocolor.Text = "ROJO" Then
              frm_largador.cbocolor.BackColor = &HFF&
           Else
              If frm_largador.cbocolor.Text = "AMARILLO" Then
                 frm_largador.cbocolor.BackColor = &HFFFF&
              Else
                 If frm_largador.cbocolor.Text = "CELESTE" Then
                    frm_largador.cbocolor.BackColor = &HFFFF00
                 Else
                    If frm_largador.cbocolor.Text = "AZUL" Then
                       frm_largador.cbocolor.BackColor = &HC00000
                    Else
                       If frm_largador.cbocolor.Text = "NEGRO" Then
                          frm_largador.cbocolor.BackColor = &H80000006
                       Else
                          frm_largador.cbocolor.BackColor = &HFFFFFF
                       End If
                    End If
                 End If
              End If
           End If
        End If
        If IsNull(Data1.Recordset("movilpas")) = False Then
           frm_largador.txt_movil.Text = Data1.Recordset("movilpas")
        Else
           frm_largador.txt_movil.Text = ""
        End If
        If IsNull(Data1.Recordset("fecpas")) = False Then
           frm_largador.mfecasig.Text = Format(Data1.Recordset("fecpas"), "dd/mm/yyyy")
        Else
           frm_largador.mfecasig.Text = "__/__/____"
        End If
        If IsNull(Data1.Recordset("horpas")) = False Then
           frm_largador.txt_horasig.Text = Format(Data1.Recordset("horpas"), "HH:mm")
        Else
           frm_largador.txt_horasig.Text = ""
        End If
        If IsNull(Data1.Recordset("fecsali")) = False Then
           frm_largador.msalida.Text = Format(Data1.Recordset("fecsali"), "dd/mm/yyyy")
        Else
           frm_largador.msalida.Text = "__/__/____"
        End If
        If IsNull(Data1.Recordset("horsali")) = False Then
           frm_largador.txt_horsal.Text = Format(Data1.Recordset("horsali"), "HH:mm")
        Else
           frm_largador.txt_horsal.Text = ""
        End If
        If IsNull(Data1.Recordset("fec_llega")) = False Then
           frm_largador.mllegada.Text = Format(Data1.Recordset("fec_llega"), "dd/mm/yyyy")
        Else
           frm_largador.mllegada.Text = "__/__/____"
        End If
        If IsNull(Data1.Recordset("hor_llega")) = False Then
           frm_largador.txt_horlle.Text = Format(Data1.Recordset("hor_llega"), "HH:mm")
        Else
           frm_largador.txt_horlle.Text = ""
        End If
        If IsNull(Data1.Recordset("fec_rea")) = False Then
           frm_largador.mtd.Text = Format(Data1.Recordset("fec_rea"), "dd/mm/yyyy")
        Else
           frm_largador.mtd.Text = "__/__/____"
        End If
        If IsNull(Data1.Recordset("hor_rea")) = False Then
           If Data1.Recordset("hor_rea") <> "" Then
              frm_largador.txt_hortd.Text = Format(Data1.Recordset("hor_rea"), "HH:mm")
           Else
              frm_largador.txt_hortd.Text = "__:__"
           End If
        Else
           frm_largador.txt_hortd.Text = "__:__"
        End If
        If IsNull(Data1.Recordset("diag")) = False Then
           frm_largador.txt_diag.Text = Data1.Recordset("diag")
        Else
           frm_largador.txt_diag.Text = ""
        End If
        If IsNull(Data1.Recordset("colormot")) = False Then
           If Data1.Recordset("colormot") = "R" Then
              frm_largador.cbocolfin.ListIndex = 2
           Else
              If Data1.Recordset("colormot") = "A" Then
                 frm_largador.cbocolfin.ListIndex = 1
              Else
                 If Data1.Recordset("colormot") = "V" Then
                    frm_largador.cbocolfin.ListIndex = 0
                 Else
                    If Data1.Recordset("colormot") = "N" Then
                       frm_largador.cbocolfin.ListIndex = 3
                    Else
                       frm_largador.cbocolfin.Text = ""
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolfin.Text = ""
        End If
        If IsNull(Data1.Recordset("nommed")) = False Then
           frm_largador.dbcbomed.Text = Data1.Recordset("nommed")
        Else
           frm_largador.dbcbomed.ListField = ""
           frm_largador.dbcbomed.BoundColumn = ""
           frm_largador.dbcbomed.Text = ""
           frm_largador.dbcbomed.ListField = "med_nombre"
           frm_largador.dbcbomed.BoundColumn = "med_nombre"
        End If
        If IsNull(Data1.Recordset("codmed")) = False Then
           frm_largador.txt_codmed.Text = Data1.Recordset("codmed")
        Else
           frm_largador.txt_codmed.Text = 0
        End If
        If IsNull(Data1.Recordset("trasla")) = False Then
           If Data1.Recordset("trasla") > 0 Then
              frm_largador.cbotras.ListIndex = Data1.Recordset("trasla")
           Else
              frm_largador.cbotras.ListIndex = 0
           End If
        Else
           frm_largador.cbotras.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("lugar")) = False Then
           frm_largador.txt_lugar.Text = Data1.Recordset("lugar")
        Else
           frm_largador.txt_lugar.Text = ""
        End If
        If IsNull(Data1.Recordset("hsald")) = False Then
           frm_largador.txt_trassal.Text = Format(Data1.Recordset("hsald"), "HH:mm")
        Else
           frm_largador.txt_trassal.Text = ""
        End If
        If IsNull(Data1.Recordset("hllega")) = False Then
           frm_largador.txt_enca.Text = Format(Data1.Recordset("hllega"), "HH:mm")
        Else
           frm_largador.txt_enca.Text = ""
        End If
        If IsNull(Data1.Recordset("hzona")) = False Then
           frm_largador.txt_enzona.Text = Format(Data1.Recordset("hzona"), "HH:mm")
        Else
           frm_largador.txt_enzona.Text = ""
        End If
        If IsNull(Data1.Recordset("movtras")) = False Then
           frm_largador.txt_movtra.Text = Data1.Recordset("movtras")
        Else
           frm_largador.txt_movtra.Text = ""
        End If
        If IsNull(Data1.Recordset("dcobr")) = False Then
           frm_largador.Combo1.Text = Data1.Recordset("dcobr")
        Else
           frm_largador.Combo1.Text = ""
        End If
        If IsNull(Data1.Recordset("totdem")) = False Then
           frm_largador.txt_demora.Text = Format(Data1.Recordset("totdem"), "HH:mm")
        Else
           frm_largador.txt_demora.Text = ""
        End If
        If IsNull(Data1.Recordset("activo")) = False Then
           frm_largador.Label3.Caption = Format(Data1.Recordset("activo"), "HH:mm:ss")
        Else
           frm_largador.Label3.Caption = "00:00:00"
        End If
        If IsNull(Data1.Recordset("timdes")) = False Then
           frm_largador.Label39.Caption = Data1.Recordset("timdes")
        Else
           frm_largador.Label39.Caption = "Sin Largar"
        End If
        If IsNull(Data1.Recordset("motmov")) = True Then
           frm_largador.txt_locali.Text = ""
        Else
           frm_largador.txt_locali.Text = Data1.Recordset("motmov")
        End If
        If IsNull(Data1.Recordset("mm")) = True Then
           frm_largador.Label41.Caption = 0
        Else
           frm_largador.Label41.Caption = Data1.Recordset("mm")
        End If
        If IsNull(Data1.Recordset("thh")) = True Then
           frm_largador.Label42.Caption = 0
        Else
           frm_largador.Label42.Caption = Data1.Recordset("thh")
        End If
        If IsNull(Data1.Recordset("tmm")) = True Then
           frm_largador.Label43.Caption = 0
        Else
           frm_largador.Label43.Caption = Data1.Recordset("tmm")
        End If
        If IsNull(Data1.Recordset("pasado")) = True Then
           frm_largador.Label44.Caption = 0
        Else
           frm_largador.Label44.Caption = Data1.Recordset("pasado")
        End If
        If IsNull(Data1.Recordset("ano")) = True Then
           frm_largador.Label45.Caption = 0
        Else
           frm_largador.Label45.Caption = Data1.Recordset("ano")
        End If
        If IsNull(Data1.Recordset("mes")) = True Then
           frm_largador.Label46.Caption = -1
        Else
           frm_largador.Label46.Caption = Data1.Recordset("mes")
        End If
        If IsNull(Data1.Recordset("timsi")) = True Then
           frm_largador.Label48.Caption = 0
        Else
           frm_largador.Label48.Caption = Data1.Recordset("timsi")
        End If
        If IsNull(Data1.Recordset("enfer")) = True Then
           frm_largador.Check2.Value = 0
        Else
           frm_largador.Check2.Value = Data1.Recordset("enfer")
        End If
        If IsNull(Data1.Recordset("motcance")) = True Then
           frm_largador.txt_quien.Text = ""
        Else
           frm_largador.txt_quien.Text = Data1.Recordset("motcance")
        End If
        If IsNull(Data1.Recordset("hh")) = True Then
           frm_largador.Combo3.ListIndex = -1
        Else
           frm_largador.Combo3.ListIndex = Data1.Recordset("hh")
        End If
        If IsNull(Data1.Recordset("cancela")) = True Then
           If IsNull(Data1.Recordset("hor_cance")) = False Then
              frm_largador.txt_salca.Text = Data1.Recordset("hor_cance")
           Else
              frm_largador.txt_salca.Text = ""
           End If
        End If
        Data3.RecordSource = "Select * from resplla where nro =" & Data1.Recordset("nrolla")
        Data3.Refresh
        If Data3.Recordset.RecordCount > 0 Then
           If IsNull(Data3.Recordset("telef")) = False Then
              If Data3.Recordset("telef") = "RECIBO" Then
                 frm_largador.Combo2.ListIndex = 0
              Else
                 If Data3.Recordset("telef") = "CONFORME" Then
                    frm_largador.Combo2.ListIndex = 1
                 Else
                    frm_largador.Combo2.ListIndex = -1
                 End If
              End If
           Else
              frm_largador.Combo2.ListIndex = -1
           End If
           
           If IsNull(Data3.Recordset("movil_rea")) = False Then
              If Data3.Recordset("movil_rea") > 0 Then
                 frm_largador.data_chof.RecordSource = "Select * from movil where nromov =" & Data3.Recordset("movil_rea")
                 frm_largador.data_chof.Refresh
                 If frm_largador.data_chof.Recordset.RecordCount > 0 Then
                    frm_largador.labcodchof.Caption = Data3.Recordset("movil_rea")
                    frm_largador.labnomchof.Caption = "Chof.:" & frm_largador.data_chof.Recordset("chofer")
                 Else
                    frm_largador.labcodchof.Caption = 0
                    frm_largador.labnomchof.Caption = ""
                 End If
              Else
                 frm_largador.labcodchof.Caption = 0
                 frm_largador.labnomchof.Caption = ""
              End If
           Else
              frm_largador.labcodchof.Caption = 0
              frm_largador.labnomchof.Caption = ""
           End If
           If IsNull(Data3.Recordset("mes")) = False Then
              frm_largador.t_codced.Text = Int(Data3.Recordset("mes"))
           Else
              frm_largador.t_codced.Text = 0
           End If
           If IsNull(Data3.Recordset("pasado")) = False Then
              frm_largador.Check4.Value = Data3.Recordset("pasado")
           Else
              frm_largador.Check4.Value = 0
           End If
           If IsNull(Data3.Recordset("hzona")) = False Then
              frm_largador.labcmt.Visible = True
              frm_largador.labcmt.Caption = "PASADO A CMT HORA:" & Format(Data3.Recordset("hzona"), "HH:mm")
              If IsNull(Data3.Recordset("mm")) = False Then
                 If Data3.Recordset("mm") = 1 Then
                    frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " NO RESUELTO H."
                    If IsNull(Data3.Recordset("hsald")) = False Then
                       frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & Data3.Recordset("hsald")
                    End If
                    If IsNull(Data3.Recordset("totend")) = False Then
                       If Data3.Recordset("totend") = "R" Then
                          frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RECLASIFICA A ROJO"
                       End If
                       If Data3.Recordset("totend") = "A" Then
                          frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RECLASIFICA A AMARILLO"
                       End If
                    End If
                 End If
                 If Data3.Recordset("mm") = 2 Then
                    frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RESUELTO HORA:"
                    If IsNull(Data3.Recordset("hor_rea")) = False Then
                       frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & Data3.Recordset("hor_rea")
                    End If
                 End If
                 If Data3.Recordset("mm") = 2 Or Data3.Recordset("mm") = 3 Then
                    frm_largador.Command5.Visible = True
                    frm_largador.Frame2.Visible = False
                 Else
                    If Data3.Recordset("mm") = 1 Then
                       frm_largador.Command5.Visible = False
                       If WDespa = 1 Then
                          frm_largador.Frame2.Visible = False
                       Else
                          frm_largador.Frame2.Visible = True
                       End If
                    Else
                       frm_largador.Command5.Visible = True
                       frm_largador.Frame2.Visible = False
                    End If
                 End If
              Else
                 frm_largador.Command5.Visible = True
                 frm_largador.Frame2.Visible = False
              End If
           Else
              frm_largador.labcmt.Caption = ""
              frm_largador.labcmt.Visible = False
              frm_largador.Command5.Visible = False
              If WDespa = 1 Then
                 frm_largador.Frame2.Visible = False
              Else
                 frm_largador.Frame2.Visible = True
              End If
           End If
           If IsNull(Data3.Recordset("movilpas")) = False Then
              Data4.Recordset.FindFirst "med_cod =" & Data3.Recordset("movilpas")
              If Not Data4.Recordset.NoMatch Then
                 frm_largador.dbcbomed2.Text = Data4.Recordset("med_nombre")
              Else
                 frm_largador.dbcbomed2.ListField = ""
                 frm_largador.dbcbomed2.BoundColumn = ""
                 frm_largador.dbcbomed2.Text = ""
                 frm_largador.dbcbomed2.ListField = "med_nombre"
                 frm_largador.dbcbomed2.BoundColumn = "med_nombre"
              End If
              frm_largador.txt_codmed2.Text = Data3.Recordset("movilpas")
           Else
              frm_largador.txt_codmed2.Text = 0
              frm_largador.dbcbomed2.ListField = ""
              frm_largador.dbcbomed2.BoundColumn = ""
              frm_largador.dbcbomed2.Text = ""
              frm_largador.dbcbomed2.ListField = "med_nombre"
             frm_largador.dbcbomed2.BoundColumn = "med_nombre"
           End If
           If IsNull(Data3.Recordset("fec_llega")) = False Then
              frm_largador.mftrassol.Text = Format(Data3.Recordset("fec_llega"), "dd/mm/yyyy")
           Else
              frm_largador.mftrassol.Text = "__/__/____"
           End If
           If IsNull(Data3.Recordset("hor_llega")) = False Then
              frm_largador.mhtrassol.Text = Format(Data3.Recordset("hor_llega"), "HH:mm")
           Else
              frm_largador.mhtrassol.Text = "__:__"
           End If
        Else
           frm_largador.txt_codmed2.Text = 0
           frm_largador.Check4.Value = 0
           frm_largador.dbcbomed2.Text = ""
           frm_largador.mftrassol.Text = "__/__/____"
           frm_largador.mhtrassol.Text = "__:__"
           frm_largador.t_codced.Text = 0
           frm_largador.labcmt.Caption = ""
           frm_largador.labcmt.Visible = False
        End If
        Data3.Recordset.Close
        Data1.Recordset.Close
        Unload Me
    Else
        MsgBox "No se encontró el registro, VERIFIQUE!!", vbCritical
    End If
    

Exit Sub

Quepasoalabrir:
            If Err.Number > 0 Then
               MsgBox " ERROR: " & str(Err.Number) & " Vuelva a intentar abrir o cierre el sistema"
               If tabver.SelectedItem.index = 1 Then
                  DBGrid1.Visible = True
                  DBGrid2.Visible = False
                  DBGrid3.Visible = False
                  DBGrid1.BackColor = &HFFC0C0
                  Data1.RecordSource = "Select * from llamado where pend <>" & 2 & " And pend <>" & 1 & " and pend <>" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by nrolla"
                  Data1.Refresh
                  labtoregi.Caption = Data1.Recordset.RecordCount
               Else
                  If tabver.SelectedItem.index = 2 Then
                     DBGrid1.Visible = True
                     DBGrid2.Visible = False
                     DBGrid3.Visible = False
                     DBGrid1.BackColor = &HC0C0FF
                     Data1.RecordSource = "Select * from llamado where pend =" & 1 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by nrolla"
                     Data1.Refresh
                     labtoregi.Caption = Data1.Recordset.RecordCount
                  Else
                     If tabver.SelectedItem.index = 3 Then
                        DBGrid1.BackColor = &HC0FFC0
                        DBGrid1.Visible = False
                        DBGrid2.Visible = False
                        DBGrid3.Visible = True
'                       Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/12/2011", "yyyy/mm/dd") & "# and trasla >=" & 1 & " and hzona ='" & Null & "' order by nrolla"
                        Data1.RecordSource = "Select * from llamado where trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14) and fecha >=#" & Format("08/08/2019", "yyyy/mm/dd") & "# and hzona ='" & Null & "' and codmot <>'" & "Z" & "' order by nrolla"
                        Data1.Refresh
                        labtoregi.Caption = Data1.Recordset.RecordCount
                     Else
                        If tabver.SelectedItem.index = 4 Then
                           DBGrid1.Visible = True
                           DBGrid2.Visible = False
                           DBGrid3.Visible = False
                           DBGrid1.BackColor = &H80FF&
                           If Trim(labcodus.Caption) <> "" Then
                              Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & Val(labcodus.Caption) & " order by fecha,hora"
                           Else
                              If ControlUsuario("Utilitarios despacho") = 1 Then
                                 Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
                              Else
                                 Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & 999999 & " order by fecha,hora"
                              End If
                           End If
                           Data1.Refresh
                           labtoregi.Caption = Data1.Recordset.RecordCount
                        Else
                           If tabver.SelectedItem.index = 5 Then
                              DBGrid1.Visible = True
                              DBGrid2.Visible = False
                              DBGrid3.Visible = False
                              DBGrid1.BackColor = &HC000&
'                              Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) and (isopa_result is null or isopa_result in ('Negativo')) order by fecha,hora"
                              Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and prox_control >=#" & Format(XfecCovid, "yyyy/mm/dd") & "# and prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# and (isopa_result is null or isopa_result in ('Negativo')) order by nombre,fecha,hora"
'                              Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null or prox_control) and prox_control >=#" & Format(XfecCovid, "yyyy/mm/dd") & "# and (isopa_result is null or isopa_result in ('Negativo')) order by fecha,hora"
                              Data1.Refresh
                              labtoregi.Caption = Data1.Recordset.RecordCount
                           Else
                              If tabver.SelectedItem.index = 6 Then
                                 DBGrid1.Visible = True
                                 DBGrid2.Visible = False
                                 DBGrid3.Visible = False
                                 DBGrid1.BackColor = &H80FF80
                                 Data1.RecordSource = "Select * from llamado where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null order by fecha,hora"
                                 Data1.Refresh
                                 labtoregi.Caption = Data1.Recordset.RecordCount
                              Else
                                 If tabver.SelectedItem.index = 7 Then
                                    DBGrid1.Visible = True
                                    DBGrid2.Visible = False
                                    DBGrid3.Visible = False
                                    DBGrid1.BackColor = &H8080FF
                                    Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
                                    Data1.Refresh
                                    labtoregi.Caption = Data1.Recordset.RecordCount
                                 Else
                                    If tabver.SelectedItem.index = 8 Then
                                       DBGrid1.Visible = False
                                       DBGrid2.Visible = True
                                       DBGrid3.Visible = False
                                       Data6.RecordSource = "select * from sol_hisopos where deriva in (1) and fecha_cierre is null order by fecha"
                                       Data6.Refresh
                                       labtoregi.Caption = Data6.Recordset.RecordCount
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               Unload Me
            End If
           
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Quepasaalver

If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick

End If

Exit Sub

Quepasaalver:
             If Err.Number > 0 Then
                MsgBox "Error al abrir, reintente o cierre ésta pantalla y vuelva a abrir", vbInformation
             End If
             
End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Data1.Recordset.RecordCount > 0 Then
    If IsNull(Data1.Recordset("direcc")) = False Then
       If IsNull(Data1.Recordset("referen")) = False Then
          Label4.Caption = Data1.Recordset("referen")
       Else
          Label4.Caption = Data1.Recordset("direcc")
       End If
       If IsNull(Data1.Recordset("telef")) = False Then
          Label6.Caption = Data1.Recordset("telef")
       Else
          Label6.Caption = ""
       End If
       If IsNull(Data1.Recordset("nomcat")) = False Then
          Label8.Caption = Data1.Recordset("nomcat")
       Else
          Label8.Caption = ""
       End If
       If IsNull(Data1.Recordset("motcon")) = False Then
          If IsNull(Data1.Recordset("obsmot")) = False Then
             Label10.Caption = Data1.Recordset("motcon") + " " + Data1.Recordset("obsmot")
          Else
             Label10.Caption = Data1.Recordset("motcon")
          End If
       Else
          Label10.Caption = ""
       End If
       
    End If
    
    If tabver.SelectedItem.index = 4 Then
       Data5.Connect = "odbc;dsn=" & Xconexrmt & ";"
        
       If IsNull(Data1.Recordset("codmedcmt")) = False Then
          If Data1.Recordset("codmedcmt") = 9999 Then
             Label10.Caption = Label10.Caption & " PASADO A: Atención al socio"
          Else
             Data5.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
             Data5.Refresh
             If Data5.Recordset.RecordCount > 0 Then
                Label10.Caption = Label10.Caption & " PASADO A:" & Data5.Recordset("nom_med")
             End If
          End If
       End If
    End If
    Data5.Connect = "odbc;dsn=" & Xconexrmt & ";"
    Data5.RecordSource = "Select * from resplla where nro =" & DBGrid1.Columns(13)
    Data5.Refresh
    If Data5.Recordset.RecordCount > 0 Then
       If IsNull(Data5.Recordset("hzona")) = False Then
          If IsNull(Data5.Recordset("mm")) = False Then
             If Data5.Recordset("mm") = 1 Then
                If IsNull(Data5.Recordset("totend")) = False Then
                   If Data5.Recordset("totend") = "R" Then
                      Label12.ForeColor = &HFF&
                      Label12.Caption = " ROJO"
                   End If
                   If Data5.Recordset("totend") = "A" Then
                      Label12.ForeColor = &H80FF&
                      Label12.Caption = Label12.Caption & "AMARILLO"
                   End If
                Else
                   Label12.ForeColor = &HFF&
                   Label12.Caption = ""
                End If
             Else
                Label12.ForeColor = &HFF&
                Label12.Caption = ""
             End If
          Else
             Label12.ForeColor = &HFF&
             Label12.Caption = ""
          End If
       Else
          Label12.ForeColor = &HFF&
          Label12.Caption = ""
       End If
    Else
       Label12.ForeColor = &HFF&
       Label12.Caption = ""
    End If

End If

End Sub


Private Sub DBGrid2_DblClick()
Dim DeseaCerrar As String
DeseaCerrar = MsgBox("Desea cerrar el CMT como realizado?", vbInformation + vbYesNo, "CMT Resultados")
If DeseaCerrar = vbYes Then
   frm_cmtresult.Show vbModal
   Unload Me

End If

End Sub

Private Sub DBGrid3_DblClick()
Dim XfecCovid As Date
XfecCovid = Date - 5

On Error GoTo Quepasoalabrir
    
        frm_largador.txt_nro.Text = Data1.Recordset("nrolla")
        frm_largador.mfecha.Text = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
        frm_largador.txt_hora.Text = Format(Data1.Recordset("hora"), "HH:mm")
        frm_largador.txt_usua.Text = Data1.Recordset("usuario")
        If IsNull(Data1.Recordset("hora_anterior")) = False Then
           frm_largador.labanthor.Caption = Data1.Recordset("hora_anterior")
        Else
           frm_largador.labanthor.Caption = ""
        End If
        If IsNull(Data1.Recordset("matric")) = False Then
           frm_largador.txt_mat.Text = Data1.Recordset("matric")
        Else
           frm_largador.txt_mat.Text = 0
        End If
        If IsNull(Data1.Recordset("nomodif")) = False Then
           frm_largador.Check1.Value = Data1.Recordset("nomodif")
        Else
           frm_largador.Check1.Value = 0
        End If
        If IsNull(Data1.Recordset("segui_covid")) = False Then
           frm_largador.chcovid.Value = Data1.Recordset("segui_covid")
        Else
           frm_largador.chcovid.Value = 0
        End If
        If IsNull(Data1.Recordset("timbre")) = False Then
           frm_largador.cbotimbre.ListIndex = Data1.Recordset("timbre")
        Else
           frm_largador.cbotimbre.ListIndex = -1
        End If
        If IsNull(Data1.Recordset("valor_timbre")) = False Then
           frm_largador.t_timbre.Text = Data1.Recordset("valor_timbre")
        Else
           frm_largador.t_timbre.Text = ""
        End If
        If IsNull(Data1.Recordset("nombre")) = False Then
           frm_largador.txt_nomb.Text = Data1.Recordset("nombre")
        Else
           frm_largador.txt_nomb.Text = ""
        End If
        If IsNull(Data1.Recordset("edad")) = False Then
           frm_largador.txt_edad.Text = Data1.Recordset("edad")
        Else
           frm_largador.txt_edad.Text = ""
        End If
        If IsNull(Data1.Recordset("mes")) = False Then
           If Data1.Recordset("mes") > 10 Then
              frm_largador.txt_costo.Text = Data1.Recordset("mes")
           Else
              frm_largador.txt_costo.Text = 0
           End If
        Else
           frm_largador.txt_costo.Text = 0
        End If
        If Format(Data1.Recordset("fecha"), "yyyy/mm/dd") >= Format("2016/12/01", "yyyy/mm/dd") Then
           If IsNull(Data1.Recordset("realiza")) = False Then
              frm_largador.chtmut.Value = Data1.Recordset("realiza")
           Else
              frm_largador.chtmut.Value = 0
           End If
        Else
           frm_largador.chtmut.Value = 0
        End If
        If IsNull(Data1.Recordset("ano")) = False Then
           If Data1.Recordset("ano") > 10 Then
              frm_largador.txt_boleta.Text = Data1.Recordset("ano")
           Else
              frm_largador.txt_boleta.Text = 0
           End If
        Else
           frm_largador.txt_boleta.Text = 0
        End If
        If Data1.Recordset("pend") = 4 Then
           frm_largador.Command5.Visible = True
           frm_largador.Frame2.Visible = False
        Else
           frm_largador.Command5.Visible = False
           frm_largador.Frame2.Visible = True
        End If
        If IsNull(Data1.Recordset("unied")) = False Then
           If Data1.Recordset("unied") = 3 Then
              frm_largador.cboed.ListIndex = 0
           Else
              If Data1.Recordset("unied") = 2 Then
                 frm_largador.cboed.ListIndex = 1
              Else
                 If Data1.Recordset("unied") = 1 Then
                    frm_largador.cboed.ListIndex = 2
                 Else
                    frm_largador.cboed.ListIndex = 0
                 End If
              End If
           End If
        Else
           frm_largador.cboed.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("categ")) = False Then
           frm_largador.txt_cat.Text = Data1.Recordset("categ")
        Else
           frm_largador.txt_cat.Text = ""
        End If
        If IsNull(Data1.Recordset("nomcat")) = False Then
           frm_largador.txt_nomcat.Text = Data1.Recordset("nomcat")
        Else
           frm_largador.txt_nomcat.Text = ""
        End If
        If IsNull(Data1.Recordset("aft")) = False Then
           frm_largador.Label40.Caption = "AFT:" & Data1.Recordset("aft")
        Else
           frm_largador.Label40.Caption = ""
        End If
        If IsNull(Data1.Recordset("ci")) = False Then
           frm_largador.txt_ced.Text = Int(Data1.Recordset("ci"))
        Else
           frm_largador.txt_ced.Text = 0
        End If
        If IsNull(Data1.Recordset("telef")) = False Then
           frm_largador.txt_tel.Text = Data1.Recordset("telef")
        Else
           frm_largador.txt_tel.Text = ""
        End If
        If IsNull(Data1.Recordset("codzon")) = False Then
           If Data1.Recordset("codzon") = 2 Then
              frm_largador.cbozona.ListIndex = 1
           Else
              If Data1.Recordset("codzon") = 3 Then
                 frm_largador.cbozona.ListIndex = 2
              Else
                 If Data1.Recordset("codzon") = 4 Then
                    frm_largador.cbozona.ListIndex = 3
                 Else
                    If Data1.Recordset("codzon") = 5 Then
                       frm_largador.cbozona.ListIndex = 4
                    Else
                       If Data1.Recordset("codzon") = 6 Then
                          frm_largador.cbozona.ListIndex = 5
                       Else
                          If Data1.Recordset("codzon") = 7 Then
                             frm_largador.cbozona.ListIndex = 6
                          Else
                             frm_largador.cbozona.ListIndex = 0
                          End If
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbozona.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("base")) = False Then
           frm_largador.cbobase.Text = Data1.Recordset("base")
        Else
           frm_largador.cbobase.Text = 0
        End If
        If IsNull(Data1.Recordset("referen")) = False Then
           frm_largador.txt_direc.Text = Data1.Recordset("referen")
        Else
           frm_largador.txt_direc.Text = ""
        End If
        If IsNull(Data1.Recordset("motcon")) = False Then
           frm_largador.txt_ante.Text = Data1.Recordset("motcon")
        Else
           frm_largador.txt_ante.Text = ""
        End If
        If IsNull(Data1.Recordset("obs")) = False Then
           frm_largador.txt_obs.Text = Data1.Recordset("obs")
        Else
           frm_largador.txt_obs.Text = ""
        End If
        If IsNull(Data1.Recordset("obsmot")) = False Then
           frm_largador.txt_mot.Text = Data1.Recordset("obsmot")
        Else
           frm_largador.txt_mot.Text = ""
        End If
        If IsNull(Data1.Recordset("codmot")) = False Then
           If Data1.Recordset("codmot") = "R" Then
              frm_largador.cbocolor.ListIndex = 2
           Else
              If Data1.Recordset("codmot") = "A" Then
                 frm_largador.cbocolor.ListIndex = 1
              Else
                 If Data1.Recordset("codmot") = "C" Then
                    frm_largador.cbocolor.ListIndex = 3
                 Else
                    If Data1.Recordset("codmot") = "Z" Then
                       frm_largador.cbocolor.ListIndex = 4
                    Else
                       If Data1.Recordset("codmot") = "N" Then
                          frm_largador.cbocolor.ListIndex = 5
                       Else
                          frm_largador.cbocolor.ListIndex = 0
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolor.ListIndex = 0
        End If
        If frm_largador.cbocolor.Text = "VERDE" Then
           frm_largador.cbocolor.BackColor = &HC000&
        Else
           If frm_largador.cbocolor.Text = "ROJO" Then
              frm_largador.cbocolor.BackColor = &HFF&
           Else
              If frm_largador.cbocolor.Text = "AMARILLO" Then
                 frm_largador.cbocolor.BackColor = &HFFFF&
              Else
                 If frm_largador.cbocolor.Text = "CELESTE" Then
                    frm_largador.cbocolor.BackColor = &HFFFF00
                 Else
                    If frm_largador.cbocolor.Text = "AZUL" Then
                       frm_largador.cbocolor.BackColor = &HC00000
                    Else
                       If frm_largador.cbocolor.Text = "NEGRO" Then
                          frm_largador.cbocolor.BackColor = &H80000006
                       Else
                          frm_largador.cbocolor.BackColor = &HFFFFFF
                       End If
                    End If
                 End If
              End If
           End If
        End If
        If IsNull(Data1.Recordset("movilpas")) = False Then
           frm_largador.txt_movil.Text = Data1.Recordset("movilpas")
        Else
           frm_largador.txt_movil.Text = ""
        End If
        If IsNull(Data1.Recordset("fecpas")) = False Then
           frm_largador.mfecasig.Text = Format(Data1.Recordset("fecpas"), "dd/mm/yyyy")
        Else
           frm_largador.mfecasig.Text = "__/__/____"
        End If
        If IsNull(Data1.Recordset("horpas")) = False Then
           frm_largador.txt_horasig.Text = Format(Data1.Recordset("horpas"), "HH:mm")
        Else
           frm_largador.txt_horasig.Text = ""
        End If
        If IsNull(Data1.Recordset("fecsali")) = False Then
           frm_largador.msalida.Text = Format(Data1.Recordset("fecsali"), "dd/mm/yyyy")
        Else
           frm_largador.msalida.Text = "__/__/____"
        End If
        If IsNull(Data1.Recordset("horsali")) = False Then
           frm_largador.txt_horsal.Text = Format(Data1.Recordset("horsali"), "HH:mm")
        Else
           frm_largador.txt_horsal.Text = ""
        End If
        If IsNull(Data1.Recordset("fec_llega")) = False Then
           frm_largador.mllegada.Text = Format(Data1.Recordset("fec_llega"), "dd/mm/yyyy")
        Else
           frm_largador.mllegada.Text = "__/__/____"
        End If
        If IsNull(Data1.Recordset("hor_llega")) = False Then
           frm_largador.txt_horlle.Text = Format(Data1.Recordset("hor_llega"), "HH:mm")
        Else
           frm_largador.txt_horlle.Text = ""
        End If
        If IsNull(Data1.Recordset("fec_rea")) = False Then
           frm_largador.mtd.Text = Format(Data1.Recordset("fec_rea"), "dd/mm/yyyy")
        Else
           frm_largador.mtd.Text = "__/__/____"
        End If
        If IsNull(Data1.Recordset("hor_rea")) = False Then
           If Data1.Recordset("hor_rea") <> "" Then
              frm_largador.txt_hortd.Text = Format(Data1.Recordset("hor_rea"), "HH:mm")
           Else
              frm_largador.txt_hortd.Text = "__:__"
           End If
        Else
           frm_largador.txt_hortd.Text = "__:__"
        End If
        If IsNull(Data1.Recordset("diag")) = False Then
           frm_largador.txt_diag.Text = Data1.Recordset("diag")
        Else
           frm_largador.txt_diag.Text = ""
        End If
        If IsNull(Data1.Recordset("colormot")) = False Then
           If Data1.Recordset("colormot") = "R" Then
              frm_largador.cbocolfin.ListIndex = 2
           Else
              If Data1.Recordset("colormot") = "A" Then
                 frm_largador.cbocolfin.ListIndex = 1
              Else
                 If Data1.Recordset("colormot") = "V" Then
                    frm_largador.cbocolfin.ListIndex = 0
                 Else
                    If Data1.Recordset("colormot") = "N" Then
                       frm_largador.cbocolfin.ListIndex = 3
                    Else
                       frm_largador.cbocolfin.Text = ""
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolfin.Text = ""
        End If
        If IsNull(Data1.Recordset("nommed")) = False Then
           frm_largador.dbcbomed.Text = Data1.Recordset("nommed")
        Else
           frm_largador.dbcbomed.ListField = ""
           frm_largador.dbcbomed.BoundColumn = ""
           frm_largador.dbcbomed.Text = ""
           frm_largador.dbcbomed.ListField = "med_nombre"
           frm_largador.dbcbomed.BoundColumn = "med_nombre"
        End If
        If IsNull(Data1.Recordset("codmed")) = False Then
           frm_largador.txt_codmed.Text = Data1.Recordset("codmed")
        Else
           frm_largador.txt_codmed.Text = 0
        End If
        If IsNull(Data1.Recordset("trasla")) = False Then
           If Data1.Recordset("trasla") > 0 Then
              frm_largador.cbotras.ListIndex = Data1.Recordset("trasla")
           Else
              frm_largador.cbotras.ListIndex = 0
           End If
        Else
           frm_largador.cbotras.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("lugar")) = False Then
           frm_largador.txt_lugar.Text = Data1.Recordset("lugar")
        Else
           frm_largador.txt_lugar.Text = ""
        End If
        If IsNull(Data1.Recordset("hsald")) = False Then
           frm_largador.txt_trassal.Text = Format(Data1.Recordset("hsald"), "HH:mm")
        Else
           frm_largador.txt_trassal.Text = ""
        End If
        If IsNull(Data1.Recordset("hllega")) = False Then
           frm_largador.txt_enca.Text = Format(Data1.Recordset("hllega"), "HH:mm")
        Else
           frm_largador.txt_enca.Text = ""
        End If
        If IsNull(Data1.Recordset("hzona")) = False Then
           frm_largador.txt_enzona.Text = Format(Data1.Recordset("hzona"), "HH:mm")
        Else
           frm_largador.txt_enzona.Text = ""
        End If
        If IsNull(Data1.Recordset("movtras")) = False Then
           frm_largador.txt_movtra.Text = Data1.Recordset("movtras")
        Else
           frm_largador.txt_movtra.Text = ""
        End If
        If IsNull(Data1.Recordset("dcobr")) = False Then
           frm_largador.Combo1.Text = Data1.Recordset("dcobr")
        Else
           frm_largador.Combo1.Text = ""
        End If
        If IsNull(Data1.Recordset("totdem")) = False Then
           frm_largador.txt_demora.Text = Format(Data1.Recordset("totdem"), "HH:mm")
        Else
           frm_largador.txt_demora.Text = ""
        End If
        If IsNull(Data1.Recordset("activo")) = False Then
           frm_largador.Label3.Caption = Format(Data1.Recordset("activo"), "HH:mm:ss")
        Else
           frm_largador.Label3.Caption = "00:00:00"
        End If
        If IsNull(Data1.Recordset("timdes")) = False Then
           frm_largador.Label39.Caption = Data1.Recordset("timdes")
        Else
           frm_largador.Label39.Caption = "Sin Largar"
        End If
        If IsNull(Data1.Recordset("motmov")) = True Then
           frm_largador.txt_locali.Text = ""
        Else
           frm_largador.txt_locali.Text = Data1.Recordset("motmov")
        End If
        If IsNull(Data1.Recordset("mm")) = True Then
           frm_largador.Label41.Caption = 0
        Else
           frm_largador.Label41.Caption = Data1.Recordset("mm")
        End If
        If IsNull(Data1.Recordset("thh")) = True Then
           frm_largador.Label42.Caption = 0
        Else
           frm_largador.Label42.Caption = Data1.Recordset("thh")
        End If
        If IsNull(Data1.Recordset("tmm")) = True Then
           frm_largador.Label43.Caption = 0
        Else
           frm_largador.Label43.Caption = Data1.Recordset("tmm")
        End If
        If IsNull(Data1.Recordset("pasado")) = True Then
           frm_largador.Label44.Caption = 0
        Else
           frm_largador.Label44.Caption = Data1.Recordset("pasado")
        End If
        If IsNull(Data1.Recordset("ano")) = True Then
           frm_largador.Label45.Caption = 0
        Else
           frm_largador.Label45.Caption = Data1.Recordset("ano")
        End If
        If IsNull(Data1.Recordset("mes")) = True Then
           frm_largador.Label46.Caption = -1
        Else
           frm_largador.Label46.Caption = Data1.Recordset("mes")
        End If
        If IsNull(Data1.Recordset("timsi")) = True Then
           frm_largador.Label48.Caption = 0
        Else
           frm_largador.Label48.Caption = Data1.Recordset("timsi")
        End If
        If IsNull(Data1.Recordset("enfer")) = True Then
           frm_largador.Check2.Value = 0
        Else
           frm_largador.Check2.Value = Data1.Recordset("enfer")
        End If
        If IsNull(Data1.Recordset("motcance")) = True Then
           frm_largador.txt_quien.Text = ""
        Else
           frm_largador.txt_quien.Text = Data1.Recordset("motcance")
        End If
        If IsNull(Data1.Recordset("hh")) = True Then
           frm_largador.Combo3.ListIndex = -1
        Else
           frm_largador.Combo3.ListIndex = Data1.Recordset("hh")
        End If
        If IsNull(Data1.Recordset("cancela")) = True Then
           If IsNull(Data1.Recordset("hor_cance")) = False Then
              frm_largador.txt_salca.Text = Data1.Recordset("hor_cance")
           Else
              frm_largador.txt_salca.Text = ""
           End If
        End If
        Data3.RecordSource = "Select * from resplla where nro =" & Data1.Recordset("nrolla")
        Data3.Refresh
        If Data3.Recordset.RecordCount > 0 Then
           If IsNull(Data3.Recordset("telef")) = False Then
              If Data3.Recordset("telef") = "RECIBO" Then
                 frm_largador.Combo2.ListIndex = 0
              Else
                 If Data3.Recordset("telef") = "CONFORME" Then
                    frm_largador.Combo2.ListIndex = 1
                 Else
                    frm_largador.Combo2.ListIndex = -1
                 End If
              End If
           Else
              frm_largador.Combo2.ListIndex = -1
           End If
           
           If IsNull(Data3.Recordset("movil_rea")) = False Then
              If Data3.Recordset("movil_rea") > 0 Then
                 frm_largador.data_chof.RecordSource = "Select * from movil where nromov =" & Data3.Recordset("movil_rea")
                 frm_largador.data_chof.Refresh
                 If frm_largador.data_chof.Recordset.RecordCount > 0 Then
                    frm_largador.labcodchof.Caption = Data3.Recordset("movil_rea")
                    frm_largador.labnomchof.Caption = "Chof.:" & frm_largador.data_chof.Recordset("chofer")
                 Else
                    frm_largador.labcodchof.Caption = 0
                    frm_largador.labnomchof.Caption = ""
                 End If
              Else
                 frm_largador.labcodchof.Caption = 0
                 frm_largador.labnomchof.Caption = ""
              End If
           Else
              frm_largador.labcodchof.Caption = 0
              frm_largador.labnomchof.Caption = ""
           End If
           If IsNull(Data3.Recordset("mes")) = False Then
              frm_largador.t_codced.Text = Int(Data3.Recordset("mes"))
           Else
              frm_largador.t_codced.Text = 0
           End If
           If IsNull(Data3.Recordset("pasado")) = False Then
              frm_largador.Check4.Value = Data3.Recordset("pasado")
           Else
              frm_largador.Check4.Value = 0
           End If
           If IsNull(Data3.Recordset("hzona")) = False Then
              frm_largador.labcmt.Visible = True
              frm_largador.labcmt.Caption = "PASADO A CMT HORA:" & Format(Data3.Recordset("hzona"), "HH:mm")
              If IsNull(Data3.Recordset("mm")) = False Then
                 If Data3.Recordset("mm") = 1 Then
                    frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " NO RESUELTO H."
                    If IsNull(Data3.Recordset("hsald")) = False Then
                       frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & Data3.Recordset("hsald")
                    End If
                    If IsNull(Data3.Recordset("totend")) = False Then
                       If Data3.Recordset("totend") = "R" Then
                          frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RECLASIFICA A ROJO"
                       End If
                       If Data3.Recordset("totend") = "A" Then
                          frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RECLASIFICA A AMARILLO"
                       End If
                    End If
                 End If
                 If Data3.Recordset("mm") = 2 Then
                    frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RESUELTO HORA:"
                    If IsNull(Data3.Recordset("hor_rea")) = False Then
                       frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & Data3.Recordset("hor_rea")
                    End If
                 End If
                 If Data3.Recordset("mm") = 2 Or Data3.Recordset("mm") = 3 Then
                    frm_largador.Command5.Visible = True
                    frm_largador.Frame2.Visible = False
                 Else
                    If Data3.Recordset("mm") = 1 Then
                       frm_largador.Command5.Visible = False
                       If WDespa = 1 Then
                          frm_largador.Frame2.Visible = False
                       Else
                          frm_largador.Frame2.Visible = True
                       End If
                    Else
                       frm_largador.Command5.Visible = True
                       frm_largador.Frame2.Visible = False
                    End If
                 End If
              Else
                 frm_largador.Command5.Visible = True
                 frm_largador.Frame2.Visible = False
              End If
           Else
              frm_largador.labcmt.Caption = ""
              frm_largador.labcmt.Visible = False
              frm_largador.Command5.Visible = False
              If WDespa = 1 Then
                 frm_largador.Frame2.Visible = False
              Else
                 frm_largador.Frame2.Visible = True
              End If
           End If
           If IsNull(Data3.Recordset("movilpas")) = False Then
              Data4.Recordset.FindFirst "med_cod =" & Data3.Recordset("movilpas")
              If Not Data4.Recordset.NoMatch Then
                 frm_largador.dbcbomed2.Text = Data4.Recordset("med_nombre")
              Else
                 frm_largador.dbcbomed2.ListField = ""
                 frm_largador.dbcbomed2.BoundColumn = ""
                 frm_largador.dbcbomed2.Text = ""
                 frm_largador.dbcbomed2.ListField = "med_nombre"
                 frm_largador.dbcbomed2.BoundColumn = "med_nombre"
              End If
              frm_largador.txt_codmed2.Text = Data3.Recordset("movilpas")
           Else
              frm_largador.txt_codmed2.Text = 0
              frm_largador.dbcbomed2.ListField = ""
              frm_largador.dbcbomed2.BoundColumn = ""
              frm_largador.dbcbomed2.Text = ""
              frm_largador.dbcbomed2.ListField = "med_nombre"
             frm_largador.dbcbomed2.BoundColumn = "med_nombre"
           End If
           If IsNull(Data3.Recordset("fec_llega")) = False Then
              frm_largador.mftrassol.Text = Format(Data3.Recordset("fec_llega"), "dd/mm/yyyy")
           Else
              frm_largador.mftrassol.Text = "__/__/____"
           End If
           If IsNull(Data3.Recordset("hor_llega")) = False Then
              frm_largador.mhtrassol.Text = Format(Data3.Recordset("hor_llega"), "HH:mm")
           Else
              frm_largador.mhtrassol.Text = "__:__"
           End If
        Else
           frm_largador.txt_codmed2.Text = 0
           frm_largador.Check4.Value = 0
           frm_largador.dbcbomed2.Text = ""
           frm_largador.mftrassol.Text = "__/__/____"
           frm_largador.mhtrassol.Text = "__:__"
           frm_largador.t_codced.Text = 0
           frm_largador.labcmt.Caption = ""
           frm_largador.labcmt.Visible = False
        End If
        
        
        Unload Me

Exit Sub

Quepasoalabrir:
            If Err.Number > 0 Then
               MsgBox " ERROR: " & str(Err.Number) & " Vuelva a intentar abrir o cierre el sistema"
               If tabver.SelectedItem.index = 1 Then
                  DBGrid1.Visible = True
                  DBGrid2.Visible = False
                  DBGrid3.Visible = False
                  DBGrid1.BackColor = &HFFC0C0
                  Data1.RecordSource = "Select * from llamado where pend <>" & 2 & " And pend <>" & 1 & " and pend <>" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by nrolla"
                  Data1.Refresh
                  labtoregi.Caption = Data1.Recordset.RecordCount
               Else
                  If tabver.SelectedItem.index = 2 Then
                     DBGrid1.Visible = True
                     DBGrid2.Visible = False
                     DBGrid3.Visible = False
                     DBGrid1.BackColor = &HC0C0FF
                     Data1.RecordSource = "Select * from llamado where pend =" & 1 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by nrolla"
                     Data1.Refresh
                     labtoregi.Caption = Data1.Recordset.RecordCount
                  Else
                     If tabver.SelectedItem.index = 3 Then
                        DBGrid1.BackColor = &HC0FFC0
                        DBGrid1.Visible = False
                        DBGrid2.Visible = False
                        DBGrid3.Visible = True
'                       Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/12/2011", "yyyy/mm/dd") & "# and trasla >=" & 1 & " and hzona ='" & Null & "' order by nrolla"
                        Data1.RecordSource = "Select * from llamado where trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14) and fecha >=#" & Format("08/08/2019", "yyyy/mm/dd") & "# and hzona ='" & Null & "' and codmot <>'" & "Z" & "' order by nrolla"
                        Data1.Refresh
                        labtoregi.Caption = Data1.Recordset.RecordCount
                     Else
                        If tabver.SelectedItem.index = 4 Then
                           DBGrid1.Visible = True
                           DBGrid2.Visible = False
                           DBGrid3.Visible = False
                           DBGrid1.BackColor = &H80FF&
                           If Trim(labcodus.Caption) <> "" Then
                              Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & Val(labcodus.Caption) & " order by fecha,hora"
                           Else
                              If ControlUsuario("Utilitarios despacho") = 1 Then
                                 Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
                              Else
                                 Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & 999999 & " order by fecha,hora"
                              End If
                           End If
                           Data1.Refresh
                           labtoregi.Caption = Data1.Recordset.RecordCount
                        Else
                           If tabver.SelectedItem.index = 5 Then
                              DBGrid1.Visible = True
                              DBGrid2.Visible = False
                              DBGrid3.Visible = False
                              DBGrid1.BackColor = &HC000&
'                              Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) and (isopa_result is null or isopa_result in ('Negativo')) order by fecha,hora"
                              Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and prox_control >=#" & Format(XfecCovid, "yyyy/mm/dd") & "# and prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# and (isopa_result is null or isopa_result in ('Negativo')) order by nombre,fecha,hora"
'                              Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null or prox_control) and prox_control >=#" & Format(XfecCovid, "yyyy/mm/dd") & "# and (isopa_result is null or isopa_result in ('Negativo')) order by fecha,hora"
                              Data1.Refresh
                              labtoregi.Caption = Data1.Recordset.RecordCount
                           Else
                              If tabver.SelectedItem.index = 6 Then
                                 DBGrid1.Visible = True
                                 DBGrid2.Visible = False
                                 DBGrid3.Visible = False
                                 DBGrid1.BackColor = &H80FF80
                                 Data1.RecordSource = "Select * from llamado where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null order by fecha,hora"
                                 Data1.Refresh
                                 labtoregi.Caption = Data1.Recordset.RecordCount
                              Else
                                 If tabver.SelectedItem.index = 7 Then
                                    DBGrid1.Visible = True
                                    DBGrid2.Visible = False
                                    DBGrid3.Visible = False
                                    DBGrid1.BackColor = &H8080FF
                                    Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
                                    Data1.Refresh
                                    labtoregi.Caption = Data1.Recordset.RecordCount
                                 Else
                                    If tabver.SelectedItem.index = 8 Then
                                       DBGrid1.Visible = False
                                       DBGrid2.Visible = True
                                       DBGrid3.Visible = False
                                       Data6.RecordSource = "select * from sol_hisopos where deriva in (1) and fecha_cierre is null order by fecha"
                                       Data6.Refresh
                                       labtoregi.Caption = Data6.Recordset.RecordCount
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               Unload Me
            End If

End Sub

Private Sub DBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Quepasaalver

If KeyCode = vbKeyReturn Then
   DBGrid3_DblClick

End If

Exit Sub

Quepasaalver:
             If Err.Number > 0 Then
                MsgBox "Error al abrir, reintente o cierre ésta pantalla y vuelva a abrir", vbInformation
             End If
End Sub

Private Sub Form_Deactivate()
'Timer1.Enabled = False


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   If ControlUsuario("Utilitarios despacho") = 1 Then
      If tabver.SelectedItem.index = 1 Then
         tabver.Tabs(2).Selected = True
      Else
         tabver.Tabs(1).Selected = True
      End If
   Else
      tabver.Tabs(4).Selected = True
   End If
End If

End Sub

Private Sub Form_Load()
Dim Xlaf, XfecCovid As Date
Xlaf = Date - 15
XfecCovid = Date - 5
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data6.Connect = "odbc;dsn=" & Xconexrmt & ";"
labcodus.Caption = Devuelve_user()

'Data1.ConnectionString = "dsn=" & Xconexrmt
'data_entrantes.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_entrantes.ConnectionString = "dsn=" & Xconexrmt

'Data1.RecordSource = "llamado"
'Data1.Refresh
tabver.Refresh
'tabver.SelectedItem.Selected = True
If ControlUsuario("Utilitarios despacho") <> 1 Then
   tabver.Tabs(4).Selected = True
End If

If tabver.SelectedItem.index = 1 Then
   DBGrid1.Visible = True
   DBGrid2.Visible = False
   DBGrid3.Visible = False
   DBGrid1.BackColor = &HFFC0C0
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend not in (1,2,4) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by hora"
   Data1.Refresh
   labtoregi.Caption = Data1.Recordset.RecordCount
Else
   If tabver.SelectedItem.index = 2 Then
      DBGrid1.Visible = True
      DBGrid2.Visible = False
      DBGrid3.Visible = False
      DBGrid1.BackColor = &HC0C0FF
      Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend in (1) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by hora"
      Data1.Refresh
      labtoregi.Caption = Data1.Recordset.RecordCount
   Else
      If tabver.SelectedItem.index = 3 Then
         DBGrid1.BackColor = &HC0FFC0
         DBGrid1.Visible = False
         DBGrid2.Visible = False
         DBGrid3.Visible = True
         Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14) and fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# and hzona ='" & Null & "' and codmot <>'" & "Z" & "' order by hora"
         Data1.Refresh
         labtoregi.Caption = Data1.Recordset.RecordCount
      Else
         If tabver.SelectedItem.index = 4 Then
            DBGrid1.Visible = True
            DBGrid2.Visible = False
            DBGrid3.Visible = False
            DBGrid1.BackColor = &H80FF&
            If Trim(labcodus.Caption) <> "" Then
               Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & Val(labcodus.Caption) & " order by fecha,hora"
            Else
               If ControlUsuario("Utilitarios despacho") = 1 Then
                  Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
               Else
                  Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & 999999 & " order by fecha,hora"
               End If
            End If
            Data1.Refresh
            labtoregi.Caption = Data1.Recordset.RecordCount
         Else
            If tabver.SelectedItem.index = 5 Then
               DBGrid1.Visible = True
               DBGrid2.Visible = False
               DBGrid3.Visible = False
               DBGrid1.BackColor = &HC000&
               Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and prox_control >=#" & Format(XfecCovid, "yyyy/mm/dd") & "# and prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# and (isopa_result is null or isopa_result in ('Negativo')) order by nombre,fecha,hora"
               Data1.Refresh
               labtoregi.Caption = Data1.Recordset.RecordCount
            Else
               If tabver.SelectedItem.index = 6 Then
                  DBGrid1.Visible = True
                  DBGrid2.Visible = False
                  DBGrid3.Visible = False
                  DBGrid1.BackColor = &H80FF80
                  Data1.RecordSource = "Select * from llamado where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null order by fecha,hora"
                  Data1.Refresh
                  labtoregi.Caption = Data1.Recordset.RecordCount
               Else
                  If tabver.SelectedItem.index = 7 Then
                     DBGrid1.Visible = True
                     DBGrid2.Visible = False
                     DBGrid3.Visible = False
                     DBGrid1.BackColor = &H8080FF
                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
                     Data1.Refresh
                     labtoregi.Caption = Data1.Recordset.RecordCount
                  Else
                     If tabver.SelectedItem.index = 8 Then
                        DBGrid1.Visible = False
                        DBGrid2.Visible = True
                        DBGrid3.Visible = False
                        Data6.RecordSource = "select * from sol_hisopos where deriva in (1) and fecha_cierre is null order by fecha"
                        Data6.Refresh
                        labtoregi.Caption = Data6.Recordset.RecordCount
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If
'&H0080FF80&
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data4.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data4.RecordSource = "medicos"
Data4.Refresh

Data5.Connect = "odbc;dsn=" & Xconexrmt & ";"

If ControlUsuario("Utilitarios despacho") = 1 Then
   DBGrid1.Columns(0).Visible = True
   DBGrid1.Columns(1).Visible = True
   DBGrid1.Columns(3).Visible = True
   DBGrid1.Columns(12).Visible = True
Else
   DBGrid1.Columns(0).Visible = False
   DBGrid1.Columns(1).Visible = False
   DBGrid1.Columns(3).Visible = False
   DBGrid1.Columns(12).Visible = False
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

Private Sub Form_Unload(Cancel As Integer)
'Timer1.Enabled = False

End Sub

Private Sub tabver_Click()
Dim Xlaf, XfecCovid As Date
Xlaf = Date - 15
XfecCovid = Date - 5
If ControlUsuario("Utilitarios despacho") <> 1 Then
   If tabver.SelectedItem.index = 1 Or tabver.SelectedItem.index = 2 Or tabver.SelectedItem.index = 3 Then
      MsgBox "Usuario no autorizado"
      tabver.Tabs(4).Selected = True
   End If
End If

If tabver.SelectedItem.index = 1 Then
   DBGrid1.Visible = True
   DBGrid2.Visible = False
   DBGrid3.Visible = False
   DBGrid1.BackColor = &HFFC0C0
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend not in (1,2,4) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by hora"
   Data1.Refresh
   labtoregi.Caption = Data1.Recordset.RecordCount
Else
   If tabver.SelectedItem.index = 2 Then
      DBGrid1.Visible = True
      DBGrid2.Visible = False
      DBGrid3.Visible = False
      DBGrid1.BackColor = &HC0C0FF
      Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend in (1) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by hora"
      Data1.Refresh
      labtoregi.Caption = Data1.Recordset.RecordCount
   Else
      If tabver.SelectedItem.index = 3 Then
         DBGrid1.BackColor = &HC0FFC0
         DBGrid1.Visible = False
         DBGrid2.Visible = False
         DBGrid3.Visible = True
         Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# and hzona ='" & Null & "' and codmot <>'" & "Z" & "' order by hora"
         Data1.Refresh
         labtoregi.Caption = Data1.Recordset.RecordCount
      Else
         If tabver.SelectedItem.index = 4 Then
            DBGrid1.Visible = True
            DBGrid2.Visible = False
            DBGrid3.Visible = False
            DBGrid1.BackColor = &H80FF&
            If Trim(labcodus.Caption) <> "" Then
               Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & Val(labcodus.Caption) & " order by fecha,hora"
            Else
               If ControlUsuario("Utilitarios despacho") = 1 Then
                  Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
               Else
                  Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & 999999 & " order by fecha,hora"
               End If
            End If
            Data1.Refresh
            labtoregi.Caption = Data1.Recordset.RecordCount
         Else
            If tabver.SelectedItem.index = 5 Then
               DBGrid1.Visible = True
               DBGrid2.Visible = False
               DBGrid3.Visible = False
               DBGrid1.BackColor = &HC000&
               Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and prox_control >=#" & Format(XfecCovid, "yyyy/mm/dd") & "# and prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# and (isopa_result is null or isopa_result in ('Negativo')) order by nombre,fecha,hora"
'Select * from llamado where fecha >='2019-07-01' and segui_covid in (1) and cierre_hora is null and (prox_control <='2021-01-19' or prox_control is null) and prox_control >='2021-01-04' and (isopa_result is null or isopa_result in ('Negativo')) order by fecha,hora
               Data1.Refresh
               labtoregi.Caption = Data1.Recordset.RecordCount
            Else
               If tabver.SelectedItem.index = 6 Then
                  DBGrid1.Visible = True
                  DBGrid2.Visible = False
                  DBGrid3.Visible = False
                  DBGrid1.BackColor = &H80FF80
                  Data1.RecordSource = "Select * from llamado where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null order by fecha,hora"
                  Data1.Refresh
                  labtoregi.Caption = Data1.Recordset.RecordCount
               Else
                  If tabver.SelectedItem.index = 7 Then
                     DBGrid1.Visible = True
                     DBGrid2.Visible = False
                     DBGrid3.Visible = False
                     DBGrid1.BackColor = &H8080FF
                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
                     Data1.Refresh
                     labtoregi.Caption = Data1.Recordset.RecordCount
                  Else
                     If tabver.SelectedItem.index = 8 Then
                        DBGrid1.Visible = False
                        DBGrid2.Visible = True
                        DBGrid3.Visible = False
                        Data6.RecordSource = "select * from sol_hisopos where deriva in (1) and fecha_cierre is null order by fecha"
                        Data6.Refresh
                        labtoregi.Caption = Data6.Recordset.RecordCount
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If

'&H00C0C000&

End Sub

Private Sub tabver_GotFocus()
Dim XfecCovid As Date
XfecCovid = Date - 5
Xlaf = Date - 15

If tabver.SelectedItem.index = 1 Then
   DBGrid1.Visible = True
   DBGrid2.Visible = False
   DBGrid3.Visible = False
   DBGrid1.BackColor = &HFFC0C0
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend not in (1,2,4) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by hora"
   Data1.Refresh
   labtoregi.Caption = Data1.Recordset.RecordCount
Else
   If tabver.SelectedItem.index = 2 Then
      DBGrid1.Visible = True
      DBGrid2.Visible = False
      DBGrid3.Visible = False
      DBGrid1.BackColor = &HC0C0FF
      Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend in (1) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by hora"
      Data1.Refresh
      labtoregi.Caption = Data1.Recordset.RecordCount
   Else
      If tabver.SelectedItem.index = 3 Then
         DBGrid1.BackColor = &HC0FFC0
         DBGrid1.Visible = False
         DBGrid2.Visible = False
         DBGrid3.Visible = True
         Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14) and fecha >=#" & Format("08/08/2012", "yyyy/mm/dd") & "# and hzona ='" & Null & "' and codmot <>'" & "Z" & "' order by hora"
         Data1.Refresh
         labtoregi.Caption = Data1.Recordset.RecordCount
      Else
         If tabver.SelectedItem.index = 4 Then
            DBGrid1.Visible = True
            DBGrid2.Visible = False
            DBGrid3.Visible = False
            DBGrid1.BackColor = &H80FF&
            If Trim(labcodus.Caption) <> "" Then
               Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & Val(labcodus.Caption) & " order by fecha,hora"
            Else
               If ControlUsuario("Utilitarios despacho") = 1 Then
                  Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
               Else
                  Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) and codmedcmt =" & 999999 & " order by fecha,hora"
               End If
            End If
            Data1.Refresh
            labtoregi.Caption = Data1.Recordset.RecordCount
         Else
            If tabver.SelectedItem.index = 5 Then
               DBGrid1.Visible = True
               DBGrid2.Visible = False
               DBGrid3.Visible = False
               DBGrid1.BackColor = &HC000&
'               Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null or prox_control) and prox_control >=#" & Format(XfecCovid, "yyyy/mm/dd") & "# and (isopa_result is null or isopa_result in ('Negativo')) order by fecha,hora"
               Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and prox_control >=#" & Format(XfecCovid, "yyyy/mm/dd") & "# and prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# and (isopa_result is null or isopa_result in ('Negativo')) order by nombre,fecha,hora"
'               Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) and (isopa_result is null or isopa_result in ('Negativo')) order by fecha,hora"
               Data1.Refresh
               labtoregi.Caption = Data1.Recordset.RecordCount
            Else
               If tabver.SelectedItem.index = 6 Then
                  DBGrid1.Visible = True
                  DBGrid2.Visible = False
                  DBGrid3.Visible = False
                  DBGrid1.BackColor = &H80FF80
                  Data1.RecordSource = "Select * from llamado where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null order by fecha,hora"
                  Data1.Refresh
                  labtoregi.Caption = Data1.Recordset.RecordCount
               Else
                  If tabver.SelectedItem.index = 7 Then
                     DBGrid1.Visible = True
                     DBGrid2.Visible = False
                     DBGrid3.Visible = False
                     DBGrid1.BackColor = &H8080FF
                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
                     Data1.Refresh
                     labtoregi.Caption = Data1.Recordset.RecordCount
                  Else
                     If tabver.SelectedItem.index = 8 Then
                        DBGrid1.Visible = False
                        DBGrid2.Visible = True
                        DBGrid3.Visible = False
                        Data6.RecordSource = "select * from sol_hisopos where deriva in (1) and fecha_cierre is null order by fecha"
                        Data6.Refresh
                        labtoregi.Caption = Data6.Recordset.RecordCount
                     
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If

End Sub

Private Sub tabver_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBGrid1.SetFocus
End If

End Sub


Private Sub Timer1_Timer()
'Data1.Refresh

data_entrantes.RecordSource = "Select * from llamado where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and codmot ='" & "R" & "' and movilpas =" & 0
data_entrantes.Refresh
If data_entrantes.Recordset.RecordCount > 0 Then
   MsgBox "TIENE LLAMADOS ROJOS SIN DESPACHAR, VERIFIQUE!!", vbCritical, "DESPACHO"
   Xentrantes = Xentrantes + 1
   If Xentrantes = 1 Then
      Timer1.Interval = Timer1.Interval + 4000
   End If
   If Xentrantes = 2 Then
      Timer1.Interval = Timer1.Interval + 5000
   End If
   If Xentrantes = 3 Then
      Timer1.Interval = Timer1.Interval + 6000
   End If
   If Xentrantes = 4 Then
      Xentrantes = 0
      Timer1.Interval = 2000
   End If
End If

End Sub

Public Function Devuelve_user() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from usuarios where usuario ='" & WElusuario & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("codmed")) = False Then
      Devuelve_user = Xrecclii("codmed")
   Else
      Devuelve_user = ""
   End If
Else
   Devuelve_user = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

