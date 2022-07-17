VERSION 5.00
Begin VB.Form frm_sincroniza 
   Caption         =   "Sincronizador despacho"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6810
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2760
      Top             =   1560
   End
   Begin VB.Data data_llarem2 
      Caption         =   "data_llarem2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data data_llaloc2 
      Caption         =   "data_llaloc2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar datos"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data data_llaloc 
      Caption         =   "data_llaloc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data data_llarem 
      Caption         =   "data_llarem"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frm_sincroniza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XX As Integer

Private Sub Command1_Click()
Dim Xlaf As Date
Dim Xmodi As Integer
Xmodi = 0
Xlaf = Date - 1
'         data_lla.Recordset("nrolla") = txt_nro.Text
'         data_lla.Recordset("nro") = txt_nro.Text
Timer1.Enabled = False
If XX >= 6 Then

    data_llarem.RecordSource = "Select * from llamado where fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# order by fecha"
    data_llarem.Refresh
'    MsgBox "Comienza"
    If data_llarem.Recordset.RecordCount > 0 Then
       data_llarem.Recordset.MoveFirst
       Do While Not data_llarem.Recordset.EOF
          data_llaloc.RecordSource = "Select * from llamado where nrolla =" & data_llarem.Recordset("nrolla")
          data_llaloc.Refresh
          If data_llaloc.Recordset.RecordCount > 0 Then
             Xmodi = 0
             data_llaloc.Recordset.Edit
             If IsNull(data_llaloc.Recordset("fecha")) = False And IsNull(data_llarem.Recordset("fecha")) = False Then
                If data_llaloc.Recordset("fecha") <> data_llarem.Recordset("fecha") Then
                   Xmodi = 6
                   data_llaloc.Recordset("fecha") = data_llarem.Recordset("fecha")
                End If
             Else
                If IsNull(data_llarem.Recordset("fecha")) = True Then
                   If IsNull(data_llaloc.Recordset("fecha")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("fecha") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("fecha")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("fecha") = data_llarem.Recordset("fecha")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("hora")) = False And IsNull(data_llarem.Recordset("hora")) = False Then
                If data_llaloc.Recordset("hora") <> data_llarem.Recordset("hora") Then
                   Xmodi = 6
                   data_llaloc.Recordset("hora") = data_llarem.Recordset("hora")
                End If
             Else
                If IsNull(data_llarem.Recordset("hora")) = True Then
                   If IsNull(data_llaloc.Recordset("hora")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("hora") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("hora")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("hora") = data_llarem.Recordset("hora")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("usuario")) = False And IsNull(data_llarem.Recordset("usuario")) = False Then
                If data_llaloc.Recordset("usuario") <> data_llarem.Recordset("usuario") Then
                   Xmodi = 6
                   data_llaloc.Recordset("usuario") = data_llarem.Recordset("usuario")
                End If
             Else
                Xmodi = 6
                data_llaloc.Recordset("usuario") = data_llarem.Recordset("usuario")
             End If
             If IsNull(data_llaloc.Recordset("matric")) = False And IsNull(data_llarem.Recordset("matric")) = False Then
                If data_llaloc.Recordset("matric") <> data_llarem.Recordset("matric") Then
                   Xmodi = 6
                   data_llaloc.Recordset("matric") = data_llarem.Recordset("matric")
                End If
             Else
                If IsNull(data_llarem.Recordset("matric")) = True Then
                   If IsNull(data_llaloc.Recordset("matric")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("matric") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("matric")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("matric") = data_llarem.Recordset("matric")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("nombre")) = False And IsNull(data_llarem.Recordset("nombre")) = False Then
                If data_llaloc.Recordset("nombre") <> data_llarem.Recordset("nombre") Then
                   Xmodi = 6
                   data_llaloc.Recordset("nombre") = data_llarem.Recordset("nombre")
                End If
             Else
                Xmodi = 6
                data_llaloc.Recordset("nombre") = data_llarem.Recordset("nombre")
             End If
             If IsNull(data_llaloc.Recordset("edad")) = False And IsNull(data_llarem.Recordset("edad")) = False Then
                If data_llaloc.Recordset("edad") <> data_llarem.Recordset("edad") Then
                   Xmodi = 6
                   data_llaloc.Recordset("edad") = data_llarem.Recordset("edad")
                End If
             Else
                If IsNull(data_llarem.Recordset("edad")) = True Then
                   If IsNull(data_llaloc.Recordset("edad")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("edad") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("edad")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("edad") = data_llarem.Recordset("edad")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("unied")) = False And IsNull(data_llarem.Recordset("unied")) = False Then
                If data_llaloc.Recordset("unied") <> data_llarem.Recordset("unied") Then
                   Xmodi = 6
                   data_llaloc.Recordset("unied") = data_llarem.Recordset("unied")
                End If
             Else
                If IsNull(data_llarem.Recordset("unied")) = True Then
                   If IsNull(data_llaloc.Recordset("unied")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("unied") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("unied")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("unied") = data_llarem.Recordset("unied")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("categ")) = False And IsNull(data_llarem.Recordset("categ")) = False Then
                If data_llaloc.Recordset("categ") <> data_llarem.Recordset("categ") Then
                   Xmodi = 6
                   data_llaloc.Recordset("categ") = data_llarem.Recordset("categ")
                   data_llaloc.Recordset("nomcat") = data_llarem.Recordset("nomcat")
                End If
             Else
                Xmodi = 6
                data_llaloc.Recordset("categ") = data_llarem.Recordset("categ")
                data_llaloc.Recordset("nomcat") = data_llarem.Recordset("nomcat")
             End If
             If IsNull(data_llaloc.Recordset("ci")) = False And IsNull(data_llarem.Recordset("ci")) = False Then
                If data_llaloc.Recordset("ci") <> data_llarem.Recordset("ci") Then
                   Xmodi = 6
                   data_llaloc.Recordset("ci") = data_llarem.Recordset("ci")
                End If
             Else
                If IsNull(data_llarem.Recordset("ci")) = True Then
                   If IsNull(data_llaloc.Recordset("ci")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("ci") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("ci")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("ci") = data_llarem.Recordset("ci")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("telef")) = False And IsNull(data_llarem.Recordset("telef")) = False Then
                If data_llaloc.Recordset("telef") <> data_llarem.Recordset("telef") Then
                   Xmodi = 6
                   data_llaloc.Recordset("telef") = data_llarem.Recordset("telef")
                End If
             Else
                If IsNull(data_llarem.Recordset("telef")) = True Then
                   If IsNull(data_llaloc.Recordset("telef")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("telef") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("telef")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("telef") = data_llarem.Recordset("telef")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("codzon")) = False And IsNull(data_llarem.Recordset("codzon")) = False Then
                If data_llaloc.Recordset("codzon") <> data_llarem.Recordset("codzon") Then
                   Xmodi = 6
                   data_llaloc.Recordset("codzon") = data_llarem.Recordset("codzon")
                End If
             Else
                   Xmodi = 6
                   data_llaloc.Recordset("codzon") = data_llarem.Recordset("codzon")
             End If
             If IsNull(data_llaloc.Recordset("base")) = False And IsNull(data_llarem.Recordset("base")) = False Then
                If data_llaloc.Recordset("base") <> data_llarem.Recordset("base") Then
                   Xmodi = 6
                   data_llaloc.Recordset("base") = data_llarem.Recordset("base")
                End If
             Else
                   Xmodi = 6
                   data_llaloc.Recordset("base") = data_llarem.Recordset("base")
             End If
             If IsNull(data_llaloc.Recordset("referen")) = False And IsNull(data_llarem.Recordset("referen")) = False Then
                If data_llaloc.Recordset("referen") <> data_llarem.Recordset("referen") Then
                   Xmodi = 6
                   data_llaloc.Recordset("referen") = data_llarem.Recordset("referen")
                End If
             Else
                If IsNull(data_llarem.Recordset("referen")) = True Then
                   If IsNull(data_llaloc.Recordset("referen")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("referen") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("referen")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("referen") = data_llarem.Recordset("referen")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("realiza")) = False And IsNull(data_llarem.Recordset("realiza")) = False Then
                If data_llaloc.Recordset("realiza") <> data_llarem.Recordset("realiza") Then
                   Xmodi = 6
                   data_llaloc.Recordset("realiza") = data_llarem.Recordset("realiza")
                End If
             Else
                If IsNull(data_llarem.Recordset("realiza")) = True Then
                   If IsNull(data_llaloc.Recordset("realiza")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("realiza") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("realiza")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("realiza") = data_llarem.Recordset("realiza")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("motcon")) = False And IsNull(data_llarem.Recordset("motcon")) = False Then
                If data_llaloc.Recordset("motcon") <> data_llarem.Recordset("motcon") Then
                   Xmodi = 6
                   data_llaloc.Recordset("motcon") = data_llarem.Recordset("motcon")
                End If
             Else
                If IsNull(data_llarem.Recordset("motcon")) = True Then
                   If IsNull(data_llaloc.Recordset("motcon")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("motcon") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("motcon")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("motcon") = data_llarem.Recordset("motcon")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("obsmot")) = False And IsNull(data_llarem.Recordset("obsmot")) = False Then
                If data_llaloc.Recordset("obsmot") <> data_llarem.Recordset("obsmot") Then
                   Xmodi = 6
                   data_llaloc.Recordset("obsmot") = data_llarem.Recordset("obsmot")
                End If
             Else
                If IsNull(data_llarem.Recordset("obsmot")) = True Then
                   If IsNull(data_llaloc.Recordset("obsmot")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("obsmot") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("obsmot")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("obsmot") = data_llarem.Recordset("obsmot")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("codmot")) = False And IsNull(data_llarem.Recordset("codmot")) = False Then
                If data_llaloc.Recordset("codmot") <> data_llarem.Recordset("codmot") Then
                   Xmodi = 6
                   data_llaloc.Recordset("codmot") = data_llarem.Recordset("codmot")
                   data_llaloc.Recordset("descol") = data_llarem.Recordset("descol")
                End If
             Else
                   Xmodi = 6
                   data_llaloc.Recordset("codmot") = data_llarem.Recordset("codmot")
                   data_llaloc.Recordset("descol") = data_llarem.Recordset("descol")
             End If
             If IsNull(data_llaloc.Recordset("movilpas")) = False And IsNull(data_llarem.Recordset("movilpas")) = False Then
                If data_llaloc.Recordset("movilpas") <> data_llarem.Recordset("movilpas") Then
                   Xmodi = 6
                   data_llaloc.Recordset("movilpas") = data_llarem.Recordset("movilpas")
                End If
             Else
                If IsNull(data_llarem.Recordset("movilpas")) = True Then
                   If IsNull(data_llaloc.Recordset("movilpas")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("movilpas") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("movilpas")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("movilpas") = data_llarem.Recordset("movilpas")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("pend")) = False And IsNull(data_llarem.Recordset("pend")) = False Then
                If data_llaloc.Recordset("pend") <> data_llarem.Recordset("pend") Then
                   Xmodi = 6
                   data_llaloc.Recordset("pend") = data_llarem.Recordset("pend")
                End If
             Else
                   Xmodi = 6
                   data_llaloc.Recordset("pend") = data_llarem.Recordset("pend")
             End If
             If IsNull(data_llaloc.Recordset("timdes")) = False And IsNull(data_llarem.Recordset("timdes")) = False Then
                If data_llaloc.Recordset("timdes") <> data_llarem.Recordset("timdes") Then
                   Xmodi = 6
                   data_llaloc.Recordset("timdes") = data_llarem.Recordset("timdes")
                End If
             Else
                If IsNull(data_llarem.Recordset("timdes")) = True Then
                   If IsNull(data_llaloc.Recordset("timdes")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("timdes") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("timdes")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("timdes") = data_llarem.Recordset("timdes")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("fec_rea")) = False And IsNull(data_llarem.Recordset("fec_rea")) = False Then
                If data_llaloc.Recordset("fec_rea") <> data_llarem.Recordset("fec_rea") Then
                   Xmodi = 6
                   data_llaloc.Recordset("fec_rea") = data_llarem.Recordset("fec_rea")
                End If
             Else
                If IsNull(data_llarem.Recordset("fec_rea")) = True Then
                   If IsNull(data_llaloc.Recordset("fec_rea")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("fec_rea") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("fec_rea")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("fec_rea") = data_llarem.Recordset("fec_rea")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("hor_rea")) = False And IsNull(data_llarem.Recordset("hor_rea")) = False Then
                If data_llaloc.Recordset("hor_rea") <> data_llarem.Recordset("hor_rea") Then
                   Xmodi = 6
                   data_llaloc.Recordset("hor_rea") = data_llarem.Recordset("hor_rea")
                End If
             Else
                If IsNull(data_llarem.Recordset("hor_rea")) = True Then
                   If IsNull(data_llaloc.Recordset("hor_rea")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("hor_rea") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("hor_rea")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("hor_rea") = data_llarem.Recordset("hor_rea")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("fecpas")) = False And IsNull(data_llarem.Recordset("fecpas")) = False Then
                If data_llaloc.Recordset("fecpas") <> data_llarem.Recordset("fecpas") Then
                   Xmodi = 6
                   data_llaloc.Recordset("fecpas") = data_llarem.Recordset("fecpas")
                End If
             Else
                If IsNull(data_llarem.Recordset("fecpas")) = True Then
                   If IsNull(data_llaloc.Recordset("fecpas")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("fecpas") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("fecpas")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("fecpas") = data_llarem.Recordset("fecpas")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("horpas")) = False And IsNull(data_llarem.Recordset("horpas")) = False Then
                If data_llaloc.Recordset("horpas") <> data_llarem.Recordset("horpas") Then
                   Xmodi = 6
                   data_llaloc.Recordset("horpas") = data_llarem.Recordset("horpas")
                End If
             Else
                If IsNull(data_llarem.Recordset("horpas")) = True Then
                   If IsNull(data_llaloc.Recordset("horpas")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("horpas") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("horpas")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("horpas") = data_llarem.Recordset("horpas")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("fecsali")) = False And IsNull(data_llarem.Recordset("fecsali")) = False Then
                If data_llaloc.Recordset("fecsali") <> data_llarem.Recordset("fecsali") Then
                   Xmodi = 6
                   data_llaloc.Recordset("fecsali") = data_llarem.Recordset("fecsali")
                End If
             Else
                If IsNull(data_llarem.Recordset("fecsali")) = True Then
                   If IsNull(data_llaloc.Recordset("fecsali")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("fecsali") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("fecsali")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("fecsali") = data_llarem.Recordset("fecsali")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("horsali")) = False And IsNull(data_llarem.Recordset("horsali")) = False Then
                If data_llaloc.Recordset("horsali") <> data_llarem.Recordset("horsali") Then
                   Xmodi = 6
                   data_llaloc.Recordset("horsali") = data_llarem.Recordset("horsali")
                End If
             Else
                If IsNull(data_llarem.Recordset("horsali")) = True Then
                   If IsNull(data_llaloc.Recordset("horsali")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("horsali") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("horsali")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("horsali") = data_llarem.Recordset("horsali")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("fec_llega")) = False And IsNull(data_llarem.Recordset("fec_llega")) = False Then
                If data_llaloc.Recordset("fec_llega") <> data_llarem.Recordset("fec_llega") Then
                   Xmodi = 6
                   data_llaloc.Recordset("fec_llega") = data_llarem.Recordset("fec_llega")
                End If
             Else
                If IsNull(data_llarem.Recordset("fec_llega")) = True Then
                   If IsNull(data_llaloc.Recordset("fec_llega")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("fec_llega") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("fec_llega")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("fec_llega") = data_llarem.Recordset("fec_llega")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("hor_llega")) = False And IsNull(data_llarem.Recordset("hor_llega")) = False Then
                If data_llaloc.Recordset("hor_llega") <> data_llarem.Recordset("hor_llega") Then
                   Xmodi = 6
                   data_llaloc.Recordset("hor_llega") = data_llarem.Recordset("hor_llega")
                End If
             Else
                If IsNull(data_llarem.Recordset("hor_llega")) = True Then
                   If IsNull(data_llaloc.Recordset("hor_llega")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("hor_llega") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("hor_llega")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("hor_llega") = data_llarem.Recordset("hor_llega")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("diag")) = False And IsNull(data_llarem.Recordset("diag")) = False Then
                If data_llaloc.Recordset("diag") <> data_llarem.Recordset("diag") Then
                   Xmodi = 6
                   data_llaloc.Recordset("diag") = data_llarem.Recordset("diag")
                End If
             Else
                If IsNull(data_llarem.Recordset("diag")) = True Then
                   If IsNull(data_llaloc.Recordset("diag")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("diag") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("diag")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("diag") = data_llarem.Recordset("diag")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("colormot")) = False And IsNull(data_llarem.Recordset("colormot")) = False Then
                If data_llaloc.Recordset("colormot") <> data_llarem.Recordset("colormot") Then
                   Xmodi = 6
                   data_llaloc.Recordset("colormot") = data_llarem.Recordset("colormot")
                End If
             Else
                If IsNull(data_llarem.Recordset("colormot")) = True Then
                   If IsNull(data_llaloc.Recordset("colormot")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("colormot") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("colormot")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("colormot") = data_llarem.Recordset("colormot")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("codmed")) = False And IsNull(data_llarem.Recordset("codmed")) = False Then
                If data_llaloc.Recordset("codmed") <> data_llarem.Recordset("codmed") Then
                   Xmodi = 6
                   data_llaloc.Recordset("codmed") = data_llarem.Recordset("codmed")
                End If
             Else
                If IsNull(data_llarem.Recordset("codmed")) = True Then
                   If IsNull(data_llaloc.Recordset("codmed")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("codmed") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("codmed")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("codmed") = data_llarem.Recordset("codmed")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("obs")) = False And IsNull(data_llarem.Recordset("obs")) = False Then
                If data_llaloc.Recordset("obs") <> data_llarem.Recordset("obs") Then
                   Xmodi = 6
                   data_llaloc.Recordset("obs") = data_llarem.Recordset("obs")
                End If
             Else
                If IsNull(data_llarem.Recordset("obs")) = True Then
                   If IsNull(data_llaloc.Recordset("obs")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("obs") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("obs")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("obs") = data_llarem.Recordset("obs")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("nommed")) = False And IsNull(data_llarem.Recordset("nommed")) = False Then
                If data_llaloc.Recordset("nommed") <> data_llarem.Recordset("nommed") Then
                   Xmodi = 6
                   data_llaloc.Recordset("nommed") = data_llarem.Recordset("nommed")
                End If
             Else
                If IsNull(data_llarem.Recordset("nommed")) = True Then
                   If IsNull(data_llaloc.Recordset("nommed")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("nommed") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("nommed")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("nommed") = data_llarem.Recordset("nommed")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("trasla")) = False And IsNull(data_llarem.Recordset("trasla")) = False Then
                If data_llaloc.Recordset("trasla") <> data_llarem.Recordset("trasla") Then
                   Xmodi = 6
                   data_llaloc.Recordset("trasla") = data_llarem.Recordset("trasla")
                End If
             Else
                If IsNull(data_llarem.Recordset("trasla")) = True Then
                   If IsNull(data_llaloc.Recordset("trasla")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("trasla") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("trasla")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("trasla") = data_llarem.Recordset("trasla")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("lugar")) = False And IsNull(data_llarem.Recordset("lugar")) = False Then
                If data_llaloc.Recordset("lugar") <> data_llarem.Recordset("lugar") Then
                   Xmodi = 6
                   data_llaloc.Recordset("lugar") = data_llarem.Recordset("lugar")
                End If
             Else
                If IsNull(data_llarem.Recordset("lugar")) = True Then
                   If IsNull(data_llaloc.Recordset("lugar")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("lugar") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("lugar")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("lugar") = data_llarem.Recordset("lugar")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("hsald")) = False And IsNull(data_llarem.Recordset("hsald")) = False Then
                If data_llaloc.Recordset("hsald") <> data_llarem.Recordset("hsald") Then
                   Xmodi = 6
                   data_llaloc.Recordset("hsald") = data_llarem.Recordset("hsald")
                End If
             Else
                If IsNull(data_llarem.Recordset("hsald")) = True Then
                   If IsNull(data_llaloc.Recordset("hsald")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("hsald") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("hsald")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("hsald") = data_llarem.Recordset("hsald")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("hllega")) = False And IsNull(data_llarem.Recordset("hllega")) = False Then
                If data_llaloc.Recordset("hllega") <> data_llarem.Recordset("hllega") Then
                   Xmodi = 6
                   data_llaloc.Recordset("hllega") = data_llarem.Recordset("hllega")
                End If
             Else
                If IsNull(data_llarem.Recordset("hllega")) = True Then
                   If IsNull(data_llaloc.Recordset("hllega")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("hllega") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("hllega")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("hllega") = data_llarem.Recordset("hllega")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("hzona")) = False And IsNull(data_llarem.Recordset("hzona")) = False Then
                If data_llaloc.Recordset("hzona") <> data_llarem.Recordset("hzona") Then
                   Xmodi = 6
                   data_llaloc.Recordset("hzona") = data_llarem.Recordset("hzona")
                End If
             Else
                If IsNull(data_llarem.Recordset("hzona")) = True Then
                   If IsNull(data_llaloc.Recordset("hzona")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("hzona") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("hzona")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("hzona") = data_llarem.Recordset("hzona")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("movtras")) = False And IsNull(data_llarem.Recordset("movtras")) = False Then
                If data_llaloc.Recordset("movtras") <> data_llarem.Recordset("movtras") Then
                   Xmodi = 6
                   data_llaloc.Recordset("movtras") = data_llarem.Recordset("movtras")
                End If
             Else
                If IsNull(data_llarem.Recordset("movtras")) = True Then
                   If IsNull(data_llaloc.Recordset("movtras")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("movtras") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("movtras")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("movtras") = data_llarem.Recordset("movtras")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("totdem")) = False And IsNull(data_llarem.Recordset("totdem")) = False Then
                If data_llaloc.Recordset("totdem") <> data_llarem.Recordset("totdem") Then
                   Xmodi = 6
                   data_llaloc.Recordset("totdem") = data_llarem.Recordset("totdem")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("dcobr")) = False And IsNull(data_llarem.Recordset("dcobr")) = False Then
                If data_llaloc.Recordset("dcobr") <> data_llarem.Recordset("dcobr") Then
                   Xmodi = 6
                   data_llaloc.Recordset("dcobr") = data_llarem.Recordset("dcobr")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("enfer")) = False And IsNull(data_llarem.Recordset("enfer")) = False Then
                If data_llaloc.Recordset("enfer") <> data_llarem.Recordset("enfer") Then
                   Xmodi = 6
                   data_llaloc.Recordset("enfer") = data_llarem.Recordset("enfer")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("motmov")) = False And IsNull(data_llarem.Recordset("motmov")) = False Then
                If data_llaloc.Recordset("motmov") <> data_llarem.Recordset("motmov") Then
                   Xmodi = 6
                   data_llaloc.Recordset("motmov") = data_llarem.Recordset("motmov")
                End If
             Else
                If IsNull(data_llarem.Recordset("motmov")) = True Then
                   If IsNull(data_llaloc.Recordset("motmov")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("motmov") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("motmov")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("motmov") = data_llarem.Recordset("motmov")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("ncobr")) = False And IsNull(data_llarem.Recordset("ncobr")) = False Then
                If data_llaloc.Recordset("ncobr") <> data_llarem.Recordset("ncobr") Then
                   Xmodi = 6
                   data_llaloc.Recordset("ncobr") = data_llarem.Recordset("ncobr")
                End If
             Else
                If IsNull(data_llarem.Recordset("ncobr")) = True Then
                   If IsNull(data_llaloc.Recordset("ncobr")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("ncobr") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("ncobr")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("ncobr") = data_llarem.Recordset("ncobr")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("cancela")) = False And IsNull(data_llarem.Recordset("cancela")) = False Then
                If data_llaloc.Recordset("cancela") <> data_llarem.Recordset("cancela") Then
                   Xmodi = 6
                   data_llaloc.Recordset("cancela") = data_llarem.Recordset("cancela")
                End If
             Else
                If IsNull(data_llarem.Recordset("cancela")) = True Then
                   If IsNull(data_llaloc.Recordset("cancela")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("cancela") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("cancela")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("cancela") = data_llarem.Recordset("cancela")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("motcance")) = False And IsNull(data_llarem.Recordset("motcance")) = False Then
                If data_llaloc.Recordset("motcance") <> data_llarem.Recordset("motcance") Then
                   Xmodi = 6
                   data_llaloc.Recordset("motcance") = data_llarem.Recordset("motcance")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("mm")) = False And IsNull(data_llarem.Recordset("mm")) = False Then
                If data_llaloc.Recordset("mm") <> data_llarem.Recordset("mm") Then
                   Xmodi = 6
                   data_llaloc.Recordset("mm") = data_llarem.Recordset("mm")
                End If
             Else
                If IsNull(data_llarem.Recordset("mm")) = True Then
                   If IsNull(data_llaloc.Recordset("mm")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("mm") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("mm")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("mm") = data_llarem.Recordset("mm")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("thh")) = False And IsNull(data_llarem.Recordset("thh")) = False Then
                If data_llaloc.Recordset("thh") <> data_llarem.Recordset("thh") Then
                   Xmodi = 6
                   data_llaloc.Recordset("thh") = data_llarem.Recordset("thh")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("tmm")) = False And IsNull(data_llarem.Recordset("tmm")) = False Then
                If data_llaloc.Recordset("tmm") <> data_llarem.Recordset("tmm") Then
                   Xmodi = 6
                   data_llaloc.Recordset("tmm") = data_llarem.Recordset("tmm")
                End If
             Else
                If IsNull(data_llarem.Recordset("tmm")) = True Then
                   If IsNull(data_llaloc.Recordset("tmm")) = True Then
                   Else
                      Xmodi = 6
                      data_llaloc.Recordset("tmm") = Null
                   End If
                Else
                   If IsNull(data_llaloc.Recordset("tmm")) = True Then
                      Xmodi = 6
                      data_llaloc.Recordset("tmm") = data_llarem.Recordset("tmm")
                   End If
                End If
             End If
             If IsNull(data_llaloc.Recordset("pasado")) = False And IsNull(data_llarem.Recordset("pasado")) = False Then
                If data_llaloc.Recordset("pasado") <> data_llarem.Recordset("pasado") Then
                   Xmodi = 6
                   data_llaloc.Recordset("pasado") = data_llarem.Recordset("pasado")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("ano")) = False And IsNull(data_llarem.Recordset("ano")) = False Then
                If data_llaloc.Recordset("ano") <> data_llarem.Recordset("ano") Then
                   Xmodi = 6
                   data_llaloc.Recordset("ano") = data_llarem.Recordset("ano")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("mes")) = False And IsNull(data_llarem.Recordset("mes")) = False Then
                If data_llaloc.Recordset("mes") <> data_llarem.Recordset("mes") Then
                   Xmodi = 6
                   data_llaloc.Recordset("mes") = data_llarem.Recordset("mes")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("timsi")) = False And IsNull(data_llarem.Recordset("timsi")) = False Then
                If data_llaloc.Recordset("timsi") <> data_llarem.Recordset("timsi") Then
                   Xmodi = 6
                   data_llaloc.Recordset("timsi") = data_llarem.Recordset("timsi")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("hor_cance")) = False And IsNull(data_llarem.Recordset("hor_cance")) = False Then
                If data_llaloc.Recordset("hor_cance") <> data_llarem.Recordset("hor_cance") Then
                   Xmodi = 6
                   data_llaloc.Recordset("hor_cance") = data_llarem.Recordset("hor_cance")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("movil_rea")) = False And IsNull(data_llarem.Recordset("movil_rea")) = False Then
                If data_llaloc.Recordset("movil_rea") <> data_llarem.Recordset("movil_rea") Then
                   Xmodi = 6
                   data_llaloc.Recordset("movil_rea") = data_llarem.Recordset("movil_rea")
                End If
             Else
             End If
             If IsNull(data_llaloc.Recordset("hh")) = False And IsNull(data_llarem.Recordset("hh")) = False Then
                If data_llaloc.Recordset("hh") <> data_llarem.Recordset("hh") Then
                   Xmodi = 6
                   data_llaloc.Recordset("hh") = data_llarem.Recordset("hh")
                End If
             Else
             End If
             If Xmodi = 6 Then
                data_llaloc.Recordset.Update
             End If
             Xmodi = 0
             data_llarem2.RecordSource = "Select * from resplla where nro =" & data_llarem.Recordset("nrolla")
             data_llarem2.Refresh
             If data_llarem2.Recordset.RecordCount > 0 Then
                data_llaloc2.RecordSource = "select * from resplla where nro =" & data_llarem.Recordset("nrolla")
                data_llaloc2.Refresh
                If data_llaloc2.Recordset.RecordCount > 0 Then
                   data_llaloc2.Recordset.MoveFirst
                   data_llaloc2.Recordset.Edit
                   If IsNull(data_llaloc2.Recordset("fecha")) = False And IsNull(data_llarem2.Recordset("fecha")) = False Then
                      If data_llaloc2.Recordset("fecha") <> data_llarem2.Recordset("fecha") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("fecha") = data_llarem2.Recordset("fecha")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("telef")) = False And IsNull(data_llarem2.Recordset("telef")) = False Then
                      If data_llaloc2.Recordset("telef") <> data_llarem2.Recordset("telef") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("telef") = data_llarem2.Recordset("telef")
                      End If
                   Else
                      If IsNull(data_llarem2.Recordset("telef")) = True Then
                         If IsNull(data_llaloc2.Recordset("telef")) = True Then
                         Else
                            Xmodi = 6
                            data_llaloc2.Recordset("telef") = Null
                         End If
                      Else
                         If IsNull(data_llaloc2.Recordset("telef")) = True Then
                            Xmodi = 6
                            data_llaloc2.Recordset("telef") = data_llarem2.Recordset("telef")
                         End If
                      End If
                   End If
                   If IsNull(data_llaloc2.Recordset("mes")) = False And IsNull(data_llarem2.Recordset("mes")) = False Then
                      If data_llaloc2.Recordset("mes") <> data_llarem2.Recordset("mes") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("mes") = data_llarem2.Recordset("mes")
                      End If
                   Else
                      If IsNull(data_llarem2.Recordset("mes")) = True Then
                         If IsNull(data_llaloc2.Recordset("mes")) = True Then
                         Else
                            Xmodi = 6
                            data_llaloc2.Recordset("mes") = Null
                         End If
                      Else
                         If IsNull(data_llaloc2.Recordset("mes")) = True Then
                            Xmodi = 6
                            data_llaloc2.Recordset("mes") = data_llarem2.Recordset("mes")
                         End If
                      End If
                   End If
                   If IsNull(data_llaloc2.Recordset("fec_llega")) = False And IsNull(data_llarem2.Recordset("fec_llega")) = False Then
                      If data_llaloc2.Recordset("fec_llega") <> data_llarem2.Recordset("fec_llega") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("fec_llega") = data_llarem2.Recordset("fec_llega")
                      End If
                   Else
                      If IsNull(data_llarem2.Recordset("fec_llega")) = True Then
                         If IsNull(data_llaloc2.Recordset("fec_llega")) = True Then
                         Else
                            Xmodi = 6
                            data_llaloc2.Recordset("fec_llega") = Null
                         End If
                      Else
                         If IsNull(data_llaloc2.Recordset("fec_llega")) = True Then
                            Xmodi = 6
                            data_llaloc2.Recordset("fec_llega") = data_llarem2.Recordset("fec_llega")
                         End If
                      End If
                   End If
                   If IsNull(data_llaloc2.Recordset("hor_llega")) = False And IsNull(data_llarem2.Recordset("hor_llega")) = False Then
                      If data_llaloc2.Recordset("hor_llega") <> data_llarem2.Recordset("hor_llega") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("hor_llega") = data_llarem2.Recordset("hor_llega")
                      End If
                   Else
                      If IsNull(data_llarem2.Recordset("hor_llega")) = True Then
                         If IsNull(data_llaloc2.Recordset("hor_llega")) = True Then
                         Else
                            Xmodi = 6
                            data_llaloc2.Recordset("hor_llega") = Null
                         End If
                      Else
                         If IsNull(data_llaloc2.Recordset("hor_llega")) = True Then
                            Xmodi = 6
                            data_llaloc2.Recordset("hor_llega") = data_llarem2.Recordset("hor_llega")
                         End If
                      End If
                   End If
                   If IsNull(data_llaloc2.Recordset("matric")) = False And IsNull(data_llarem2.Recordset("matric")) = False Then
                      If data_llaloc2.Recordset("matric") <> data_llarem2.Recordset("matric") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("matric") = data_llarem2.Recordset("matric")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("edad")) = False And IsNull(data_llarem2.Recordset("edad")) = False Then
                      If data_llaloc2.Recordset("edad") <> data_llarem2.Recordset("edad") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("edad") = data_llarem2.Recordset("edad")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("direcc")) = False And IsNull(data_llarem2.Recordset("direcc")) = False Then
                      If data_llaloc2.Recordset("direcc") <> data_llarem2.Recordset("direcc") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("direcc") = data_llarem2.Recordset("direcc")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("pasado")) = False And IsNull(data_llarem2.Recordset("pasado")) = False Then
                      If data_llaloc2.Recordset("pasado") <> data_llarem2.Recordset("pasado") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("pasado") = data_llarem2.Recordset("pasado")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("direcc")) = False And IsNull(data_llarem2.Recordset("direcc")) = False Then
                      If data_llaloc2.Recordset("direcc") <> data_llarem2.Recordset("direcc") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("direcc") = data_llarem2.Recordset("direcc")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("movilpas")) = False And IsNull(data_llarem2.Recordset("movilpas")) = False Then
                      If data_llaloc2.Recordset("movilpas") <> data_llarem2.Recordset("movilpas") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("movilpas") = data_llarem2.Recordset("movilpas")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("nommed")) = False And IsNull(data_llarem2.Recordset("nommed")) = False Then
                      If data_llaloc2.Recordset("nommed") <> data_llarem2.Recordset("nommed") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("nommed") = data_llarem2.Recordset("nommed")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("movil_rea")) = False And IsNull(data_llarem2.Recordset("movil_rea")) = False Then
                      If data_llaloc2.Recordset("movil_rea") <> data_llarem2.Recordset("movil_rea") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("movil_rea") = data_llarem2.Recordset("movil_rea")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("hora")) = False And IsNull(data_llarem2.Recordset("hora")) = False Then
                      If data_llaloc2.Recordset("hora") <> data_llarem2.Recordset("hora") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("hora") = data_llarem2.Recordset("hora")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("usuario")) = False And IsNull(data_llarem2.Recordset("usuario")) = False Then
                      If data_llaloc2.Recordset("usuario") <> data_llarem2.Recordset("usuario") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("usuario") = data_llarem2.Recordset("usuario")
                      End If
                   Else
                   End If
                   If Xmodi = 6 Then
                      data_llaloc2.Recordset.Update
                   End If
                   Xmodi = 0
                Else
                   data_llaloc2.Recordset.AddNew
                   data_llaloc2.Recordset("nro") = data_llarem2.Recordset("nro")
                   data_llaloc2.Recordset("fecha") = data_llarem2.Recordset("fecha")
                   data_llaloc2.Recordset("telef") = data_llarem2.Recordset("telef")
                   data_llaloc2.Recordset("mes") = data_llarem2.Recordset("mes")
                   data_llaloc2.Recordset("fec_llega") = data_llarem2.Recordset("fec_llega")
                   data_llaloc2.Recordset("hor_llega") = data_llarem2.Recordset("hor_llega")
                   data_llaloc2.Recordset("matric") = data_llarem2.Recordset("matric")
                   data_llaloc2.Recordset("edad") = data_llarem2.Recordset("edad")
                   data_llaloc2.Recordset("direcc") = data_llarem2.Recordset("direcc")
                   data_llaloc2.Recordset("pasado") = data_llarem2.Recordset("pasado")
                   data_llaloc2.Recordset("direcc") = data_llarem2.Recordset("direcc")
                   data_llaloc2.Recordset("movilpas") = data_llarem2.Recordset("movilpas")
                   data_llaloc2.Recordset("nommed") = data_llarem2.Recordset("nommed")
                   data_llaloc2.Recordset("movil_rea") = data_llarem2.Recordset("movil_rea")
                   data_llaloc2.Recordset("hora") = data_llarem2.Recordset("hora")
                   data_llaloc2.Recordset("usuario") = data_llarem2.Recordset("usuario")
                   data_llaloc2.Recordset.Update
                End If
             End If
          Else
             data_llaloc.Recordset.AddNew
             If IsNull(data_llarem.Recordset("nrolla")) = False Then
                data_llaloc.Recordset("nrolla") = data_llarem.Recordset("nrolla")
             End If
             If IsNull(data_llarem.Recordset("nro")) = False Then
                data_llaloc.Recordset("nro") = data_llarem.Recordset("nro")
             End If
             If IsNull(data_llarem.Recordset("fecha")) = False Then
                data_llaloc.Recordset("fecha") = data_llarem.Recordset("fecha")
             End If
             If IsNull(data_llarem.Recordset("hora")) = False Then
                data_llaloc.Recordset("hora") = data_llarem.Recordset("hora")
             End If
             If IsNull(data_llarem.Recordset("usuario")) = False Then
                data_llaloc.Recordset("usuario") = data_llarem.Recordset("usuario")
             End If
             If IsNull(data_llarem.Recordset("matric")) = False Then
                data_llaloc.Recordset("matric") = data_llarem.Recordset("matric")
             End If
             If IsNull(data_llarem.Recordset("nombre")) = False Then
                data_llaloc.Recordset("nombre") = data_llarem.Recordset("nombre")
             End If
             If IsNull(data_llarem.Recordset("edad")) = False Then
                data_llaloc.Recordset("edad") = data_llarem.Recordset("edad")
             End If
             If IsNull(data_llarem.Recordset("unied")) = False Then
                data_llaloc.Recordset("unied") = data_llarem.Recordset("unied")
             End If
             If IsNull(data_llarem.Recordset("categ")) = False Then
                data_llaloc.Recordset("categ") = data_llarem.Recordset("categ")
                data_llaloc.Recordset("nomcat") = data_llarem.Recordset("nomcat")
             End If
             If IsNull(data_llarem.Recordset("ci")) = False Then
                data_llaloc.Recordset("ci") = data_llarem.Recordset("ci")
             End If
             If IsNull(data_llarem.Recordset("telef")) = False Then
                data_llaloc.Recordset("telef") = data_llarem.Recordset("telef")
             End If
             If IsNull(data_llarem.Recordset("codzon")) = False Then
                data_llaloc.Recordset("codzon") = data_llarem.Recordset("codzon")
             End If
             If IsNull(data_llarem.Recordset("base")) = False Then
                data_llaloc.Recordset("base") = data_llarem.Recordset("base")
             End If
             If IsNull(data_llarem.Recordset("referen")) = False Then
                data_llaloc.Recordset("referen") = data_llarem.Recordset("referen")
             End If
             If IsNull(data_llarem.Recordset("realiza")) = False Then
                data_llaloc.Recordset("realiza") = data_llarem.Recordset("realiza")
             End If
             If IsNull(data_llarem.Recordset("motcon")) = False Then
                data_llaloc.Recordset("motcon") = data_llarem.Recordset("motcon")
             End If
             If IsNull(data_llarem.Recordset("obsmot")) = False Then
                data_llaloc.Recordset("obsmot") = data_llarem.Recordset("obsmot")
             End If
             If IsNull(data_llarem.Recordset("codmot")) = False Then
                data_llaloc.Recordset("codmot") = data_llarem.Recordset("codmot")
                data_llaloc.Recordset("descol") = data_llarem.Recordset("descol")
             End If
             If IsNull(data_llarem.Recordset("movilpas")) = False Then
                data_llaloc.Recordset("movilpas") = data_llarem.Recordset("movilpas")
             End If
             If IsNull(data_llarem.Recordset("pend")) = False Then
                data_llaloc.Recordset("pend") = data_llarem.Recordset("pend")
             End If
             If IsNull(data_llarem.Recordset("timdes")) = False Then
                data_llaloc.Recordset("timdes") = data_llarem.Recordset("timdes")
             End If
             If IsNull(data_llarem.Recordset("fec_rea")) = False Then
                data_llaloc.Recordset("fec_rea") = data_llarem.Recordset("fec_rea")
             End If
             If IsNull(data_llarem.Recordset("hor_rea")) = False Then
                data_llaloc.Recordset("hor_rea") = data_llarem.Recordset("hor_rea")
             End If
             If IsNull(data_llarem.Recordset("fecpas")) = False Then
                data_llaloc.Recordset("fecpas") = data_llarem.Recordset("fecpas")
             End If
             If IsNull(data_llarem.Recordset("horpas")) = False Then
                data_llaloc.Recordset("horpas") = data_llarem.Recordset("horpas")
             End If
             If IsNull(data_llarem.Recordset("fecsali")) = False Then
                data_llaloc.Recordset("fecsali") = data_llarem.Recordset("fecsali")
             End If
             If IsNull(data_llarem.Recordset("horsali")) = False Then
                data_llaloc.Recordset("horsali") = data_llarem.Recordset("horsali")
             End If
             If IsNull(data_llarem.Recordset("fec_llega")) = False Then
                data_llaloc.Recordset("fec_llega") = data_llarem.Recordset("fec_llega")
             End If
             If IsNull(data_llarem.Recordset("hor_llega")) = False Then
                data_llaloc.Recordset("hor_llega") = data_llarem.Recordset("hor_llega")
             End If
             If IsNull(data_llarem.Recordset("diag")) = False Then
                data_llaloc.Recordset("diag") = data_llarem.Recordset("diag")
             End If
             If IsNull(data_llarem.Recordset("colormot")) = False Then
                data_llaloc.Recordset("colormot") = data_llarem.Recordset("colormot")
             End If
             If IsNull(data_llarem.Recordset("codmed")) = False Then
                data_llaloc.Recordset("codmed") = data_llarem.Recordset("codmed")
             End If
             If IsNull(data_llarem.Recordset("obs")) = False Then
                data_llaloc.Recordset("obs") = data_llarem.Recordset("obs")
             End If
             If IsNull(data_llarem.Recordset("nommed")) = False Then
                data_llaloc.Recordset("nommed") = data_llarem.Recordset("nommed")
             End If
             If IsNull(data_llarem.Recordset("trasla")) = False Then
                data_llaloc.Recordset("trasla") = data_llarem.Recordset("trasla")
             End If
             If IsNull(data_llarem.Recordset("lugar")) = False Then
                data_llaloc.Recordset("lugar") = data_llarem.Recordset("lugar")
             End If
             If IsNull(data_llarem.Recordset("hsald")) = False Then
                data_llaloc.Recordset("hsald") = data_llarem.Recordset("hsald")
             End If
             If IsNull(data_llarem.Recordset("hllega")) = False Then
                data_llaloc.Recordset("hllega") = data_llarem.Recordset("hllega")
             End If
             If IsNull(data_llarem.Recordset("hzona")) = False Then
                data_llaloc.Recordset("hzona") = data_llarem.Recordset("hzona")
             End If
             If IsNull(data_llarem.Recordset("movtras")) = False Then
                data_llaloc.Recordset("movtras") = data_llarem.Recordset("movtras")
             End If
             If IsNull(data_llarem.Recordset("totdem")) = False Then
                data_llaloc.Recordset("totdem") = data_llarem.Recordset("totdem")
             End If
             If IsNull(data_llarem.Recordset("dcobr")) = False Then
                data_llaloc.Recordset("dcobr") = data_llarem.Recordset("dcobr")
             End If
             If IsNull(data_llarem.Recordset("enfer")) = False Then
                data_llaloc.Recordset("enfer") = data_llarem.Recordset("enfer")
             End If
             If IsNull(data_llarem.Recordset("motmov")) = False Then
                data_llaloc.Recordset("motmov") = data_llarem.Recordset("motmov")
             End If
             If IsNull(data_llarem.Recordset("ncobr")) = False Then
                data_llaloc.Recordset("ncobr") = data_llarem.Recordset("ncobr")
             End If
             If IsNull(data_llarem.Recordset("cancela")) = False Then
                data_llaloc.Recordset("cancela") = data_llarem.Recordset("cancela")
             End If
             If IsNull(data_llarem.Recordset("motcance")) = False Then
                data_llaloc.Recordset("motcance") = data_llarem.Recordset("motcance")
             End If
             If IsNull(data_llarem.Recordset("obs")) = False Then
                data_llaloc.Recordset("obs") = data_llarem.Recordset("obs")
             End If
             If IsNull(data_llarem.Recordset("mm")) = False Then
                data_llaloc.Recordset("mm") = data_llarem.Recordset("mm")
             End If
             If IsNull(data_llarem.Recordset("thh")) = False Then
                data_llaloc.Recordset("thh") = data_llarem.Recordset("thh")
             End If
             If IsNull(data_llarem.Recordset("tmm")) = False Then
                data_llaloc.Recordset("tmm") = data_llarem.Recordset("tmm")
             End If
             If IsNull(data_llarem.Recordset("pasado")) = False Then
                data_llaloc.Recordset("pasado") = data_llarem.Recordset("pasado")
             End If
             If IsNull(data_llarem.Recordset("ano")) = False Then
                data_llaloc.Recordset("ano") = data_llarem.Recordset("ano")
             End If
             If IsNull(data_llarem.Recordset("mes")) = False Then
                data_llaloc.Recordset("mes") = data_llarem.Recordset("mes")
             End If
             If IsNull(data_llarem.Recordset("timsi")) = False Then
                data_llaloc.Recordset("timsi") = data_llarem.Recordset("timsi")
             End If
             If IsNull(data_llarem.Recordset("hor_cance")) = False Then
                data_llaloc.Recordset("hor_cance") = data_llarem.Recordset("hor_cance")
             End If
             If IsNull(data_llarem.Recordset("movil_rea")) = False Then
                data_llaloc.Recordset("movil_rea") = data_llarem.Recordset("movil_rea")
             End If
             If IsNull(data_llarem.Recordset("hh")) = False Then
                data_llaloc.Recordset("hh") = data_llarem.Recordset("hh")
             End If
             data_llaloc.Recordset.Update
             Xmodi = 0
             data_llarem2.RecordSource = "Select * from resplla where nro =" & data_llarem.Recordset("nrolla")
             data_llarem2.Refresh
             If data_llarem2.Recordset.RecordCount > 0 Then
                data_llaloc2.RecordSource = "select * from resplla where nro =" & data_llarem.Recordset("nrolla")
                data_llaloc2.Refresh
                If data_llaloc2.Recordset.RecordCount > 0 Then
                   data_llaloc2.Recordset.MoveFirst
                   data_llaloc2.Recordset.Edit
                   If IsNull(data_llaloc2.Recordset("fecha")) = False And IsNull(data_llarem2.Recordset("fecha")) = False Then
                      If data_llaloc2.Recordset("fecha") <> data_llarem2.Recordset("fecha") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("fecha") = data_llarem2.Recordset("fecha")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("telef")) = False And IsNull(data_llarem2.Recordset("telef")) = False Then
                      If data_llaloc2.Recordset("telef") <> data_llarem2.Recordset("telef") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("telef") = data_llarem2.Recordset("telef")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("mes")) = False And IsNull(data_llarem2.Recordset("mes")) = False Then
                      If data_llaloc2.Recordset("mes") <> data_llarem2.Recordset("mes") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("mes") = data_llarem2.Recordset("mes")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("fec_llega")) = False And IsNull(data_llarem2.Recordset("fec_llega")) = False Then
                      If data_llaloc2.Recordset("fec_llega") <> data_llarem2.Recordset("fec_llega") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("fec_llega") = data_llarem2.Recordset("fec_llega")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("hor_llega")) = False And IsNull(data_llarem2.Recordset("hor_llega")) = False Then
                      If data_llaloc2.Recordset("hor_llega") <> data_llarem2.Recordset("hor_llega") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("hor_llega") = data_llarem2.Recordset("hor_llega")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("matric")) = False And IsNull(data_llarem2.Recordset("matric")) = False Then
                      If data_llaloc2.Recordset("matric") <> data_llarem2.Recordset("matric") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("matric") = data_llarem2.Recordset("matric")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("edad")) = False And IsNull(data_llarem2.Recordset("edad")) = False Then
                      If data_llaloc2.Recordset("edad") <> data_llarem2.Recordset("edad") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("edad") = data_llarem2.Recordset("edad")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("direcc")) = False And IsNull(data_llarem2.Recordset("direcc")) = False Then
                      If data_llaloc2.Recordset("direcc") <> data_llarem2.Recordset("direcc") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("direcc") = data_llarem2.Recordset("direcc")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("pasado")) = False And IsNull(data_llarem2.Recordset("pasado")) = False Then
                      If data_llaloc2.Recordset("pasado") <> data_llarem2.Recordset("pasado") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("pasado") = data_llarem2.Recordset("pasado")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("direcc")) = False And IsNull(data_llarem2.Recordset("direcc")) = False Then
                      If data_llaloc2.Recordset("direcc") <> data_llarem2.Recordset("direcc") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("direcc") = data_llarem2.Recordset("direcc")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("movilpas")) = False And IsNull(data_llarem2.Recordset("movilpas")) = False Then
                      If data_llaloc2.Recordset("movilpas") <> data_llarem2.Recordset("movilpas") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("movilpas") = data_llarem2.Recordset("movilpas")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("nommed")) = False And IsNull(data_llarem2.Recordset("nommed")) = False Then
                      If data_llaloc2.Recordset("nommed") <> data_llarem2.Recordset("nommed") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("nommed") = data_llarem2.Recordset("nommed")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("movil_rea")) = False And IsNull(data_llarem2.Recordset("movil_rea")) = False Then
                      If data_llaloc2.Recordset("movil_rea") <> data_llarem2.Recordset("movil_rea") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("movil_rea") = data_llarem2.Recordset("movil_rea")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("hora")) = False And IsNull(data_llarem2.Recordset("hora")) = False Then
                      If data_llaloc2.Recordset("hora") <> data_llarem2.Recordset("hora") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("hora") = data_llarem2.Recordset("hora")
                      End If
                   Else
                   End If
                   If IsNull(data_llaloc2.Recordset("usuario")) = False And IsNull(data_llarem2.Recordset("usuario")) = False Then
                      If data_llaloc2.Recordset("usuario") <> data_llarem2.Recordset("usuario") Then
                         Xmodi = 6
                         data_llaloc2.Recordset("usuario") = data_llarem2.Recordset("usuario")
                      End If
                   Else
                   End If
                   If Xmodi = 6 Then
                      data_llaloc2.Recordset.Update
                   End If
                   Xmodi = 0
                Else
                   data_llaloc2.Recordset.AddNew
                   data_llaloc2.Recordset("nro") = data_llarem2.Recordset("nro")
                   data_llaloc2.Recordset("fecha") = data_llarem2.Recordset("fecha")
                   data_llaloc2.Recordset("telef") = data_llarem2.Recordset("telef")
                   data_llaloc2.Recordset("mes") = data_llarem2.Recordset("mes")
                   data_llaloc2.Recordset("fec_llega") = data_llarem2.Recordset("fec_llega")
                   data_llaloc2.Recordset("hor_llega") = data_llarem2.Recordset("hor_llega")
                   data_llaloc2.Recordset("matric") = data_llarem2.Recordset("matric")
                   data_llaloc2.Recordset("edad") = data_llarem2.Recordset("edad")
                   data_llaloc2.Recordset("direcc") = data_llarem2.Recordset("direcc")
                   data_llaloc2.Recordset("pasado") = data_llarem2.Recordset("pasado")
                   data_llaloc2.Recordset("direcc") = data_llarem2.Recordset("direcc")
                   data_llaloc2.Recordset("movilpas") = data_llarem2.Recordset("movilpas")
                   data_llaloc2.Recordset("nommed") = data_llarem2.Recordset("nommed")
                   data_llaloc2.Recordset("movil_rea") = data_llarem2.Recordset("movil_rea")
                   data_llaloc2.Recordset("hora") = data_llarem2.Recordset("hora")
                   data_llaloc2.Recordset("usuario") = data_llarem2.Recordset("usuario")
                   data_llaloc2.Recordset.Update
                End If
             End If
          End If
          data_llarem.Recordset.MoveNext
       Loop
       
'       MsgBox "Termina"
       XX = 0
    End If
Else

End If
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
data_llarem.Connect = "odbc;dsn=sappnew;"
data_llarem2.Connect = "odbc;dsn=sappnew;"

data_llaloc.Connect = "odbc;dsn=sapploca;"
'data_llaloc.RecordSource = "llamado"
'data_llaloc.Refresh

data_llaloc2.Connect = "odbc;dsn=sapploca;"

End Sub

Private Sub Timer1_Timer()

XX = XX + 1
Command1_Click

End Sub
