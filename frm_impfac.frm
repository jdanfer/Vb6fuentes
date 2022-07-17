VERSION 5.00
Begin VB.Form frm_impfac 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Imprime"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_fac 
      Caption         =   "data_fac"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LINEAS"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton btn_no 
      BackColor       =   &H00C0E0FF&
      Caption         =   "NO IMPRIME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      Picture         =   "frm_impfac.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton btn_si 
      BackColor       =   &H00C0E0FF&
      Caption         =   "IMPRIME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Picture         =   "frm_impfac.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frm_impfac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_no_Click()
frmabm.btn_fact.Enabled = True

Unload Me

End Sub

Private Sub btn_si_Click()
Dim Xcontar As Integer
Dim XTotIV, XsubTot, Xsumo As Double
Dim Xtotal, Xsubt, XtoiV, XimpSub As Double
frmabm.btn_fact.Enabled = True

Xsubt = 0
XtoiV = 0
XimpSub = 0
Xcontar = 0
Xtotal = 0
Xcontar = 1
'MsgBox "Verifique impresora", vbCritical, "Mensaje de ERROR"
If data_fac.Recordset.RecordCount > 0 Then
   data_fac.Recordset.MoveFirst
   Do While Not data_fac.Recordset.EOF
      Xsumo = data_fac.Recordset("tot_lin") - data_fac.Recordset("imp_iva")
      XsubTot = XsubTot + Xsumo
      XTotIV = XTotIV + data_fac.Recordset("imp_iva")
      Xtotal = Xtotal + data_fac.Recordset("tot_lin")
      data_fac.Recordset.MoveNext
   Loop
   data_fac.Recordset.MoveFirst
   Do While Not data_fac.Recordset.EOF
      data_fac.Recordset.Edit
      data_fac.Recordset("costo") = XsubTot
      data_fac.Recordset("costo_prod") = XTotIV
      data_fac.Recordset.Update
      data_fac.Recordset.MoveNext
   Loop
   XsubTot = 0
   XTotIV = 0
   data_fac.Recordset.MoveFirst
    Printer.ScaleWidth = 2300
    Printer.ScaleHeight = 1000
'    Printer.FontName = "MS Sans Serif"
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 10
'-----------
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 550
    If IsNull(data_fac.Recordset("ruc")) = False Then
       Printer.Print data_fac.Recordset("ruc");
       Printer.Print "         X";
    Else
       Printer.Print "";
    End If
    Printer.CurrentX = 1600
    Printer.Print Format(data_fac.Recordset("fecha"), "dd/mm/yyyy");
    Printer.Print "    ";
    Printer.Print Format(data_fac.Recordset("hora"), "HH:mm");
    Printer.CurrentX = 2000
    Printer.Print Format(data_fac.Recordset("fecha"), "dd/mm/yyyy");
    Printer.Print "  ";
    Printer.Print Format(data_fac.Recordset("hora"), "HH:mm")
    Printer.Print ""
'        Printer.Print ""
    Printer.CurrentX = 2000
    If IsNull(data_fac.Recordset("tipo")) = True Then
       Printer.Print " ";
    Else
       If data_fac.Recordset("tipo") = "RECIBO" Then
          If data_fac.Recordset("cod_prod") = 60108 Or data_fac.Recordset("cod_prod") = 60103 Then
             Printer.Print "P/CTA ORDEN";
          Else
             Printer.Print data_fac.Recordset("tipo");
          End If
       Else
          Printer.Print data_fac.Recordset("tipo");
       End If
    End If
    Printer.Print "  " & data_fac.Recordset("factura")
    Printer.CurrentX = 1700
    If IsNull(data_fac.Recordset("tipo")) = True Then
       Printer.Print " "
    Else
       If data_fac.Recordset("tipo") = "RECIBO" Then
          If data_fac.Recordset("cod_prod") = 60108 Or data_fac.Recordset("cod_prod") = 60103 Then
             Printer.Print "P/CTA ORDEN"
          Else
             Printer.Print data_fac.Recordset("tipo")
          End If
       Else
          Printer.Print data_fac.Recordset("tipo")
       End If
    End If
    Printer.CurrentX = 2000
    Printer.Print data_fac.Recordset("cod_cli");
    Printer.Print "    " & data_fac.Recordset("convenio")
    Printer.Print data_fac.Recordset("nom_cli");
    Printer.CurrentX = 1000
    Printer.Print data_fac.Recordset("convenio");
    Printer.CurrentX = 1200
    Printer.Print data_fac.Recordset("cod_cli")
    Printer.CurrentX = 1600
    Printer.Print data_fac.Recordset("factura");
    Printer.CurrentX = 2200
    Printer.Print Format(Xsubt, "Standard")
    Printer.CurrentX = 2200
    Printer.Print Format(XtoiV, "Standard")
    Printer.Print data_fac.Recordset("fecha");
    Printer.Print " ";
    Printer.Print data_fac.Recordset("hora");
    Printer.Print " ";
    Printer.Print data_fac.Recordset("operador");
    Printer.Print "           ";
    Printer.Print data_fac.Recordset("tipo");
    Printer.CurrentX = 850
    Printer.Print "  ";
    Printer.Print data_fac.Recordset("factura");
    Printer.Print "  ";
    Printer.CurrentX = 1150
    If IsNull(data_fac.Recordset("nro_med_a")) = True Then
       Printer.Print "S/D";
    Else
       Printer.Print data_fac.Recordset("nro_med_a");
    End If
    Printer.CurrentX = 1700
    Printer.Print data_fac.Recordset("cod_cli");
    Printer.CurrentX = 2200
    Printer.Print Format(Xtotal, "Standard")
    Printer.Print ""
    
    Printer.Print data_fac.Recordset("cod_prod");
    Printer.Print "     ";
    If IsNull(data_fac.Recordset("mes_paga")) = True Then
       Printer.Print data_fac.Recordset("nom_prod");
    Else
       Printer.Print data_fac.Recordset("nom_prod");
       Printer.Print " MES: " + Trim(Str(data_fac.Recordset("mes_paga"))) + "/" + Trim(Str(data_fac.Recordset("ano_paga")));
    End If
    Printer.CurrentX = 1350
    Xtotal = 0
    Printer.Print Format(data_fac.Recordset("tot_lin"), "Standard");
    Printer.CurrentX = 1750
    Printer.Print data_fac.Recordset("convenio")
    Xtotal = Xtotal + data_fac.Recordset("tot_lin")
    data_fac.Recordset.MoveNext
    Do While Not data_fac.Recordset.RecordCount = Xcontar
       Printer.Print data_fac.Recordset("cod_prod");
       Printer.Print "     ";
       If IsNull(data_fac.Recordset("mes_paga")) = True Then
          Printer.Print data_fac.Recordset("nom_prod");
       Else
          Printer.Print data_fac.Recordset("nom_prod");
          Printer.Print " MES: " + Trim(Str(data_fac.Recordset("mes_paga"))) + "/" + Trim(Str(data_fac.Recordset("ano_paga")));
       End If
       Printer.CurrentX = 1350
       Xcontar = Xcontar + 1
       Xtotal = Xtotal + data_fac.Recordset("tot_lin")
       If Xcontar = 2 Then
          Printer.Print Format(data_fac.Recordset("tot_lin"), "Standard")
       Else
          If Xcontar = 3 Then
             Printer.Print Format(data_fac.Recordset("tot_lin"), "Standard")
          Else
             If Xcontar = 4 Then
                Printer.Print Format(data_fac.Recordset("tot_lin"), "Standard")
             Else
                If Xcontar = 5 Then
                   Printer.Print Format(data_fac.Recordset("tot_lin"), "Standard");
                   Printer.CurrentX = 1600
                   If IsNull(data_fac.Recordset("nom_med_a")) = True Then
                      Printer.Print "S/D"
                   Else
                      Printer.Print data_fac.Recordset("nom_med_a")
                   End If
                   Printer.Print ""
                   Printer.CurrentX = 1750
                   Printer.Print Format(Xsubt, "Standard")
                End If
             End If
          End If
       End If
'       Xcontar = Xcontar + 1
       data_fac.Recordset.MoveNext
    Loop
    data_fac.Recordset.MoveFirst
    XTotIV = 0
    Do While Not data_fac.Recordset.EOF
       XTotIV = XTotIV + data_fac.Recordset("imp_iva")
       data_fac.Recordset.MoveNext
    Loop
    data_fac.Recordset.MoveFirst
    XsubTot = Xtotal - XTotIV
    If Xcontar = 1 Then
       Printer.Print ""
       Printer.Print ""
       Printer.CurrentX = 1600
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D"
       Else
          Printer.Print data_fac.Recordset("nro_med_a")
       End If
       Printer.Print ""
       Printer.Print ""
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xsubt, "Standard")
       Printer.Print ""
       Printer.CurrentX = 50
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D";
       Else
          Printer.Print data_fac.Recordset("nro_med_a");
       End If
       Printer.CurrentX = 750
       Printer.Print Format(Xsubt, "Standard");
       Printer.CurrentX = 1150
       Printer.Print Format(XTotIV, "Standard");
       Printer.CurrentX = 1350
       Printer.Print Format(Xtotal, "Standard");
       Printer.CurrentX = 1750
       Printer.Print Format(XTotIV, "Standard")
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xtotal, "Standard")
    End If
    If Xcontar = 2 Then
       Printer.Print ""
       Printer.CurrentX = 1600
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D"
       Else
          Printer.Print data_fac.Recordset("nro_med_a")
       End If
       Printer.Print ""
       Printer.Print ""
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xsubt, "Standard")
       Printer.Print ""
       Printer.CurrentX = 50
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D";
       Else
          Printer.Print data_fac.Recordset("nro_med_a");
       End If
       Printer.CurrentX = 750
       Printer.Print Format(Xsubt, "Standard");
       Printer.CurrentX = 1150
       Printer.Print Format(XTotIV, "Standard");
       Printer.CurrentX = 1350
       Printer.Print Format(Xtotal, "Standard");
       Printer.CurrentX = 1750
       Printer.Print Format(XTotIV, "Standard")
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xtotal, "Standard")
    
    End If
    If Xcontar = 3 Then
       Printer.Print ""
       Printer.CurrentX = 1600
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D"
       Else
          Printer.Print data_fac.Recordset("nro_med_a")
       End If
       Printer.Print ""
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xsubt, "Standard")
       Printer.Print ""
       Printer.CurrentX = 50
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D";
       Else
          Printer.Print data_fac.Recordset("nro_med_a");
       End If
'           Printer.Print ""
       Printer.CurrentX = 750
       Printer.Print Format(Xsubt, "Standard");
       Printer.CurrentX = 1150
       Printer.Print Format(XTotIV, "Standard");
       Printer.CurrentX = 1350
       Printer.Print Format(Xtotal, "Standard");
       Printer.CurrentX = 1750
       Printer.Print Format(XTotIV, "Standard")
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xtotal, "Standard")
    
    End If
    If Xcontar = 4 Then
       Printer.CurrentX = 1600
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D"
       Else
          Printer.Print data_fac.Recordset("nro_med_a")
       End If
       Printer.Print ""
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xsubt, "Standard")
       Printer.Print ""
       Printer.CurrentX = 50
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D";
       Else
          Printer.Print data_fac.Recordset("nro_med_a");
       End If
       Printer.CurrentX = 750
       Printer.Print Format(Xsubt, "Standard");
       Printer.CurrentX = 1150
       Printer.Print Format(XTotIV, "Standard");
       Printer.CurrentX = 1350
       Printer.Print Format(Xtotal, "Standard");
       Printer.CurrentX = 1750
       Printer.Print Format(XTotIV, "Standard")
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xtotal, "Standard")
    
    End If
    If Xcontar = 5 Then
       Printer.Print ""
       Printer.CurrentX = 50
       If IsNull(data_fac.Recordset("nro_med_a")) = True Then
          Printer.Print "S/D";
       Else
          Printer.Print data_fac.Recordset("nro_med_a");
       End If
       Printer.Print ""
       Printer.CurrentX = 750
       Printer.Print Format(Xsubt, "Standard");
       Printer.CurrentX = 1150
       Printer.Print Format(XTotIV, "Standard");
       Printer.CurrentX = 1350
       Printer.Print Format(Xtotal, "Standard");
       Printer.CurrentX = 1750
       Printer.Print Format(XTotIV, "Standard")
       Printer.Print ""
       Printer.CurrentX = 1750
       Printer.Print Format(Xtotal, "Standard")
    
    End If

    Printer.EndDoc
End If
Unload Me

End Sub

Private Sub Form_Load()
data_fac.DatabaseName = App.Path & "\factura.mdb"
data_fac.Refresh

End Sub
