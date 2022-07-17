VERSION 5.00
Begin VB.Form frmquefac 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "que factura"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_conce 
      Caption         =   "data_conce"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DEVOLUCION RECIBO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RECIBO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ND E-FACTURA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NC E-FACTURA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E-FACTURA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ND de E-TICKET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NC de E-TICKET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      ItemData        =   "frmquefac.frx":0000
      Left            =   3120
      List            =   "frmquefac.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E-TICKET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Data data_ctr 
      Caption         =   "data_ctr"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   5880
      Picture         =   "frmquefac.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancelar facturación"
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forma de Pago:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   360
      Picture         =   "frmquefac.frx":05AA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmquefac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmabm.btn_fact.Enabled = True
Unload Me

End Sub

Private Sub Command2_Click()
XQuefac = 101
If Combo1.ListIndex = 0 Then
   Xfpago = 1
Else
   If Combo1.ListIndex = 1 Then
      Xfpago = 2
   Else
      Xfpago = 1
   End If
End If
frm_factura.Show vbModal

Unload frmquefac

      
End Sub


Private Sub Command3_Click()
XQuefac = 102
If Combo1.ListIndex = 0 Then
   Xfpago = 1
Else
   If Combo1.ListIndex = 1 Then
      Xfpago = 2
   Else
      Xfpago = 1
   End If
End If
frm_factura.Show vbModal
Unload frmquefac


End Sub

Private Sub Command4_Click()
XQuefac = 103
If Combo1.ListIndex = 0 Then
   Xfpago = 1
Else
   If Combo1.ListIndex = 1 Then
      Xfpago = 2
   Else
      Xfpago = 1
   End If
End If
frm_factura.Show vbModal
Unload frmquefac

End Sub

Private Sub Command5_Click()
Dim Xelrut As String
Dim Xelrquees As String

Xelrut = ""
XQuefac = 111
If Combo1.ListIndex = 0 Then
   Xfpago = 1
Else
   If Combo1.ListIndex = 1 Then
      Xfpago = 2
   Else
      Xfpago = 1
   End If
End If
Dim Xdig, Xtot, Xfactor, Xtot2, Xrut, i As Integer
Xdig = 0
Xtot = 0
Xfactor = 0
Xtot2 = 0
Data1.RecordSource = "Select * from convenio where cnv_codigo ='" & frmabm.txt_codcnv.Text & "'"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   If IsNull(Data1.Recordset("cnv_entre")) = False Then
      If Trim(Data1.Recordset("cnv_entre")) <> "" Then
         If Val(Data1.Recordset("cnv_cuenta")) = Val(frmabm.txt_mat.Caption) Then
            If IsNull(Data1.Recordset("cnv_ruc")) = False Then
               Xelrquees = Data1.Recordset("cnv_ruc")
            Else
               Xelrquees = ""
            End If
         Else
            Xelrquees = ""
         End If
      Else
         Xelrquees = ""
      End If
   Else
      Xelrquees = ""
   End If
Else
   Xelrquees = ""
End If
If Xelrquees <> "" Then
   Xelrut = InputBox("Ingrese un número de RUT válido", , Xelrquees)
Else
   Xelrut = InputBox("Ingrese un número de RUT válido")
End If
If Xelrut <> "" Then
   If Len(Trim(Xelrut)) = 12 Then
      If IsNumeric(Xelrut) Then
         Xdig = Val(Mid(Xelrut, 12, 1))
         Xrut = Val(Mid(Xelrut, 1, 12))
         Xtot = 0
         Xfactor = 2
         For i = 1 To 11
             If i = 1 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 4
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 2 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 3
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 3 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 2
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 4 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 9
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 5 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 8
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 6 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 7
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 7 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 6
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 8 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 5
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 9 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 4
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 10 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 3
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 11 Then
                Xtot = Val(Mid(Xelrut, i, 1)) * 2
                Xtot2 = Xtot2 + Xtot
             End If
         Next
         Xtot = Xtot2 Mod 11
         If Xtot > 0 Then
            Xtot = 11 - Xtot
         Else
            Xdig = 0
         End If
         If Xtot = 11 Then
            Xdig = 0
         Else
            Xdig = Xtot
         End If
         If Xdig = Val(Mid(Xelrut, 12, 1)) Then
            Xelrutfact = Xelrut
            frm_factura.Show vbModal
            Unload frmquefac
         Else
            MsgBox "El RUT ingresado no es válido", vbCritical
         End If
      Else
         MsgBox "El RUT ingresado debe contener solo números"
      End If
   Else
      MsgBox "La cantidad de dígitos del RUT no es correcta"
   End If
Else
   MsgBox "No ingresó RUT", vbCritical
End If

End Sub

Private Sub Command6_Click()
XQuefac = 112
If Combo1.ListIndex = 0 Then
   Xfpago = 1
Else
   If Combo1.ListIndex = 1 Then
      Xfpago = 2
   Else
      Xfpago = 1
   End If
End If

frm_factura.Show vbModal
Unload frmquefac

End Sub

Private Sub Command7_Click()
XQuefac = 113
If Combo1.ListIndex = 0 Then
   Xfpago = 1
Else
   If Combo1.ListIndex = 1 Then
      Xfpago = 2
   Else
      Xfpago = 1
   End If
End If
frm_factura.Show vbModal
Unload frmquefac

End Sub

Private Sub Command8_Click()
XQuefac = 4
If Combo1.ListIndex = 0 Then
   Xfpago = 1
Else
   If Combo1.ListIndex = 1 Then
      Xfpago = 2
   Else
      Xfpago = 1
   End If
End If
frm_factura.Show vbModal
Unload frmquefac

End Sub

Private Sub Command9_Click()
XQuefac = 21
If Combo1.ListIndex = 0 Then
   Xfpago = 1
Else
   If Combo1.ListIndex = 1 Then
      Xfpago = 2
   Else
      Xfpago = 1
   End If
End If
frm_factura.Show vbModal
Unload frmquefac

End Sub

Private Sub Form_Load()
Dim Xfec1 As Date
Dim Xdiferen As Long
Combo1.ListIndex = 0

Data2.DatabaseName = App.path & "\parse.mdb"
Data2.RecordSource = "parsec0"
Data2.Refresh

    data_ctr.DatabaseName = App.path & "\ctrf.mdb"
    data_ctr.RecordSource = "ctrf"
    data_ctr.Refresh
    Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
    
    Xfec1 = data_ctr.Recordset("fecha")
    Xfec2 = Date
    Xdiferen = DateDiff("d", Xfec1, Xfec2)
    data_conce.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_conce.RecordSource = "Select * from convenio where cnv_codigo ='" & frmabm.txt_codcnv.Text & "'"
    data_conce.Refresh
    If data_conce.Recordset.RecordCount > 0 Then
       If IsNull(data_conce.Recordset("cnv_aran")) = False Then
          Xop1 = data_conce.Recordset("cnv_aran")
       Else
          Xop1 = 0
       End If
    Else
       Xop1 = 0
    End If
    If WElusuario = "JFERNAN" Or WElusuario = "MCOSTA" Or WElusuario = "RREGUEIRA" Or WElusuario = "CDEMORAES" Or frm_menu.data_parse.Recordset("base") = 11 Then
    Else
        If Xdiferen >= 15 Then
           MsgBox "Hay un error en las fechas, VERIFIQUE!!", vbCritical
           End
        Else
           If Xdiferen <= 5 Then
           Else
              MsgBox "Hay un error en las fechas, VERIFIQUE!!", vbCritical
              Unload Me
           End If
        End If
    End If
    If Data2.Recordset("base") = 20 Then
       Combo1.ListIndex = 1
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


