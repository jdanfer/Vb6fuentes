VERSION 5.00
Begin VB.Form frm_quefactcnv22 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
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
   Moveable        =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboforma 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      ItemData        =   "frm_quefactcnv.frx":0000
      Left            =   2160
      List            =   "frm_quefactcnv.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000000C0&
      Caption         =   "ND de E-FACTURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF00&
      Caption         =   "NC de E-TICKET"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "ND de E-TICKET"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "NC de E-FACTURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   6360
      Picture         =   "frm_quefactcnv.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "E-TICKET"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "E-FACTURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORMA DE PAGO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2400
      Picture         =   "frm_quefactcnv.frx":05AA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "frm_quefactcnv22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ErrorComman1

If cboforma.ListIndex = 0 Then
   Xformapcnv = 1
Else
   If cboforma.ListIndex = 1 Then
      Xformapcnv = 2
   Else
      MsgBox "No seleccionó forma de pago"
      End
   End If
End If

XAlta = 10
frm_factconve22.Show vbModal
Unload frm_quefactcnv22

Exit Sub

ErrorComman1:
            MsgBox ("Error en Command1:" & Err.Description)
            
End Sub

Private Sub Command2_Click()
On Error GoTo ErrorComman2

If cboforma.ListIndex = 0 Then
   Xformapcnv = 1
Else
   If cboforma.ListIndex = 1 Then
      Xformapcnv = 2
   Else
      MsgBox "No seleccionó forma de pago"
      End
   End If
End If

XAlta = 11
frm_factconve22.Show vbModal
Unload frm_quefactcnv22

Exit Sub

ErrorComman2:
            MsgBox ("Error en command2: " & Err.Description)
            
End Sub

Private Sub Command3_Click()
On Error GoTo ErrorComman3

Unload Me

Exit Sub

ErrorComman3:
            MsgBox ("Error: " & Err.Description)
            
End Sub

Private Sub Command4_Click()
On Error GoTo ErrorComman4

If cboforma.ListIndex = 0 Then
   Xformapcnv = 1
Else
   If cboforma.ListIndex = 1 Then
      Xformapcnv = 2
   Else
      MsgBox "No seleccionó forma de pago"
      End
   End If
End If

XAlta = 14
frm_factconve22.Show vbModal
Unload frm_quefactcnv22

Exit Sub

ErrorComman4:
            MsgBox ("Error en Command4: " & Err.Description)
            
End Sub

Private Sub Command5_Click()
On Error GoTo ErrorComman5

If cboforma.ListIndex = 0 Then
   Xformapcnv = 1
Else
   If cboforma.ListIndex = 1 Then
      Xformapcnv = 2
   Else
      MsgBox "No seleccionó forma de pago"
      End
   End If
End If

XAlta = 16
frm_factconve22.Show vbModal
Unload frm_quefactcnv22

Exit Sub

ErrorComman5:
            MsgBox ("Error en Command5: " & Err.Description)
            
            
End Sub

Private Sub Command6_Click()
On Error GoTo ErrorComman6

If cboforma.ListIndex = 0 Then
   Xformapcnv = 1
Else
   If cboforma.ListIndex = 1 Then
      Xformapcnv = 2
   Else
      MsgBox "No seleccionó forma de pago"
      End
   End If
End If

XAlta = 12
frm_factconve22.Show vbModal
Unload frm_quefactcnv22

Exit Sub

ErrorComman6:
            MsgBox ("Error en Command6: " & Err.Description)
            
            
End Sub

Private Sub Command7_Click()
On Error GoTo ErrorComman7

If cboforma.ListIndex = 0 Then
   Xformapcnv = 1
Else
   If cboforma.ListIndex = 1 Then
      Xformapcnv = 2
   Else
      MsgBox "No seleccionó forma de pago"
      End
   End If
End If

XAlta = 15
frm_factconve22.Show vbModal
Unload frm_quefactcnv22

Exit Sub

ErrorComman7:
            MsgBox ("Error en Command7:" & Err.Description)
            
            
End Sub

Private Sub Form_Resize()
On Error GoTo ErrorResi

With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With
If XcomoFactura = 1 Then
   Command1.SetFocus
   Command2.Enabled = False
   Command6.Enabled = False
   Command5.Enabled = False
Else
   Command1.Enabled = False
   Command4.Enabled = False
   Command7.Enabled = False
   Command2.SetFocus
End If
cboforma.ListIndex = 0

Exit Sub

ErrorResi:
            MsgBox ("Error en RESIZE: " & Err.Description)
            
            
End Sub
