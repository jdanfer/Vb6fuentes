VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_autorizaesp 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Autorizacion"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   4920
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Max             =   50
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3360
      Top             =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Socio Casa de Galicia. Se consultará autorización. Aguarde..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frm_autorizaesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If CgalDesde = 1 Then 'despacho
   Label1.Caption = frm_largador.txt_nomb.Text & vbCrLf
   Label1.Caption = Label1.Caption & "Lamento informarle que debemos pedir autorización para asistirle." & vbCrLf
   Label1.Caption = Label1.Caption & "Los problemas de Casa de Galicia, que no nos paga, nos obligan a esta situación." & vbCrLf
   Label1.Caption = Label1.Caption & "Vuelva a llamar en unos minutos." & vbCrLf
   Label1.Caption = Label1.Caption & "Sepa disculparnos." & vbCrLf
   Label1.Caption = Label1.Caption & "GRACIAS."
Else
   If CgalDesde = 2 Then
      Label1.Caption = frmabm.txt_apellid.Text & vbCrLf
      Label1.Caption = Label1.Caption & "Lamento informarle que debemos pedir autorización para asistirle." & vbCrLf
      Label1.Caption = Label1.Caption & "Los problemas de Casa de Galicia, que no nos paga, nos obligan a esta situación." & vbCrLf
      Label1.Caption = Label1.Caption & "Aguarde unos minutos." & vbCrLf
      Label1.Caption = Label1.Caption & "Sepa disculparnos." & vbCrLf
      Label1.Caption = Label1.Caption & "GRACIAS."
   Else
      If CgalDesde = 3 Then
         Label1.Caption = frm_especialistas.t_nompac.Text & vbCrLf
         Label1.Caption = Label1.Caption & "Lamento informarle que debemos pedir autorización para asistirle." & vbCrLf
         Label1.Caption = Label1.Caption & "Los problemas de Casa de Galicia, que no nos paga, nos obligan a esta situación." & vbCrLf
         Label1.Caption = Label1.Caption & "Vuelva a llamar en unos minutos." & vbCrLf
         Label1.Caption = Label1.Caption & "Sepa disculparnos." & vbCrLf
         Label1.Caption = Label1.Caption & "GRACIAS."
      Else
         Label1.Caption = "Señor/a,"
         Label1.Caption = Label1.Caption & "Lamento informarle que debemos pedir autorización para asistirle." & vbCrLf
         Label1.Caption = Label1.Caption & "Los problemas de Casa de Galicia, que no nos paga, nos obligan a esta situación." & vbCrLf
         Label1.Caption = Label1.Caption & "Aguarde unos minutos." & vbCrLf
         Label1.Caption = Label1.Caption & "Sepa disculparnos." & vbCrLf
         Label1.Caption = Label1.Caption & "GRACIAS."
      End If
   End If
End If

End Sub

Private Sub Timer1_Timer()
If pb1.Value = 50 Then
   Timer1.Enabled = False
   MsgBox "La autorización será comunicada desde el sistema SAPP.", vbInformation
   Unload Me
Else
   pb1.Value = pb1.Value + 1
End If

End Sub
