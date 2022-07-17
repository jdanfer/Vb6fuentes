VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_usuarios 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Usuarios"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9825
   Icon            =   "frm_usuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from usuarios order by nombre"
      Top             =   5880
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
      Height          =   375
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarios"
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton b_can 
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
      Left            =   9000
      Picture         =   "frm_usuarios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton b_acep 
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
      Left            =   240
      Picture         =   "frm_usuarios.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Guarda los datos ingresados y/o modificados"
      Top             =   5520
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9255
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Agenda Hisopados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   29
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "frm_usuarios.frx":0F56
         Left            =   6360
         List            =   "frm_usuarios.frx":0F60
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox t_correo 
         Height          =   375
         Left            =   2160
         TabIndex        =   27
         Top             =   2280
         Width           =   4215
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_usuarios.frx":0F6C
         Height          =   1935
         Left            =   5520
         OleObjectBlob   =   "frm_usuarios.frx":0F80
         TabIndex        =   23
         Top             =   3240
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   21
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4920
         Picture         =   "frm_usuarios.frx":1963
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Agregar los permisos básicos"
         Top             =   4680
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4920
         Picture         =   "frm_usuarios.frx":1EED
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Agregar todos los permisos"
         Top             =   3720
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Doble click para borrar una opción"
         Top             =   3720
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4920
         Picture         =   "frm_usuarios.frx":2477
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Agregar módulo a la lista"
         Top             =   3120
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frm_usuarios.frx":2A01
         Left            =   240
         List            =   "frm_usuarios.frx":2A03
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3360
         Width           =   4575
      End
      Begin VB.TextBox t_ced 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   14
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_usuarios.frx":2A05
         Left            =   2160
         List            =   "frm_usuarios.frx":2A21
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txt_repcon 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   7200
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txt_contra 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txt_nomb 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txt_usua 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF0000&
         Caption         =   "Servicios AP:"
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
         Left            =   4320
         TabIndex        =   26
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF0000&
         Caption         =   "Correo electrónico:"
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
         TabIndex        =   25
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         Caption         =   "Lista de usuarios"
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
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "Código médico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6360
         TabIndex        =   22
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   3
         X1              =   0
         X2              =   9240
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "Opciones autorizadas"
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
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   4575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Documento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "Miembro de:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Repetir contraseña:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Contraseña:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
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
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Nombre completo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
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
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2880
      Picture         =   "frm_usuarios.frx":2A98
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2415
   End
End
Attribute VB_Name = "frm_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xban, XX As Integer
Xban = 0

If List1.ListCount >= 1 Then
   For XX = 1 To List1.ListCount
       List1.ListIndex = XX - 1
       If List1.List(List1.ListIndex) = Combo2.Text Then
          Xban = 1
       End If
   Next
Else
   Xban = 0
End If

If Combo2.ListIndex >= 0 And Xban = 0 Then
   List1.AddItem Combo2.Text
End If

End Sub

Private Sub b_acep_Click()
Dim Xcontar As Long
Dim Xidus As Long
Dim XX As Integer

If txt_usua.Text <> "" Then
   If txt_contra.Text <> "" Then
      b_acep.Enabled = False
      If Trim(txt_contra.Text) = Trim(txt_repcon.Text) Then
         Data1.Recordset.FindFirst "usuario ='" & txt_usua.Text & "'"
         If Not Data1.Recordset.NoMatch Then
            If txt_nomb.Text <> Data1.Recordset("nombre") Then
               Data1.Recordset.Edit
               Data1.Recordset("nombre") = txt_nomb.Text
               Data1.Recordset.Update
            End If
            If Combo1.Text = "USUARIOS DESP" Then
               If IsNull(Data1.Recordset("med")) = False Then
                  If Data1.Recordset("med") <> "S" Then
                     Data1.Recordset.Edit
                     Data1.Recordset("med") = "S"
                     Data1.Recordset.Update
                  End If
               Else
                  Data1.Recordset.Edit
                  Data1.Recordset("med") = "S"
                  Data1.Recordset.Update
               End If
            Else
               If Check1.Value = 1 Then
                  If IsNull(Data1.Recordset("med")) = False Then
                     If Data1.Recordset("med") <> "R" Then
                        Data1.Recordset.Edit
                        Data1.Recordset("med") = "R"
                        Data1.Recordset.Update
                     End If
                  Else
                     Data1.Recordset.Edit
                     Data1.Recordset("med") = "R"
                     Data1.Recordset.Update
                  End If
               Else
                  If IsNull(Data1.Recordset("med")) = False Then
                     If Data1.Recordset("med") <> "A" Then
                        Data1.Recordset.Edit
                        Data1.Recordset("med") = "A"
                        Data1.Recordset.Update
                     End If
                  Else
                     Data1.Recordset.Edit
                     Data1.Recordset("med") = "A"
                     Data1.Recordset.Update
                  End If
               End If
            End If
            If Combo3.ListIndex = 1 Then
               If IsNull(Data1.Recordset("serv_ap")) = False Then
                  If Data1.Recordset("serv_ap") <> "S" Then
                     Data1.Recordset.Edit
                     Data1.Recordset("serv_ap") = "S"
                     Data1.Recordset("correo_ap") = t_correo.Text
                     Data1.Recordset.Update
                  Else
                     If IsNull(Data1.Recordset("correo_ap")) = False Then
                        If Data1.Recordset("correo_ap") <> t_correo.Text Then
                           Data1.Recordset.Edit
                           Data1.Recordset("correo_ap") = t_correo.Text
                           Data1.Recordset.Update
                        End If
                     Else
                        Data1.Recordset.Edit
                        Data1.Recordset("correo_ap") = t_correo.Text
                        Data1.Recordset.Update
                     End If
                  End If
               Else
                  Data1.Recordset.Edit
                  Data1.Recordset("serv_ap") = "S"
                  Data1.Recordset("correo_ap") = t_correo.Text
                  Data1.Recordset.Update
               End If
            Else
               If IsNull(Data1.Recordset("serv_ap")) = False Then
                  Data1.Recordset.Edit
                  Data1.Recordset("serv_ap") = Null
                  Data1.Recordset("correo_ap") = Null
                  Data1.Recordset.Update
               End If
            End If
            If Text1.Text <> "" Then
               If IsNull(Data1.Recordset("codmed")) = False Then
                  If Data1.Recordset("codmed") <> Text1.Text Then
                     Data1.Recordset.Edit
                     Data1.Recordset("codmed") = Text1.Text
                     Data1.Recordset.Update
                  End If
               Else
                  Data1.Recordset.Edit
                  Data1.Recordset("codmed") = Text1.Text
                  Data1.Recordset.Update
               End If
            Else
               If IsNull(Data1.Recordset("codmed")) = False Then
                  Data1.Recordset.Edit
                  Data1.Recordset("codmed") = Null
                  Data1.Recordset.Update
               End If
            End If
            If txt_contra.Text <> Data1.Recordset("clave") Then
               Data1.Recordset.Edit
               Data1.Recordset("clave") = txt_contra.Text
               Data1.Recordset.Update
            End If
            If Combo1.Text <> Data1.Recordset("tipo") Then
               Data1.Recordset.Edit
               Data1.Recordset("tipo") = Combo1.Text
               Data1.Recordset.Update
            End If
            If List1.ListCount >= 1 Then
               For XX = 1 To List1.ListCount
                   List1.ListIndex = XX - 1
                   Data2.RecordSource = "select * from usua_permisos where id_usuario =" & Data1.Recordset("id") & " and opcion ='" & List1.List(List1.ListIndex) & "' order by opcion"
                   Data2.Refresh
                   If Data2.Recordset.RecordCount > 0 Then
                                        
                   Else
                      Data2.Recordset.AddNew
                      Data2.Recordset("id_usuario") = Data1.Recordset("id")
                      Data2.Recordset("opcion") = List1.List(List1.ListIndex)
                      Data2.Recordset.Update
                   End If
               Next
            End If
            If t_ced.Text <> "" Then
               Data2.RecordSource = "Select * from cap_ciap where des_cap ='" & txt_usua.Text & "'"
               Data2.Refresh
               If Data2.Recordset.RecordCount > 0 Then
                  If IsNull(Data2.Recordset("cod_cap")) = False Then
                     If Data2.Recordset("cod_cap") <> t_ced.Text Then
                        Data2.Recordset.Edit
                        Data2.Recordset("cod_cap") = t_ced.Text
                        Data2.Recordset.Update
                     End If
                  Else
                     Data2.Recordset.Edit
                     Data2.Recordset("cod_cap") = t_ced.Text
                     Data2.Recordset.Update
                  End If
               Else
                  Data2.RecordSource = "Select * from cap_ciap order by id DESC"
                  Data2.Refresh
                  Xidus = Data2.Recordset("id") + 1
                  Data2.Recordset.AddNew
                  Data2.Recordset("des_cap") = txt_usua.Text
                  Data2.Recordset("cod_cap") = t_ced.Text
                  Data2.Recordset.Update
               End If
            Else
            End If
         Else
            Data1.Recordset.MoveLast
            Xcontar = Data1.Recordset("id") + 1
            Data1.Recordset.AddNew
            Data1.Recordset("id") = Xcontar
            Data1.Recordset("usuario") = UCase(txt_usua.Text)
            Data1.Recordset("clave") = txt_contra.Text
            Data1.Recordset("tipo") = Combo1.Text
            Data1.Recordset("nombre") = txt_nomb.Text
            If Text1.Text <> "" Then
               Data1.Recordset("codmed") = Text1.Text
            End If
            If Combo1.Text = "USUARIOS DESP" Then
               Data1.Recordset("med") = "S"
            Else
               Data1.Recordset("med") = "A"
            End If
            If Combo3.ListIndex = 1 Then
               Data1.Recordset("serv_ap") = "S"
               If Trim(t_correo.Text) <> "" Then
                  Data1.Recordset("correo_ap") = t_correo.Text
               Else
                  Data1.Recordset("correo_ap") = "sappsistemas@gmail.com"
               End If
            End If
            Data1.Recordset.Update
            
            If List1.ListCount >= 1 Then
               For XX = 1 To List1.ListCount
                   List1.ListIndex = XX - 1
                   Data2.RecordSource = "select * from usua_permisos where id_usuario =" & Xcontar & " and opcion ='" & List1.List(List1.ListIndex) & "' order by opcion"
                   Data2.Refresh
                   If Data2.Recordset.RecordCount > 0 Then
                                        
                   Else
                      Data2.Recordset.AddNew
                      Data2.Recordset("id_usuario") = Xcontar
                      Data2.Recordset("opcion") = List1.List(List1.ListIndex)
                      Data2.Recordset.Update
                   End If
               Next
            End If
            
            Data2.RecordSource = "Select * from cap_ciap where des_cap ='" & txt_usua.Text & "'"
            Data2.Refresh
            If Data2.Recordset.RecordCount > 0 Then
               If IsNull(Data2.Recordset("cod_cap")) = False Then
                  If Data2.Recordset("cod_cap") <> t_ced.Text Then
                     Data2.Recordset.Edit
                     Data2.Recordset("cod_cap") = t_ced.Text
                     Data2.Recordset.Update
                  End If
               Else
                  Data2.Recordset.Edit
                  Data2.Recordset("cod_cap") = t_ced.Text
                  Data2.Recordset.Update
               End If
            Else
               Data2.RecordSource = "Select * from cap_ciap order by id DESC"
               Data2.Refresh
               Xidus = Data2.Recordset("id") + 1
               Data2.Recordset.AddNew
               Data2.Recordset("id") = Xidus
               Data2.Recordset("des_cap") = txt_usua.Text
               Data2.Recordset("cod_cap") = t_ced.Text
               Data2.Recordset.Update
            End If
         End If
         MsgBox "Terminado..."
         b_acep.Enabled = True
         Unload Me
      Else
         MsgBox "Verifique contraseña", vbInformation, "Mensaje"
         txt_contra.SetFocus
      End If
   Else
      MsgBox "Debe ingresar contraseña", vbInformation, "Mensaje"
      txt_contra.SetFocus
   End If
Else
   MsgBox "Debe ingresar usuario", vbInformation, "Mensaje"
   txt_usua.SetFocus
End If
       
End Sub

Private Sub b_can_Click()
Unload Me

End Sub

Private Sub Command2_Click()
List1.Clear

If Data3.Recordset.RecordCount > 0 Then
   Data3.Recordset.MoveFirst
   Do While Not Data3.Recordset.EOF
      List1.AddItem Data3.Recordset("opcion")
      Data3.Recordset.MoveNext
   Loop
End If

End Sub

Private Sub Command3_Click()

List1.Clear
List1.AddItem "Mantenimiento"
List1.AddItem "Control medicación base"
List1.AddItem "Control de consultas"
List1.AddItem "Control Actos enfermería"
List1.AddItem "Control Entrega HC"
List1.AddItem "Facturar"
List1.AddItem "Ver_Deuda"
List1.AddItem "historial"
List1.AddItem "Estadísticas"
List1.AddItem "Ingreso de caja"
List1.AddItem "Re imprimir Factura"
List1.AddItem "Control actos de enfermería"
List1.AddItem "Ver llamados a domicilio"
List1.AddItem "Cargar Electros a HCE"
List1.AddItem "Vencimientos"
List1.AddItem "Servicios"
List1.AddItem "Fechas Especialistas"
List1.AddItem "Solicitud Insumos Informáticos"
List1.AddItem "Ficha Personal"
List1.AddItem "Cambiar contraseña"
List1.AddItem "Solicitud asistencia técnica informática"
List1.AddItem "Solicitud asistencia a mantenimiento"
List1.AddItem "Solicitud a RRHH"
List1.AddItem "Solicitud a Padrón social"
List1.AddItem "Iniciativas del Personal"
List1.AddItem "Ventas por Médico"
List1.AddItem "Laboratorios"
List1.AddItem "Informes control HC MOVILES"
List1.AddItem "historial Soc"


End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data3.RecordSource = "select * from opciones_menu order by opcion"
Data3.Refresh
List1.Clear
If Data3.Recordset.RecordCount > 0 Then
   Data3.Recordset.MoveFirst
   Do While Not Data3.Recordset.EOF
      Combo2.AddItem Data3.Recordset("opcion")
      Data3.Recordset.MoveNext
   Loop
End If

Combo1.ListIndex = 0

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub List1_DblClick()
Dim Xsnlper As String
Xsnlper = MsgBox("Desea borrar la opción seleccionada?", vbExclamation + vbYesNo)
If Xsnlper = vbYes Then
   Data1.Recordset.FindFirst "usuario ='" & txt_usua.Text & "'"
   If Not Data1.Recordset.NoMatch Then
      If List1.ListIndex >= 0 Then
         Data2.RecordSource = "select * from usua_permisos where id_usuario =" & Data1.Recordset("id") & " and opcion ='" & List1.List(List1.ListIndex) & "' order by opcion"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            Data2.Recordset.Delete
            Data2.Refresh
         End If
      End If
      List1.RemoveItem List1.ListIndex
   End If
End If


End Sub

Private Sub txt_usua_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(VBA.UCase(VBA.Chr(KeyAscii)))

If KeyAscii = 13 Then
   txt_nomb.SetFocus
End If

End Sub

Private Sub txt_usua_LostFocus()
Dim Xresp As String
If txt_usua.Text <> "" Then
   Data1.Recordset.FindFirst "usuario ='" & txt_usua.Text & "'"
   If Not Data1.Recordset.NoMatch Then
      Xresp = MsgBox("Ya EXISTE, desea modificar?", vbInformation + vbYesNo, "Mensaje")
      If Xresp = vbYes Then
         txt_nomb.Text = Data1.Recordset("nombre")
         txt_contra.Text = Data1.Recordset("clave")
         txt_repcon.Text = Data1.Recordset("clave")
         If IsNull(Data1.Recordset("codmed")) = False Then
            Text1.Text = Data1.Recordset("codmed")
         Else
            Text1.Text = ""
         End If
         If IsNull(Data1.Recordset("serv_ap")) = False Then
            If Data1.Recordset("serv_ap") = "S" Then
               Combo3.ListIndex = 1
            Else
               Combo3.ListIndex = -1
            End If
         Else
            Combo3.ListIndex = -1
         End If
         If IsNull(Data1.Recordset("correo_ap")) = False Then
            t_correo.Text = Data1.Recordset("correo_ap")
         Else
            t_correo.Text = ""
         End If
         Combo1.Text = Data1.Recordset("tipo")
         If txt_usua.Text <> "" Then
            Data2.RecordSource = "Select * from cap_ciap where des_cap ='" & txt_usua.Text & "'"
            Data2.Refresh
            If Data2.Recordset.RecordCount > 0 Then
               If IsNull(Data2.Recordset("cod_cap")) = False Then
                  t_ced.Text = Data2.Recordset("cod_cap")
               Else
                  t_ced.Text = ""
               End If
            Else
               t_ced.Text = ""
            End If
            Data2.RecordSource = "select * from usua_permisos where id_usuario =" & Data1.Recordset("id") & " order by opcion"
            Data2.Refresh
            If Data2.Recordset.RecordCount > 0 Then
               List1.Clear
               Data2.Recordset.MoveFirst
               Do While Not Data2.Recordset.EOF
                  List1.AddItem Data2.Recordset("opcion")
                  Data2.Recordset.MoveNext
               Loop
            Else
               List1.Clear
            End If
         Else
            t_ced.Text = ""
         End If
      Else
         Unload Me
      End If
   Else
      List1.Clear
   End If
End If

      
      
End Sub
