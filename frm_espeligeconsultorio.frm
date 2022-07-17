VERSION 5.00
Begin VB.Form frm_espeligeconsultorio 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton reservar 
      Caption         =   "Reservar Consultorio"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   $"frm_espeligeconsultorio.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frm_espeligeconsultorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Superposiciones As String
Private urlWS As String
Private bodyWS As String
Private reservoConsultorio As String




' hacer un boton "no registrar reserva" para que siga curso normal? '
Private Sub Command1_Click()
    reservoConsultorio = False
    Unload Me
End Sub

Private Sub reservar_Click()
    Dim consultorio_seleccion As String
    Dim response As String
    Dim id_hora_consultorio As Integer
    On Error GoTo WSError
    Set p = JSON.parse(Superposiciones)
    consultorio_seleccion = JSON.toString(p.Item("superposiciones").Item(Combo1.ListIndex + 1).Item("id_consultorio"))
    'consumo servicio'
    Set obj = consumirServicio("PUT", urlWS & "/" & consultorio_seleccion & "/disponibilidades", bodyWS & "&forzar=true&superponer=true")
    response = obj.responseText
    Set p = JSON.parse(response)
    id_hora_consultorio = CInt(JSON.toString(p.Item("horaConsultorio").Item("id")))
    
    With frm_especialistas
        .Pass_id_hora_reserva = id_hora_consultorio
        Debug.Print .Pass_id_hora_reserva
    End With
    reservoConsultorio = True
    Unload Me
    Exit Sub
WSError:
    MsgBox Err.Description & " (servicio de reservas caído o sin conexion, favor, contactarse con computos), podrá seguir pero el consultorio de la base no sera reservado. "
    On Error GoTo 0 ' desactivo error handler para que el resto quede como estaba
    reservoConsultorio = True 'para que siga con la creacion de la fecha'
    GetParameters.log (Err.Description)
    Unload Me
End Sub

Private Sub Form_Load()
    reservoConsultorio = False
    On Error GoTo WSError
    Set p = JSON.parse(Superposiciones)
    Combo1.Clear
    For index = 1 To p.Item("superposiciones").count
        Combo1.AddItem JSON.toString(p.Item("superposiciones").Item(index).Item("id_consultorio")) & " - " & JSON.toString(p.Item("superposiciones").Item(index).Item("desc_consultorio")) & " -  (superposicion: " & JSON.toString(p.Item("superposiciones").Item(index).Item("superposicionHoras")) & " horas, " & JSON.toString(p.Item("superposiciones").Item(index).Item("superposicionMinutos")) & " minutos" & ")", index - 1
    Next
    Combo1.ListIndex = 0
    Exit Sub
WSError:
    MsgBox Err.Description & " (servicio de reservas caído o sin conexion, favor, contactarse con computos), podrá seguir pero el consultorio de la base no sera reservado. "
    On Error GoTo 0
    reservoConsultorio = True
    Unload Me
End Sub

Public Property Get PassVar() As String
PassVar = Superposiciones
End Property

Public Property Let PassVar(ByVal vNewValue As String)
 Superposiciones = vNewValue
End Property


Public Property Get PassUrlWS() As String
PassUrlWS = urlWS
End Property

Public Property Let PassUrlWS(ByVal vNewValue As String)
 urlWS = vNewValue
End Property

Public Property Get PassBodyWS() As String
PassBodyWS = bodyWS
End Property

Public Property Let PassBodyWS(ByVal vNewValue As String)
 bodyWS = vNewValue
End Property

Public Property Get PassReservoConsultorio() As Boolean
PassReservoConsultorio = reservoConsultorio
End Property

Public Property Let PassReservoConsultorio(ByVal vNewValue As Boolean)
 reservoConsultorio = vNewValue
End Property

