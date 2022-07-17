VERSION 5.00
Begin VB.Form frm_espmodificaconsultorio 
   Caption         =   "modificacion de consultorio"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton elegir 
      Caption         =   "elegir"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Elija consultorio nuevo"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frm_espmodificaconsultorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private base As String
Private jsonConsultoriosBase As String
Private reservado As Boolean
Private id_hora_reserva As Integer
'Private id_hora_seleccionada As Integer
Private fecha As String
Private idMedico As Integer
Private horaInicio As String
Private horaFin As String
Private url As String


Private Sub Command1_Click()
    reservado = False
    On Error GoTo 0
    Unload Me
End Sub

Private Sub elegir_Click()
    On Error GoTo WSError
    
    Dim itemSeleccionado As String
    Set p = JSON.parse(jsonConsultoriosBase)
    If Combo1.ListIndex <> Combo1.ListCount - 1 Then
         idConsultorio = JSON.toString(p.Item(Combo1.ListIndex + 1).Item("id"))
         XfecstrGuiones = Format(fecha, "yyyy-mm-dd")
         
         body = "medico=" & idMedico & "&inicio=" & XfecstrGuiones & "T" & horaInicio & ":00&fin=" & XfecstrGuiones & "T" & horaFin & ":00"
         Set obj = consumirServicio("PUT", url & "/bases/" & base & "/consultorios/" & idConsultorio & "/disponibilidades", body & "&forzar=true")
         
         estado = obj.Status
         'MsgBox estado, vbInformation
         
         Set p = JSON.parse(obj.responseText)
         Select Case estado
             Case 200
                 'MsgBox p.Item("horaConsultorio").Item("id")
                 id_hora_reserva = p.Item("horaConsultorio").Item("id")
                 reservado = True
                 On Error GoTo 0
                 Unload Me
             Case 404
                 horasSuperposicion = p.Item("superposiciones").Item(1).Item("superposicionHoras")
                 minutosSuperposicion = p.Item("superposiciones").Item(1).Item("superposicionMinutos")
                 Xsionosuperpone = MsgBox("No hay lugar en este consultorio, desea superponer las horas? (superposicion  " & horasSuperposicion & " horas, " & minutosSuperposicion & " minutos )", vbExclamation + vbYesNo)
                 If Xsionosuperpone = vbYes Then
                     body = "medico=" & idMedico & "&inicio=" & XfecstrGuiones & "T" & horaInicio & ":00&fin=" & XfecstrGuiones & "T" & horaFin & ":00"
                     Set obj = consumirServicio("PUT", url & "/bases/" & base & "/consultorios/" & idConsultorio & "/disponibilidades", body & "&forzar=true&superponer=true")
                     estado = obj.Status
                     Set pp = JSON.parse(obj.responseText)
                     id_hora_reserva = pp.Item("horaConsultorio").Item("id")
                     reservado = True
                     On Error GoTo 0
                     Unload Me
                 End If
             Case Else
                 'MsgBox "else, no se pudo reservar consultorio "
                 On Error GoTo 0
                 Unload Me
         End Select
    Else
        'MsgBox "no reservar"
        id_hora_reserva = 0
        reservado = True
        On Error GoTo 0
        Unload Me
    End If

    Exit Sub
WSError:
    MsgBox Err.Description & " (servicio de reservas caído o sin conexion, favor, contactarse con computos para poder modificar el consultorio). "
    reservado = False
    On Error GoTo 0 ' desactivo error handler para que el resto quede como estaba
    GetParameters.log (Err.Description)
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo WSError
    'obtengo consultorios de la base'
    reservado = False
    'funcion obtiene base retorna obj'
    Set obj = consumirServicio("GET", url & "/bases/" & base & "/consultorios", "")
    'obj.Open "GET", url & "/bases/" & base & "/consultorios", False
    'obj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    'obj.setRequestHeader "User-Agent", "SAPP VB6"
    'fin funcion obtiene base retorna obj'
    'obj.send
    
    jsonConsultoriosBase = obj.responseText
    Set p = JSON.parse(jsonConsultoriosBase)
    
    For i = 1 To p.count
        Combo1.AddItem JSON.toString(p.Item(i).Item("id")) & JSON.toString(p.Item(i).Item("descConsultorio")), i - 1
    Next
    Combo1.AddItem "no ocupar consultorio"
    
    Combo1.ListIndex = 0
    Exit Sub
WSError:
    MsgBox Err.Description & " (servicio de reservas caído o sin conexion, favor, contactarse con computos). "
    GetParameters.log (Err.Description)
    On Error GoTo 0
End Sub



Public Property Get PassBase() As String
PassVar = base
End Property

Public Property Let PassBase(ByVal vNewValue As String)
 base = vNewValue
End Property

Public Property Let PassFecha(ByVal vNewValue As String)
    fecha = vNewValue
End Property

Public Property Let PassIdMedico(ByVal vNewValue As Integer)
    idMedico = vNewValue
End Property

Public Property Let PassHoraInicio(ByVal vNewValue As String)
    horaInicio = vNewValue
End Property

Public Property Let PassHoraFin(ByVal vNewValue As String)
    horaFin = vNewValue
End Property


Public Property Let PassUrl(ByVal vNewValue As String)
    url = vNewValue
End Property






Public Property Get PassReservado() As Boolean
PassReservado = reservado
End Property

Public Property Get PassIdHoraReserva() As Integer
PassIdHoraReserva = id_hora_reserva
End Property

'Public Property Let PassId_hora_seleccionada(ByVal idSeleccionado As Integer)
'    id_hora_seleccionada = idSeleccionado
'End Property



