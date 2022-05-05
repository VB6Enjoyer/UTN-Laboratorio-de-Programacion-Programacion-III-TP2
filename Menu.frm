VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Menu 
   BackColor       =   &H006E943D&
   Caption         =   "Gestión de reserva de canchas"
   ClientHeight    =   7545
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   12450
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2A0E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2A681
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2AC1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2B1B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2B74F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2BCE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2C083
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2C61D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   1005
      ButtonWidth     =   1931
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reservar"
            Key             =   "reservar"
            Object.ToolTipText     =   "Hacer, editar o eliminar una reserva."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ver reservas"
            Key             =   "verreservas"
            Object.ToolTipText     =   "Ver las reservas en la base de datos."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Socios"
            Key             =   "socios"
            Object.ToolTipText     =   "Ver la lista de socios"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Añadir socio"
            Key             =   "addsocio"
            Object.ToolTipText     =   "Registrar un socio en el sistema."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Usuarios"
            Key             =   "usuarios"
            Object.ToolTipText     =   "Ver la lista de usuarios."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Configuración"
            Key             =   "config"
            Object.ToolTipText     =   "Habilitar y deshabilitar canchas y horarios."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Información"
            Key             =   "info"
            Object.ToolTipText     =   "Información del sistema."
            ImageIndex      =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar sesión"
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cerrar sesión y salir del sistema."
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   7200
   End
   Begin VB.CommandButton btnReservar 
      BackColor       =   &H006E943D&
      Caption         =   "&Reservar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      MouseIcon       =   "Menu.frx":2CBB7
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton btnConfig 
      BackColor       =   &H006E943D&
      Caption         =   "&Configuración"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      MouseIcon       =   "Menu.frx":2CEC1
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   32768
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "turno"
         Text            =   "Turno"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "cancha"
         Text            =   "Cancha"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "partido"
         Text            =   "Partido"
         Object.Width           =   15875
      EndProperty
   End
   Begin VB.CommandButton btnConsultarU 
      BackColor       =   &H006E943D&
      Caption         =   "Consultar &usuarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      MouseIcon       =   "Menu.frx":2D1CB
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton btnConsultarS 
      BackColor       =   &H006E943D&
      Caption         =   "&Consultar socios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      MouseIcon       =   "Menu.frx":2D4D5
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton btnLogoff 
      BackColor       =   &H006E943D&
      Caption         =   "Cerrar sesion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      MouseIcon       =   "Menu.frx":2D7DF
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton btnVerreservas 
      BackColor       =   &H006E943D&
      Caption         =   "&Ver reservas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      MouseIcon       =   "Menu.frx":2DAE9
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label labHora 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
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
      Height          =   480
      Left            =   2880
      TabIndex        =   10
      Top             =   960
      Width           =   945
   End
   Begin VB.Label labFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Height          =   480
      Left            =   600
      TabIndex        =   9
      Top             =   960
      Width           =   1185
   End
   Begin VB.Label labUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6360
      TabIndex        =   2
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label labPermisos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(perms)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6360
      TabIndex        =   1
      Top             =   1200
      Width           =   990
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnConsultarS_Click()
    Socios.Show
    Unload Me
End Sub

Private Sub btnConfig_Click()
    Configuracion.Show
    Unload Me
End Sub

Private Sub btnConsultarU_Click()
    Usuarios.Show
    Unload Me
End Sub

Private Sub btnLogoff_Click()
    loginUser = ""
    loginRol = 0

    Inicio.Show
    Unload Me
End Sub

Private Sub btnReservar_Click()
    Reservar.Show
    Unload Me
End Sub

Private Sub btnVerreservas_Click()
    ConsultaReservas.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    'Muestra el usuario logueado y su rol.
    labUsuario.Caption = loginUser
    
    'Muestra la fecha y la hora
    labFecha.Caption = Date
    labHora.Caption = Time
    Timer1.Enabled = True
    Timer1.Interval = 1000
        
    'Asigna permisos según el rol del usuario logueado
    Select Case loginRol
        Case 0
            labPermisos.Caption = "Admin"
        Case 1
            labPermisos.Caption = "Operador"
        Case 2
            labPermisos.Caption = "Invitado"
            btnReservar.Enabled = False
            btnConsultarU.Enabled = False
            btnConsultarS.Enabled = False
            btnConfig.Enabled = False
        Case 3
            labPermisos.Caption = "Socio"
            btnReservar.Enabled = True
            btnConsultarU.Enabled = False
            btnConsultarS.Enabled = True
            btnConfig.Enabled = False
    End Select
    
    Set cn = New ADODB.Connection
    Set rsReservas = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsReservas.Source = "Reservas"
    rsReservas.CursorType = adOpenKeyset
    rsReservas.LockType = adLockOptimistic
    rsReservas.Open "SELECT * FROM Reservas ORDER BY Turno", cn
    
    If rsReservas.EOF Then GoTo toolbarConfig
    
    rsReservas.MoveFirst
    
    Dim li As ListItem
    Do Until rsReservas.EOF
        If rsReservas.Fields("Fecha") = Date Then
            If DateDiff("h", Time, Mid(rsReservas.Fields("Turno"), 1, 5)) > 0 Then
                Set li = ListView1.ListItems.Add(, , Trim(rsReservas("Turno")))
                li.ListSubItems.Add , , Mid(Trim(rsReservas("Cancha")), 7, Len(rsReservas("Cancha")))
                
                If rsReservas.Fields("Jugador1B") <> Empty Then
                    li.ListSubItems.Add , , Trim(rsReservas("Jugador1A")) & "/" & Trim(rsReservas("Jugador1B")) & " vs " & _
                    Trim(rsReservas("Jugador2A")) & "/" & Trim(rsReservas("Jugador2B"))
                Else
                    li.ListSubItems.Add , , Trim(rsReservas("Jugador1A")) & " vs " & Trim(rsReservas("Jugador2A"))
                End If
            End If
        End If
        rsReservas.MoveNext
    Loop
    
    'Cambia el mensaje del boton de reserva si hay un socio logueado que ya hizo una reserva.
    If loginRol = 3 Then
    
        rsReservas.Close
        rsReservas.Open "SELECT * FROM Reservas WHERE Jugador1A = '" & nombreUser & "' AND Fecha LIKE '%" & Date & "%'", cn
        
        If rsReservas.EOF Then GoTo toolbarConfig
        
        If DateDiff("h", Time, Mid(rsReservas.Fields("Turno"), 1, 5)) > 0 Then btnReservar.Caption = "Editar reserva"
    End If
    
    cn.Close

toolbarConfig:
    If loginRol = 3 Then
        Toolbar1.Buttons.Item(5).Enabled = False
        Toolbar1.Buttons.Item(7).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
    End If

    If loginRol = 2 Then
        Toolbar1.Buttons.Item(1).Enabled = False
        Toolbar1.Buttons.Item(4).Enabled = False
        Toolbar1.Buttons.Item(5).Enabled = False
        Toolbar1.Buttons.Item(7).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
    End If
End Sub

Private Sub labIngresar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labIngresar.Font.Underline = True
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    If ListView1.Sorted = True And ColumnHeader.Index - 1 = ListView1.SortKey Then
        ListView1.SortOrder = 1 - ListView1.SortOrder
    Else
        ListView1.SortOrder = lvwAscending
        ListView1.SortKey = ColumnHeader.Index - 1
    End If
    
    ListView1.Sorted = True
End Sub

Private Sub Timer1_Timer()
    'Actualiza la hora
    labHora.Caption = Time
    
    'Actualiza la fecha
    If labFecha.Caption <> Date Then
        labFecha.Caption = Date
        
        Set cn = New ADODB.Connection
        Set rsReservas = New ADODB.Recordset
        
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
        
        rsReservas.Source = "Reservas"
        rsReservas.CursorType = adOpenKeyset
        rsReservas.LockType = adLockOptimistic
        rsReservas.Open "SELECT * FROM Reservas ORDER BY Fecha DESC", cn
        
        If rsReservas.EOF Then Exit Sub
        
        rsReservas.MoveFirst
        
        'Habilita todos los horarios para todas las canchas si la fecha de la última reserva es distinta a la actual
        If rsReservas.Fields("Fecha") <> Date Then
            Set rsHorarios = New ADODB.Recordset
            
            rsHorarios.Source = "Horarios"
            rsHorarios.CursorType = adOpenKeyset
            rsHorarios.LockType = adLockOptimistic
            rsHorarios.Open "SELECT * FROM Horarios", cn
            rsHorarios.MoveFirst
            
            rsHorarios.Requery
            
            Do Until rsHorarios.EOF
                With rsHorarios
                    .Fields("Cancha1") = 0
                    .Fields("Cancha2") = 0
                    .Fields("Cancha3") = 0
                    .Fields("Cancha4") = 0
                    .Fields("Cancha5") = 0
                    .Fields("Cancha6") = 0
                    .Fields("Cancha7") = 0
                    .Fields("Cancha8") = 0
                    .Fields("Cancha9") = 0
                    .Fields("Cancha10") = 0
                End With
                rsHorarios.MoveNext
            Loop
            
            rsHorarios.UpdateBatch 'Actualizamos la DB
            rsHorarios.Requery
        
            cn.Close
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Reservar
    If Button = "Reservar" Then
        Reservar.Show
        Unload Me
    End If
    
    'Consultar reservas
    If Button = "Ver reservas" Then
        ConsultaReservas.Show
        Unload Me
    End If

    'Consultar socios
    If Button = "Socios" Then
        Socios.Show
        Unload Me
    End If
    
    'Registrar un socio
    If Button = "Añadir socio" Then
        Registro.Show 1, Me
    End If
    
    'Consultar usuarios
    If Button = "Usuarios" Then
        Usuarios.Show
        Unload Me
    End If

    If Button = "Configuración" Then
        Configuracion.Show
        Unload Me
    End If

    'Información
    If Button = "Información" Then
        MsgBox ("Sistema de gestión de canchas de tenis." & Chr(10) & Chr(10) & "Creado usando Visual Basic 6 (SP6)." & Chr(10) & _
        Chr(10) & Chr(169) & "2021 Juan Ignacio Núñez.")
    End If

    'Cerrar sesión
    If Button = "Cerrar sesión" Then
        loginUser = ""
        loginRol = 0
    
        Inicio.Show
        Unload Me
    End If
End Sub
