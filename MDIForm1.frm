VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H006E943D&
   Caption         =   "Sistema de gestión de canchas de tenis"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12420
      _ExtentX        =   21908
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
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
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2536
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Reservar
    If Button = "Reservar" Then
        Reservar.Show
    End If
    
    'Consultar reservas
    If Button = "Ver reservas" Then
        ConsultaReservas.Show
    End If

    'Consultar socios
    If Button = "Socios" Then
        Socios.Show
    End If
    
    'Registrar un socio
    If Button = "Añadir socio" Then
        Registro.Show 1, Me
    End If
    
    'Consultar usuarios
    If Button = "Usuarios" Then
        Usuarios.Show
    End If

    If Button = "Configuración" Then
        Configuracion.Show
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


