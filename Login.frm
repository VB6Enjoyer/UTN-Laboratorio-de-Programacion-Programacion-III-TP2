VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H006E943D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso de usuario"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H006E943D&
      Caption         =   "Ver"
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
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   1000
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2220
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton btnIngresar 
      Caption         =   "&Ingresar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   1000
      PasswordChar    =   "·"
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Invitado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Socio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Operador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Admin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1005
      TabIndex        =   5
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1005
      TabIndex        =   1
      Top             =   1200
      Width           =   1725
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function ingresar()
    Dim Usuario As String
    Dim Contraseña As String
    
    Usuario = txtUsuario.text
    Contraseña = txtPassword.text
    
    'Len(Trim()) no sirve por alguna razón
    If Usuario = "" Then
        MsgBox ("Ingrese un nombre de usuario.")
        txtUsuario.SetFocus
        Exit Function
    ElseIf Contraseña = "" Then
        MsgBox ("Ingrese una contraseña.")
        txtPassword.SetFocus
        Exit Function
    End If
    
    Set cn = New ADODB.Connection
    Set rsUsuarios = New ADODB.Recordset
    Set rsSocios = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    rsUsuarios.Source = "Usuarios"
    rsUsuarios.CursorType = adOpenKeyset
    rsUsuarios.LockType = adLockOptimistic
    rsUsuarios.Open "SELECT * FROM Usuarios WHERE isHabilitado = 1", cn
    rsUsuarios.MoveFirst
    
    rsUsuarios.Find ("Usuario = '" & Usuario & "'")
    
    'Revisa hasta el final de la base de datos buscando si el usuario existe.
    If rsUsuarios.EOF = True Then
        MsgBox ("No se encontró el usuario")
        txtUsuario.SetFocus
    Else
        'Revisa que la contraseña ingresada sea correcta, le asigna el nombre y rol a las variables globales del modulo
        'y carga el menú. No deja ingresar a socios morosos.
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE Usuario = '" & rsUsuarios.Fields("Usuario") & "' AND isHabilitado = 1", cn
        
        If rsSocios.EOF = False Then
            rsSocios.MoveFirst
            If rsSocios.Fields("Estado") = "En deuda" Then
                MsgBox ("SOCIO MOROSO: Por favor, regularice su estado actual para ingresar al sistema.")
                Exit Function
            End If
        End If
        
        If rsUsuarios.Fields("Password") = Hash(Contraseña) Then
            loginUser = rsUsuarios.Fields("Usuario")
            loginRol = rsUsuarios.Fields("Rol")
            
            If loginRol = 3 Then
                nombreUser = rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido")
            End If
            
            Unload Me
            Unload Inicio
            Menu.Show
        Else
            MsgBox ("Contraseña incorrecta.")
            txtPassword.SetFocus
        End If
    End If
    
    ingresar = Null
End Function

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnIngresar_Click()
    Dim a
    a = ingresar()
End Sub

Private Sub Check1_Click()
    If Check1 Then txtPassword.PasswordChar = "" Else: txtPassword.PasswordChar = "·"
End Sub

Private Sub Label3_Click()
    txtUsuario.text = "ADMIN"
    txtPassword.text = "123"
End Sub

Private Sub Label4_Click()
    txtUsuario.text = "OPERADOR"
    txtPassword.text = "456"
End Sub

Private Sub Label5_Click()
    txtUsuario.text = "SOCIO"
    txtPassword.text = "TENIS"
End Sub

Private Sub Label6_Click()
    txtUsuario.text = "INVITADO"
    txtPassword.text = "789"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim a
        a = ingresar()
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.SetFocus
    End If
End Sub

Private Sub txtUsuario_LostFocus()
    txtUsuario = UCase(txtUsuario)
End Sub
