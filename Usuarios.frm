VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Usuarios 
   BackColor       =   &H006E943D&
   Caption         =   "Consulta de usuarios"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   12450
   Begin VB.CommandButton btnEditar 
      Caption         =   "Editar datos del usuario"
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
      Left            =   3900
      TabIndex        =   9
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton btnDeshabilitar 
      Caption         =   "Deshabilitar usuario"
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
      Left            =   1620
      TabIndex        =   8
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton btnVolver 
      Caption         =   "Volver al menú"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   4
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H006E943D&
      Caption         =   "Buscar por:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton btnDesh 
         Caption         =   "Mostrar usuarios deshabilitados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         TabIndex        =   10
         Top             =   300
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Usuarios.frx":0000
         Left            =   120
         List            =   "Usuarios.frx":0010
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtBuscar 
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
         TabIndex        =   2
         Top             =   1200
         Width           =   5655
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Usuarios.frx":0036
         Left            =   120
         List            =   "Usuarios.frx":0043
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Busqueda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1170
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   1620
      TabIndex        =   5
      Top             =   2160
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuario"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Apellido"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "rol"
         Text            =   "Rol"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contraseña"
         Object.Width           =   2999
      EndProperty
   End
   Begin VB.Label labRegistrar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registrar un usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   4750
      MouseIcon       =   "Usuarios.frx":005F
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   7080
      Width           =   2835
   End
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim deshabilitados As Boolean
Dim a

Function cargarUsuarios()
    If rsUsuarios.EOF Then Exit Function
    
    rsUsuarios.MoveFirst
    
    Dim i As Integer
    i = 1
    While rsUsuarios.EOF = False
        ListView1.ListItems.Add , , Trim(rsUsuarios("Usuario"))
        
        If rsUsuarios.Fields("Nombre") <> Empty Then
            ListView1.ListItems(i).SubItems(1) = Trim(rsUsuarios("Nombre"))
            ListView1.ListItems(i).SubItems(2) = Trim(rsUsuarios("Apellido"))
        End If
        
        If rsUsuarios.Fields("Rol") = 0 Then
            ListView1.ListItems(i).SubItems(3) = "ADMIN"
        Else
            If rsUsuarios.Fields("Rol") = 1 Then
                ListView1.ListItems(i).SubItems(3) = "OPERADOR"
            Else
                If rsUsuarios.Fields("Rol") = 2 Then
                    ListView1.ListItems(i).SubItems(3) = "INVITADO"
                Else
                    If rsUsuarios.Fields("Rol") = 3 Then
                        ListView1.ListItems(i).SubItems(3) = "SOCIO"
                    End If
                End If
            End If
        End If
        
        ListView1.ListItems(i).SubItems(4) = Trim(rsUsuarios("Password"))
        i = i + 1
        rsUsuarios.MoveNext
    Wend

    cargarUsuarios = Null
End Function

Private Sub btnDesh_Click()
    Set rsUsuarios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsUsuarios.Source = "Usuarios"
    rsUsuarios.CursorType = adOpenKeyset
    rsUsuarios.LockType = adLockOptimistic
    
    ListView1.ListItems.Clear

    If deshabilitados = True Then
        deshabilitados = False
        btnDesh.Caption = "Mostrar usuarios habilitados"
        txtBuscar.Enabled = False
        Combo1.Enabled = False
        Combo2.Enabled = False
        btnDeshabilitar.Caption = "Habilitar usuario"
    
        rsUsuarios.Open "SELECT * FROM Usuarios WHERE isHabilitado = 0", cn
        ListView1.ListItems.Clear
        
        a = cargarUsuarios()
    
        cn.Close
        
        If ListView1.ListItems.Count <> 0 Then
            ListView1.SelectedItem.Selected = False
        End If
    Else
        deshabilitados = True
        btnDesh.Caption = "Mostrar usuarios deshabilitados"
        txtBuscar.Enabled = True
        Combo1.Enabled = True
        Combo2.Enabled = True
        btnDeshabilitar.Caption = "Deshabilitar usuario"
    
        rsUsuarios.Open "SELECT * FROM Usuarios WHERE isHabilitado = 1", cn
        ListView1.ListItems.Clear
        
        a = cargarUsuarios()
    
        cn.Close
        
        If ListView1.ListItems.Count <> 0 Then
            ListView1.SelectedItem.Selected = False
        End If
    End If
End Sub

Private Sub btnDeshabilitar_Click()
    'Si no hay ningún elemento seleccionado no elimina nada.
    If ListView1.ListItems.Count <> 0 Then
        If ListView1.SelectedItem.Selected = False Then
            MsgBox ("Debe seleccionar un usuario para realizar esta acción.")
            Exit Sub
        End If
    Else
        MsgBox ("No hay ningun usuario para realizar la acción.")
        Exit Sub
    End If
    
    Dim Username As String
    Username = ListView1.SelectedItem.text
    
    If Username = "ADMIN" Then
        MsgBox ("No se puede deshabilitar al administrador del sistema.")
        GoTo cancelarDeshabilitar
    End If

    If deshabilitados = True Then
        Dim res As Integer
        res = MsgBox("Está por deshabilitar al usuario " & Username & Chr(10) & _
        "Esta acción no le permitirá al usuario entrar al sistema.", vbYesNo, "Confirmar acción")
               
        If res = 6 Then
            Set cn = New ADODB.Connection
            Set rsUsuarios = New ADODB.Recordset
        
            cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
            
            rsUsuarios.Source = "Usuarios"
            rsUsuarios.CursorType = adOpenKeyset
            rsUsuarios.LockType = adLockOptimistic
            rsUsuarios.Open "SELECT * FROM Usuarios WHERE Usuario = '" & Username & "'", cn
            rsUsuarios.MoveFirst
            
            With rsUsuarios
                .Requery
                
                .Fields("isHabilitado") = 0
                    
                .UpdateBatch 'Actualizamos la DB
                .Requery
            End With
            
            cn.Close
            
            GoTo cargarUsuarios
            
        Else: GoTo cancelarDeshabilitar
        
        End If
    Else
        res = MsgBox("Está por habilitar al usuario " & Username & Chr(10) & _
        "Esta acción le permitira al usuario volver a ingresar al sistema ", vbYesNo, "Confirmar acción")
    
        If res = 6 Then
            Set cn = New ADODB.Connection
            Set rsUsuarios = New ADODB.Recordset
        
            cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
            
            rsUsuarios.Source = "Usuarios"
            rsUsuarios.CursorType = adOpenKeyset
            rsUsuarios.LockType = adLockOptimistic
            rsUsuarios.Open "SELECT * FROM Usuarios WHERE Usuario = '" & Username & "'", cn
            rsUsuarios.MoveFirst
            
            With rsUsuarios
                .Requery
                
                .Fields("isHabilitado") = 1
                    
                .UpdateBatch 'Actualizamos la DB
                .Requery
            End With
            
            cn.Close
            
            GoTo cargarUsuarios
        Else: GoTo cancelarDeshabilitar
        End If
    End If

'Deselecciona la última selección.
cancelarDeshabilitar:
    If ListView1.ListItems.Count <> 0 Then
        ListView1.SelectedItem.Selected = False
    End If
    Exit Sub

'Carga los usuarios
cargarUsuarios:
    If deshabilitados = True Then
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
        rsUsuarios.Source = "Usuarios"
        rsUsuarios.CursorType = adOpenKeyset
        rsUsuarios.LockType = adLockOptimistic
        rsUsuarios.Open "SELECT * FROM Usuarios WHERE isHabilitado = 1", cn
        ListView1.ListItems.Clear
        
        a = cargarUsuarios()
    
        cn.Close
        
        If ListView1.ListItems.Count <> 0 Then
            ListView1.SelectedItem.Selected = False
        End If
    Else
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
        rsUsuarios.Source = "Usuarios"
        rsUsuarios.CursorType = adOpenKeyset
        rsUsuarios.LockType = adLockOptimistic
        rsUsuarios.Open "SELECT * FROM Usuarios WHERE isHabilitado = 0", cn
        ListView1.ListItems.Clear
        
        a = cargarUsuarios()
    
        cn.Close
        
        If ListView1.ListItems.Count <> 0 Then
            ListView1.SelectedItem.Selected = False
        End If
    End If
End Sub

Private Sub btnEditar_Click()
    'Si no hay ningún elemento seleccionado no hace nada.
    If ListView1.ListItems.Count <> 0 Then
        If ListView1.SelectedItem.Selected = False Then
            MsgBox ("Debe seleccionar un usuario para realizar esta acción.")
            Exit Sub
        End If
    Else
        MsgBox ("No hay ningun usuario para realizar la acción.")
        Exit Sub
    End If
    
    Dim Rol As String
    Rol = ListView1.SelectedItem.ListSubItems.Item(3).text

    If loginRol = 1 Then
        If Rol = "ADMIN" Then
            MsgBox ("Solo el administrador del sistema puede editarse a si mismo.")
            Exit Sub
        End If
        
        If Rol = "OPERADOR" Then
            MsgBox ("Solo el administrador del sistema puede editar operadores.")
            Exit Sub
        End If
    End If
    
    editUser = ListView1.SelectedItem.text
    
    editando2 = True
    
    Registro.Show 1, Me
    
    If ListView1.ListItems.Count > 0 Then
        ListView1.SelectedItem.Selected = False
    End If
End Sub

Private Sub btnVolver_Click()
    Menu.Show
    Unload Me
End Sub

Private Sub Combo1_Click()
    If Combo1.text = "Rol" Then
        Combo2.Visible = True
        txtBuscar.Visible = False
    Else
        Combo2.Visible = False
        txtBuscar.Visible = True
    End If
End Sub

Private Sub Combo2_Click()
    If Combo2 = Empty Then Exit Sub

    Set rsUsuarios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    Dim Rol As Integer
    
    If Combo2 = "Admin" Then Rol = 0
    If Combo2 = "Operador" Then Rol = 1
    If Combo2 = "Invitado" Then Rol = 2
    If Combo2 = "Socio" Then Rol = 3
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    query = "SELECT * FROM Usuarios"
    
    rsUsuarios.Source = "Usuarios"
    rsUsuarios.CursorType = adOpenKeyset
    rsUsuarios.LockType = adLockOptimistic
    
    query = query & " WHERE Rol = '" & Rol & "' AND isHabilitado = 1"
    
    rsUsuarios.Open query, cn
    ListView1.ListItems.Clear
    
    a = cargarUsuarios()

    cn.Close
End Sub

Private Sub Form_Activate()
    deshabilitados = True
    editando2 = False
    
    If loginRol <> 0 Then btnDeshabilitar.Enabled = False

    Set rsUsuarios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsUsuarios.Source = "Usuarios"
    rsUsuarios.CursorType = adOpenKeyset
    rsUsuarios.LockType = adLockOptimistic
    rsUsuarios.Open "SELECT * FROM Usuarios WHERE isHabilitado = 1", cn
    ListView1.ListItems.Clear
    
    a = cargarUsuarios()

    cn.Close
    
    If ListView1.ListItems.Count > 0 Then
        ListView1.SelectedItem.Selected = False
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labRegistrar.FontUnderline = False
End Sub

Private Sub labRegistrar_Click()
    Registro.Show 1, Me
End Sub

Private Sub labRegistrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labRegistrar.FontUnderline = True
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView1.Sorted = True And ColumnHeader.Index - 1 = ListView1.SortKey Then
        ListView1.SortOrder = 1 - ListView1.SortOrder
    Else
        ListView1.SortOrder = lvwAscending
        ListView1.SortKey = ColumnHeader.Index - 1
    End If
    
    ListView1.Sorted = True
End Sub

Private Sub txtBuscar_Change()
    Set rsUsuarios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    query = "SELECT * FROM Usuarios WHERE isHabilitado = 1"
    
    rsUsuarios.Source = "Usuarios"
    rsUsuarios.CursorType = adOpenKeyset
    rsUsuarios.LockType = adLockOptimistic
    
    If txtBuscar <> "" And Combo1 <> "" Then
        query = query & " AND " & Combo1 & " LIKE '%" & UCase(txtBuscar) & "%'"
    ElseIf Combo1 = "" Then
        Exit Sub
    End If
    
    rsUsuarios.Open query, cn
    ListView1.ListItems.Clear
    
    a = cargarUsuarios()

    cn.Close
End Sub

Private Sub txtBuscar_LostFocus()
    txtBuscar = UCase(txtBuscar)
End Sub
