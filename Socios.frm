VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Socios 
   BackColor       =   &H006E943D&
   Caption         =   "Consulta de socios"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   12450
   Begin VB.CommandButton btnEstado 
      Caption         =   "Cambiar estado"
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
      Left            =   5160
      TabIndex        =   14
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton btnDeshabilitar 
      Caption         =   "Deshabilitar socio"
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
      TabIndex        =   12
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "Editar datos del socio"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   6720
      Width           =   2055
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
      Height          =   2415
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton btnMostrarDesh 
         Caption         =   "Mostrar socios deshabilitados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox checkMorosos 
         BackColor       =   &H006E943D&
         Caption         =   "No mostrar morosos"
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
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Value           =   1  'Checked
         Width           =   4575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H006E943D&
         Caption         =   "Género"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   5640
         TabIndex        =   6
         Top             =   300
         Width           =   1455
         Begin VB.OptionButton optGen 
            BackColor       =   &H006E943D&
            Caption         =   "Varón"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optGen 
            BackColor       =   &H006E943D&
            Caption         =   "Mujer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optGen 
            BackColor       =   &H006E943D&
            Caption         =   "Ambos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Value           =   -1  'True
            Width           =   1215
         End
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
         ItemData        =   "Socios.frx":0000
         Left            =   120
         List            =   "Socios.frx":000D
         TabIndex        =   5
         Top             =   360
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
         TabIndex        =   4
         Top             =   1200
         Width           =   5295
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
      Left            =   8880
      TabIndex        =   1
      Top             =   6720
      Width           =   2775
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   11055
      _ExtentX        =   19500
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Apellido"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Género"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dirección"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Localidad"
         Object.Width           =   2788
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Provincia"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "DNI"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Estado"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "Socios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim deshabilitados As Boolean

Function eliminarReservas(Nombre)
    rsReservas.Source = "Reservas"
    rsReservas.CursorType = adOpenKeyset
    rsReservas.LockType = adLockOptimistic
    rsReservas.Open "SELECT * FROM Reservas WHERE Fecha LIKE '%" & Date & "%'", cn
    
    If rsReservas.EOF = False Then
        rsReservas.MoveFirst
        
        rsHorarios.Source = "Horarios"
        rsHorarios.CursorType = adOpenKeyset
        rsHorarios.LockType = adLockOptimistic
        
        Dim reservante As String
        Dim user As String
        Dim cancha As String
        Dim turno As String
        
        Do Until rsReservas.EOF
            If DateDiff("h", Time, Mid(rsReservas.Fields("Turno"), 1, 5)) > 0 Then
                If rsReservas.Fields("Jugador1A") = Nombre Or rsReservas.Fields("Jugador1B") = Nombre Or _
                rsReservas.Fields("Jugador2A") = Nombre Or rsReservas.Fields("Jugador2B") = Nombre Then
                    reservante = rsReservas.Fields("Jugador1A")
                    cancha = UCase(Mid(rsReservas.Fields("Cancha"), 1, 6)) & _
                             Mid(Trim(rsReservas.Fields("Cancha")), 11, Len(rsReservas.Fields("Cancha")))
                    turno = rsReservas.Fields("Turno")
                    
                    rsReservas.Delete
                    rsReservas.Update
                    
                    rsSocios.MoveFirst
                    
                    Do Until rsSocios.EOF
                        If (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido")) = reservante Then
                            user = rsSocios.Fields("Usuario")
                            
                            With rsSocios
                                .Requery
                                .Find "Usuario = '" & user & "'"
                                
                                .Fields("ultimaReserva") = .Fields("reservaAux")
                        
                                .UpdateBatch
                                .Requery
                            End With
                            
                            Exit Do
                        End If
                        rsSocios.MoveNext
                    Loop
                    
                    rsHorarios.Open "SELECT * FROM Horarios WHERE Hora = '" & turno & "'", cn
                    rsHorarios.MoveFirst
                    
                    With rsHorarios
                        .Requery
                        
                        .Fields(cancha) = 0
                            
                        .UpdateBatch 'Actualizamos la DB
                        .Requery
                    End With
                    
                    rsHorarios.Close
                End If
            End If
            rsReservas.Update
            If rsReservas.EOF = False Then
                rsReservas.MoveNext
            Else: Exit Do
            End If
        Loop
    End If
    
    eliminarReservas = Null
End Function

Private Sub btnDeshabilitar_Click()
    'Si no hay ningún elemento seleccionado no elimina nada.
    If ListView1.ListItems.Count <> 0 Then
        If ListView1.SelectedItem.Selected = False Then
            MsgBox ("Debe seleccionar un socio para realizar esta acción.")
            Exit Sub
        End If
    Else
        MsgBox ("No hay ningun socio para realizar la acción.")
        Exit Sub
    End If

    Dim DNI As Long
    Dim Nombre As String
    Dim Username As String
    
    'Asigno los valores del elemento seleccionado del ListView para hacer el código más sencillo.
    DNI = ListView1.SelectedItem.ListSubItems.Item(6).text
    Nombre = ListView1.SelectedItem.ListSubItems.Item(1).text & " " & ListView1.SelectedItem.text

    If deshabilitados = True Then
        Dim res As Integer
        res = MsgBox("Está por deshabilitar al socio " & Nombre & Chr(10) & "DNI: " & DNI & Chr(10) & Chr(10) & _
        "Esta acción también deshabilitara al usuario asociado al socio y cualquier reserva actual del mismo.", vbYesNo, "Confirmar acción")
               
        If res = 6 Then
            Set cn = New ADODB.Connection
            Set rsReservas = New ADODB.Recordset
            Set rsSocios = New ADODB.Recordset
            Set rsHorarios = New ADODB.Recordset
            Set rsUsuarios = New ADODB.Recordset
        
            cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
            
            rsSocios.Source = "Socios"
            rsSocios.CursorType = adOpenKeyset
            rsSocios.LockType = adLockOptimistic
            rsSocios.Open "SELECT * FROM Socios WHERE DNI = " & DNI & "", cn
            rsSocios.MoveFirst
            
            Username = rsSocios.Fields("Usuario")
            
            With rsSocios
                .Requery
                
                .Fields("isHabilitado") = 0
                
                    
                .UpdateBatch 'Actualizamos la DB
                .Requery
            End With
            
            rsSocios.Close
            rsSocios.Open "SELECT * FROM Socios", cn
            
            Dim a
            a = eliminarReservas(Nombre)
            
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
            
            GoTo cargarSocios
            
        Else: GoTo cancelarDeshabilitar
        
        End If
    Else
        res = MsgBox("Está por habilitar al socio " & Nombre & Chr(10) & "DNI: " & DNI & Chr(10) & Chr(10) & _
        "Esta acción le permitira al socio volver a usar el sistema de reservas de canchas.", vbYesNo, "Confirmar acción")
    
        If res = 6 Then
            Set cn = New ADODB.Connection
            Set rsSocios = New ADODB.Recordset
            Set rsUsuarios = New ADODB.Recordset
        
            cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
            
            rsSocios.Source = "Socios"
            rsSocios.CursorType = adOpenKeyset
            rsSocios.LockType = adLockOptimistic
            rsSocios.Open "SELECT * FROM Socios WHERE DNI = " & DNI & "", cn
            rsSocios.MoveFirst
            
            Username = rsSocios.Fields("Usuario")
            
            With rsSocios
                .Requery
                
                .Fields("isHabilitado") = 1
                
                    
                .UpdateBatch 'Actualizamos la DB
                .Requery
            End With
            
            rsSocios.Close
            
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
            
            GoTo cargarSocios
        Else: GoTo cancelarDeshabilitar
        End If
    End If

'Deselecciona la última selección.
cancelarDeshabilitar:
    If ListView1.ListItems.Count <> 0 Then
            ListView1.SelectedItem.Selected = False
        End If
    Exit Sub

'Carga los socios
cargarSocios:
    If deshabilitados = True Then
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE Estado = 'OK' AND isHabilitado = 1", cn
        rsSocios.MoveFirst
        
        ListView1.ListItems.Clear
        Dim li As ListItem
        While rsSocios.EOF = False
            Set li = ListView1.ListItems.Add(, , Trim(rsSocios("Apellido")))
            If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
            If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
            If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
            li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
            If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
            If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
            If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
            If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
            rsSocios.MoveNext
        Wend
        
        cn.Close
        If ListView1.ListItems.Count <> 0 Then
            ListView1.SelectedItem.Selected = False
        End If
    Else
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE isHabilitado = 0", cn
        ListView1.ListItems.Clear
        
        If rsSocios.EOF = False Then
            rsSocios.MoveFirst
            Do Until rsSocios.EOF
                Set li = ListView1.ListItems.Add(, , Trim(rsSocios.Fields("Apellido")))
                If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
                If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
                If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
                li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
                If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
                If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
                If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
                If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
                rsSocios.MoveNext
            Loop
        End If

        cn.Close
        If ListView1.ListItems.Count <> 0 Then
            ListView1.SelectedItem.Selected = False
        End If
    End If
End Sub

Private Sub btnEditar_Click()
    'Si no hay ningún elemento seleccionado tira un error.
    If ListView1.SelectedItem.Selected = False Then
        MsgBox ("Debe seleccionar un socio para editar.")
        Exit Sub
    End If

    Dim Nombre As String
    Dim Username As String
    
    'Asigno los valores del elemento seleccionado del ListView para hacer el código más sencillo.
    dniUser = ListView1.SelectedItem.ListSubItems.Item(6).text
    Nombre = ListView1.SelectedItem.ListSubItems.Item(1).text & " " & ListView1.SelectedItem.text
    
    editando = True
    
    Registro.Show 1, Me
    ListView1.SelectedItem.Selected = False
End Sub

Private Sub btnEstado_Click()
    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset
    Set rsReservas = New ADODB.Recordset
    Set rsHorarios = New ADODB.Recordset
    
    Dim DNI As Long
    Dim Nombre As String
    DNI = ListView1.SelectedItem.ListSubItems.Item(6).text
    Nombre = ListView1.SelectedItem.ListSubItems.Item(1).text & " " & ListView1.SelectedItem.text

    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open "SELECT * FROM Socios WHERE DNI = " & DNI & "", cn
    
    rsSocios.Requery
    
    If ListView1.SelectedItem.ListSubItems.Item(7).text = "OK" Then
        rsSocios.Fields("Estado") = "EN DEUDA"
        Dim a
        a = eliminarReservas(Nombre)
    Else
        rsSocios.Fields("Estado") = "OK"
    End If
    
    rsSocios.UpdateBatch
    rsSocios.Requery
    
    rsSocios.Close
    
    If checkMorosos.Value = 1 Then
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE Estado = 'OK' AND isHabilitado = 1", cn
        rsSocios.MoveFirst
        
        ListView1.ListItems.Clear
        
        Dim li As ListItem
        While rsSocios.EOF = False
            Set li = ListView1.ListItems.Add(, , Trim(rsSocios("Apellido")))
            If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
            If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
            If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
            li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
            If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
            If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
            If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
            If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
            rsSocios.MoveNext
        Wend
    Else
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE Estado IS NOT null AND isHabilitado = 1", cn
        
        ListView1.ListItems.Clear
        
        If rsSocios.EOF = False Then
            rsSocios.MoveFirst
            Do Until rsSocios.EOF
                Set li = ListView1.ListItems.Add(, , Trim(rsSocios.Fields("Apellido")))
                If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
                If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
                If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
                li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
                If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
                If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
                If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
                If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
                rsSocios.MoveNext
            Loop
        End If
    
        cn.Close
    End If
    
    If ListView1.ListItems.Count > 0 Then
        ListView1.SelectedItem.Selected = False
    End If
End Sub

Private Sub btnMostrarDesh_Click()
    ListView1.ListItems.Clear

    Set rsSocios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    If deshabilitados = True Then
        query = "SELECT * FROM Socios WHERE isHabilitado = 0"
        deshabilitados = False
        btnMostrarDesh.Caption = "Mostrar socios habilitados"
        btnDeshabilitar.Caption = "Habilitar socio"
        Frame1.Enabled = False
        txtBuscar.Enabled = False
        Combo1.Enabled = False
        checkMorosos.Enabled = False
    Else
        query = "SELECT * FROM Socios WHERE Estado = 'OK' AND isHabilitado = 1"
        deshabilitados = True
        btnMostrarDesh.Caption = "Mostrar socios deshabilitados"
        btnDeshabilitar.Caption = "Deshabilitar socio"
        Frame1.Enabled = True
        txtBuscar.Enabled = True
        Combo1.Enabled = True
        checkMorosos.Enabled = True
        checkMorosos.Value = 1
    End If
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open query, cn
    ListView1.ListItems.Clear
    
    Dim li As ListItem
    If rsSocios.EOF = False Then
        rsSocios.MoveFirst
        Do Until rsSocios.EOF
            Set li = ListView1.ListItems.Add(, , Trim(rsSocios.Fields("Apellido")))
            If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
            If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
            If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
            li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
            If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
            If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
            If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
            If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
            rsSocios.MoveNext
        Loop
    End If

    cn.Close
End Sub

Private Sub checkMorosos_Click()
    Set rsSocios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    query = "SELECT * FROM Socios"
    
    If checkMorosos Then
        query = query & " WHERE Estado = 'OK'"
    Else
        query = query & " WHERE Estado IS NOT null"
    End If
    
    If optGen(0) Then
         query = query & " AND Género = '" & optGen(0).Caption & "'"
    ElseIf optGen(1) Then
        query = query & " AND Género = '" & optGen(1).Caption & "'"
    End If
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    
    If txtBuscar <> "" And Combo1 <> "" Then
        query = query & " AND " & Combo1 & " LIKE '%" & txtBuscar & "%' AND isHabilitado = 1"
    End If
    
    rsSocios.Open query, cn
    ListView1.ListItems.Clear
    
    Dim li As ListItem
    If rsSocios.EOF = False Then
        rsSocios.MoveFirst
        Do Until rsSocios.EOF
            Set li = ListView1.ListItems.Add(, , Trim(rsSocios.Fields("Apellido")))
            If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
            If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
            If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
            li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
            If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
            If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
            If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
            If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
            rsSocios.MoveNext
        Loop
    End If

    cn.Close
End Sub

Private Sub Combo1_Change()
    Set rsSocios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    query = "SELECT * FROM Socios"
    
    If checkMorosos Then
        query = query & " WHERE Estado = 'OK'"
    Else
        query = query & " WHERE Estado IS NOT null"
    End If
    
    If optGen(0) Then
         query = query & " AND Género = '" & optGen(0).Caption & "'"
    ElseIf optGen(1) Then
        query = query & " AND Género = '" & optGen(1).Caption & "'"
    End If
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    
    If txtBuscar <> "" And Combo1 <> "" Then
        query = query & " AND " & Combo1 & " LIKE '%" & txtBuscar & "%' AND isHabilitado = 1"
    End If
    
    rsSocios.Open query, cn
    ListView1.ListItems.Clear
    
    Dim li As ListItem
    If rsSocios.EOF = False Then
        rsSocios.MoveFirst
        Do Until rsSocios.EOF
            Set li = ListView1.ListItems.Add(, , Trim(rsSocios.Fields("Apellido")))
            If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
            If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
            If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
            li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
            If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
            If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
            If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
            If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
            rsSocios.MoveNext
        Loop
    End If

    cn.Close
End Sub

Private Sub optGen_Click(Index As Integer)
    Set rsSocios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    query = "SELECT * FROM Socios"
    
    If checkMorosos Then
        query = query & " WHERE Estado = 'OK'"
    Else
        query = query & " WHERE Estado IS NOT null"
    End If
    
    If Index <> 2 Then
         query = query & " AND Género = '" & optGen(Index).Caption & "'"
    End If
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    
    If txtBuscar <> "" And Combo1 <> "" Then
        query = query & " AND " & Combo1 & " LIKE '%" & txtBuscar & "%' AND isHabilitado = 1"
    End If
    
    rsSocios.Open query, cn
    ListView1.ListItems.Clear
    
    Dim li As ListItem
    If rsSocios.EOF = False Then
        rsSocios.MoveFirst
        Do Until rsSocios.EOF
            Set li = ListView1.ListItems.Add(, , Trim(rsSocios.Fields("Apellido")))
            If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
            If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
            If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
            li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
            If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
            If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
            If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
            If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
            rsSocios.MoveNext
        Loop
    End If

    cn.Close
End Sub

Private Sub txtBuscar_Change()
    Set rsSocios = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    query = "SELECT * FROM Socios"
    
    If checkMorosos Then
        query = query & " WHERE Estado = 'OK'"
    Else
        query = query & " WHERE Estado IS NOT null"
    End If
    
    If optGen(0) Then
        query = query & " AND Género = '" & optGen(0).Caption & "'"
    ElseIf optGen(1) Then
        query = query & " AND Género = '" & optGen(1).Caption & "'"
    End If
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    
    If txtBuscar <> "" And Combo1 <> "" Then
        query = query & " AND " & Combo1 & " LIKE '%" & txtBuscar & "%' AND isHabilitado = 1"
    End If
    
    rsSocios.Open query, cn
    ListView1.ListItems.Clear
    
    Dim li As ListItem
    If rsSocios.EOF = False Then
        rsSocios.MoveFirst
        Do Until rsSocios.EOF
            Set li = ListView1.ListItems.Add(, , Trim(rsSocios.Fields("Apellido")))
            If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
            If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
            If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
            li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
            If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
            If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
            If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
            If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
            rsSocios.MoveNext
        Loop
    End If

    cn.Close
End Sub

Private Sub btnVolver_Click()
    Menu.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    editando = False

    If loginRol = 3 Or loginRol = 2 Then btnEditar.Enabled = False
    If loginRol = 3 Then btnEstado.Enabled = False
    If loginRol <> 0 Then btnDeshabilitar.Enabled = False

    deshabilitados = True

    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open "SELECT * FROM Socios WHERE Estado = 'OK' AND isHabilitado = 1", cn
    rsSocios.MoveFirst
    
    Dim li As ListItem
    While rsSocios.EOF = False
        Set li = ListView1.ListItems.Add(, , Trim(rsSocios("Apellido")))
        If Not Len(Trim(rsSocios.Fields("Nombre"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Nombre"))
        If Not Len(Trim(rsSocios.Fields("Género"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Género"))
        If Not (Len(Trim(rsSocios.Fields("Calle"))) = 0 And Len(Trim(rsSocios.Fields("NroCalle") = 0))) Then _
        li.ListSubItems.Add , , Trim(rsSocios("Calle")) & " " & Trim(rsSocios(("NroCalle")))
        If Not Len(Trim(rsSocios.Fields("Localidad"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Localidad"))
        If Not Len(Trim(rsSocios.Fields("Provincia"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Provincia"))
        If Not Len(Trim(rsSocios.Fields("DNI"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("DNI"))
        If Not Len(Trim(rsSocios.Fields("Estado"))) = 0 Then li.ListSubItems.Add , , Trim(rsSocios("Estado"))
        rsSocios.MoveNext
    Wend
    
    If ListView1.ListItems.Count <> 0 Then
        ListView1.SelectedItem.Selected = False
    End If
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

Private Sub txtBuscar_LostFocus()
    txtBuscar.text = UCase(txtBuscar)
End Sub
