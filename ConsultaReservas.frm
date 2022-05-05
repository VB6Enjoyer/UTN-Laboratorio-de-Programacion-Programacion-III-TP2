VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConsultaReservas 
   BackColor       =   &H006E943D&
   Caption         =   "Consulta de reservas"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   12450
   Begin VB.CommandButton btnEditar 
      Caption         =   "Editar reserva"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton btnEliminar 
      Caption         =   "Eliminar reserva"
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
      Left            =   240
      TabIndex        =   9
      Top             =   6960
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
      Left            =   9360
      TabIndex        =   4
      Top             =   6960
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
      Height          =   2655
      Left            =   1620
      TabIndex        =   0
      Top             =   120
      Width           =   9135
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
         ItemData        =   "ConsultaReservas.frx":0000
         Left            =   120
         List            =   "ConsultaReservas.frx":000D
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSComCtl2.MonthView MonthView2 
         Height          =   2370
         Left            =   6240
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   7246909
         Appearance      =   1
         StartOfWeek     =   57475074
         CurrentDate     =   44362
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   7246909
         Appearance      =   1
         StartOfWeek     =   57475074
         CurrentDate     =   44362
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
         Width           =   4575
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
         ItemData        =   "ConsultaReservas.frx":0032
         Left            =   120
         List            =   "ConsultaReservas.frx":0045
         TabIndex        =   1
         Top             =   360
         Width           =   2535
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
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Cancha"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Turno"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jugador 1 (Reservante)"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Compañero"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Jugador 2"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Compañero"
         Object.Width           =   3704
      EndProperty
   End
End
Attribute VB_Name = "ConsultaReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEditar_Click()
    Dim fecha As Date
    Dim hora As String
    Dim reservante As String
    
    'Asigno los valores del elemento seleccionado del ListView para hacer el código más sencillo.
    fecha = ListView1.SelectedItem.text
    hora = ListView1.SelectedItem.ListSubItems.Item(2).text
    reservante = ListView1.SelectedItem.ListSubItems.Item(3).text
    
    'Verificaciones
   '|----------------------------------------------------------------------|
    'Si no hay ningún elemento seleccionado le pide al usuario elegir una reserva.
    If ListView1.SelectedItem.Selected = False Then
        MsgBox ("Debe seleccionar una reserva para editar.")
        Exit Sub
    End If
    
    'No permite editar reservas de días pasados.
    If fecha <> Date Then
        MsgBox ("No puede editar reservas de días anteriores.")
        GoTo cancelarEditar
    End If
    
    'No permite editar reservas de horas pasadas en el día corriente.
    If DateDiff("h", Time, Mid(hora, 1, 5)) <= 0 Then
        MsgBox ("No puede editar reservas ya ocurridas.")
        GoTo cancelarEditar
    End If
   '|----------------------------------------------------------------------|
    
    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open "SELECT * FROM Socios", cn
    rsSocios.MoveFirst
    
    Do Until rsSocios.EOF
        If rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = reservante Then
            editUser = rsSocios.Fields("Usuario")
            Exit Do
        End If
        rsSocios.MoveNext
    Loop
    
    editando = True
    
    Reservar.Show
    Unload Me
    
    Exit Sub
    
cancelarEditar:
    ListView1.SelectedItem.Selected = False
End Sub

Private Sub btnEliminar_Click()
    Dim Index As Integer
    Dim cancha As String
    Dim fecha As Date
    Dim hora As String
    Dim reservante As String
    
    'Asigno los valores del elemento seleccionado del ListView para hacer el código más sencillo.
    Index = ListView1.SelectedItem.Index
    fecha = ListView1.SelectedItem.text
    cancha = "CANCHA" & ListView1.SelectedItem.ListSubItems.Item(1).text
    hora = ListView1.SelectedItem.ListSubItems.Item(2).text
    reservante = ListView1.SelectedItem.ListSubItems.Item(3).text
    
    'Verificaciones
   '|----------------------------------------------------------------------|
    'Si no hay ningún elemento seleccionado no elimina nada.
    If ListView1.SelectedItem.Selected = False Then
        MsgBox ("Debe seleccionar una reserva para eliminar.")
        Exit Sub
    End If
    
    'No permite eliminar reservas de días pasados.
    If fecha <> Date Then
        MsgBox ("No puede eliminar reservas de días anteriores.")
        GoTo cancelarEliminar
    End If
    
    'No permite eliminar reservas de horas pasadas en el día corriente.
    If DateDiff("h", Time, Mid(hora, 1, 5)) <= 0 Then
        MsgBox ("No puede eliminar reservas ya ocurridas.")
        GoTo cancelarEliminar
    End If
   '|----------------------------------------------------------------------|
   
    Dim res As Integer
    res = MsgBox("Está seguro que desea eliminar la siguiente reserva?" & Chr(10) & Chr(10) & _
           fecha & Chr(10) & cancha & " entre " & hora & Chr(10) & _
           "Reserva hecha por: " & reservante, vbYesNo, "Confirmar eliminación de reserva")
           
    If res = 6 Then
        Set cn = New ADODB.Connection
        Set rsReservas = New ADODB.Recordset
        Set rsSocios = New ADODB.Recordset
        Set rsHorarios = New ADODB.Recordset
    
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
            
        rsReservas.Source = "Reservas"
        rsReservas.CursorType = adOpenKeyset
        rsReservas.LockType = adLockOptimistic
        
        'Buscamos en la DB la reserva seleccionado al igualar la cancha, turno y fecha.
        rsReservas.Open "SELECT * FROM Reservas WHERE Cancha = '" & cancha & "' AND Turno = '" & hora & "' AND Fecha LIKE '%" & fecha & "%'", cn
        
        'Eliminamos la entrada encontrada y actualizamos la DB
        rsReservas.Delete
        rsReservas.Update
        
        'Cierro la conexión y abro una nueva para volver a generar la tabla
        rsReservas.Close
        rsReservas.Open "SELECT * FROM Reservas ORDER BY Fecha DESC, Turno DESC", cn
        
        If (rsReservas.EOF = True And rsReservas.BOF = False) Then Exit Sub
        
        rsReservas.MoveFirst
        
        ListView1.ListItems.Clear
        Dim li As ListItem
        While rsReservas.EOF = False
            Set li = ListView1.ListItems.Add(, , Trim(rsReservas("Fecha")))
            li.ListSubItems.Add , , Mid(Trim(rsReservas("Cancha")), 7, Len(rsReservas("Cancha")))
            li.ListSubItems.Add , , Trim(rsReservas("Turno"))
            li.ListSubItems.Add , , Trim(rsReservas("Jugador1A"))
            
            If rsReservas.Fields("Jugador1B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador1B"))
            
            li.ListSubItems.Add , , Trim(rsReservas("Jugador2A"))
            
            If rsReservas.Fields("Jugador2B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador2B"))
            
            rsReservas.MoveNext
        Wend
        
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE ultimaReserva LIKE '%" & fecha & "%'", cn
        rsSocios.MoveFirst
        
        Do Until rsSocios.EOF
            If rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = reservante Then
                'Reemplazamos la fecha de la última reserva por un guardado auxiliar para permitir al socio reservar
                'de nuevo en el día
                With rsSocios
                    .Requery
                    
                    .Fields("ultimaReserva") = .Fields("reservaAux")
                        
                    .UpdateBatch 'Actualizamos la DB
                    .Requery
                End With
                
                Exit Do
            End If
            rsSocios.MoveNext
        Loop
        
        rsHorarios.Source = "horarios"
        rsHorarios.CursorType = adOpenKeyset
        rsHorarios.LockType = adLockOptimistic
        rsHorarios.Open "SELECT * FROM horarios WHERE Hora = '" & hora & "'", cn
        rsHorarios.MoveFirst
        
        With rsHorarios
            .Requery
            
            .Fields("Cancha" & Mid(Trim(cancha), 11, Len(cancha))) = 0
                
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With
        
        cn.Close
        
        ListView1.SelectedItem.Selected = False
        
    Else: GoTo cancelarEliminar
    
    End If

'Deselecciona la última selección.
cancelarEliminar:
    ListView1.SelectedItem.Selected = False
End Sub

Private Sub btnVolver_Click()
    Menu.Show
    Unload Me
End Sub

Private Sub Combo1_Click()
    If Combo1 = "Fecha" Then
        txtBuscar.Visible = False
        MonthView1.Visible = True
        MonthView2.Visible = False
        Combo2.Visible = False
        Label2.Visible = False
    Else
        If Combo1 = "Periodo (Todos)" Then
            txtBuscar.Visible = False
            MonthView1.Visible = True
            MonthView2.Visible = True
            Combo2.Visible = False
            Label2.Visible = False
        Else
            If Combo1 = "Periodo (Individual)" Then
                txtBuscar.Visible = False
                MonthView1.Visible = True
                MonthView2.Visible = True
                Combo2.Visible = True
                Label2.Visible = True
                Label2.Caption = "Socio:"
            Else
                txtBuscar.Visible = True
                MonthView1.Visible = False
                MonthView2.Visible = False
                Combo2.Visible = False
                Label2.Visible = True
                Label2.Caption = "Busqueda:"
            End If
        End If
    End If
End Sub

Private Sub Combo2_Click()
    Set rsReservas = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    query = "SELECT * FROM Reservas WHERE Jugador1A = '" & Combo2.text & "'"
    
    rsReservas.Source = "Reservas"
    rsReservas.CursorType = adOpenKeyset
    rsReservas.LockType = adLockOptimistic
    
    rsReservas.Open query & " ORDER BY Fecha DESC, Turno DESC", cn
    
    If (rsReservas.EOF = True And rsReservas.BOF = False) Then Exit Sub
    
    ListView1.ListItems.Clear
    
    Dim li As ListItem
    If rsReservas.EOF = False Then
        rsReservas.MoveFirst
        Do Until rsReservas.EOF
            If rsReservas.Fields("Fecha") >= MonthView1.Value And rsReservas.Fields("Fecha") <= MonthView2.Value Then
                Set li = ListView1.ListItems.Add(, , Trim(rsReservas("Fecha")))
                li.ListSubItems.Add , , Mid(Trim(rsReservas("Cancha")), 7, Len(rsReservas("Cancha")))
                li.ListSubItems.Add , , Trim(rsReservas("Turno"))
                li.ListSubItems.Add , , Trim(rsReservas("Jugador1A"))
                
                If rsReservas.Fields("Jugador1B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador1B"))
                
                li.ListSubItems.Add , , Trim(rsReservas("Jugador2A"))
                
                If rsReservas.Fields("Jugador2B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador2B"))
            End If
            rsReservas.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Activate()
    If loginRol = 3 Or loginRol = 2 Then btnEditar.Enabled = False
    If loginRol <> 0 Then btnEliminar.Enabled = False

    Set cn = New ADODB.Connection
    Set rsReservas = New ADODB.Recordset
    Set rsSocios = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsReservas.Source = "Reservas"
    rsReservas.CursorType = adOpenKeyset
    rsReservas.LockType = adLockOptimistic
    rsReservas.Open "SELECT * FROM Reservas ORDER BY Fecha DESC, Turno DESC", cn
    
    If (rsReservas.EOF = True And rsReservas.BOF = False) Then Exit Sub
    
    rsReservas.MoveFirst
    
    Dim li As ListItem
    While rsReservas.EOF = False
        Set li = ListView1.ListItems.Add(, , Trim(rsReservas("Fecha")))
        li.ListSubItems.Add , , Mid(Trim(rsReservas("Cancha")), 7, Len(rsReservas("Cancha")))
        li.ListSubItems.Add , , Trim(rsReservas("Turno"))
        li.ListSubItems.Add , , Trim(rsReservas("Jugador1A"))
        
        If rsReservas.Fields("Jugador1B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador1B"))
        
        li.ListSubItems.Add , , Trim(rsReservas("Jugador2A"))
        
        If rsReservas.Fields("Jugador2B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador2B"))
        
        rsReservas.MoveNext
    Wend
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open "SELECT * FROM Socios", cn
    rsSocios.MoveFirst
    
    Dim i As Integer
    Do Until rsSocios.EOF
        Combo2.List(i) = (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido"))
        i = i + 1
        rsSocios.MoveNext
    Loop
    
    cn.Close
    
    ListView1.SelectedItem.Selected = False
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

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Set rsReservas = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    If Combo1 = "Fecha" Then
        query = "SELECT * FROM Reservas WHERE Fecha LIKE '%" & MonthView1.Value & "%'"
        
        rsReservas.Source = "Reservas"
        rsReservas.CursorType = adOpenKeyset
        rsReservas.LockType = adLockOptimistic
        
        rsReservas.Open query & " ORDER BY Fecha DESC, Turno DESC", cn
        
        If (rsReservas.EOF = True And rsReservas.BOF = False) Then Exit Sub
        
        ListView1.ListItems.Clear
        
        Dim li As ListItem
        If rsReservas.EOF = False Then
            rsReservas.MoveFirst
            Do Until rsReservas.EOF
                Set li = ListView1.ListItems.Add(, , Trim(rsReservas("Fecha")))
                li.ListSubItems.Add , , Mid(Trim(rsReservas("Cancha")), 7, Len(rsReservas("Cancha")))
                li.ListSubItems.Add , , Trim(rsReservas("Turno"))
                li.ListSubItems.Add , , Trim(rsReservas("Jugador1A"))
                
                If rsReservas.Fields("Jugador1B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador1B"))
                
                li.ListSubItems.Add , , Trim(rsReservas("Jugador2A"))
                
                If rsReservas.Fields("Jugador2B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador2B"))
                
                rsReservas.MoveNext
            Loop
        End If
    End If

    If Combo1 = "Periodo (Todos)" Or Combo1 = "Periodo (Individual)" Then
        If MonthView1.Value > MonthView2.Value Then
            MonthView1.Value = MonthView2.Value
        End If
        
        If Combo1 = "Periodo (Todos)" Then
            query = "SELECT * FROM Reservas"
        Else
            query = "SELECT * FROM Reservas WHERE Jugador1A = '" & Combo2.text & "'"
        End If
        
        rsReservas.Source = "Reservas"
        rsReservas.CursorType = adOpenKeyset
        rsReservas.LockType = adLockOptimistic
        
        rsReservas.Open query & " ORDER BY Fecha DESC, Turno DESC", cn
        
        If (rsReservas.EOF = True And rsReservas.BOF = False) Then Exit Sub
        
        ListView1.ListItems.Clear
        
        If rsReservas.EOF = False Then
            rsReservas.MoveFirst
            Do Until rsReservas.EOF
                If rsReservas.Fields("Fecha") >= MonthView1.Value And rsReservas.Fields("Fecha") <= MonthView2.Value Then
                    Set li = ListView1.ListItems.Add(, , Trim(rsReservas("Fecha")))
                    li.ListSubItems.Add , , Mid(Trim(rsReservas("Cancha")), 7, Len(rsReservas("Cancha")))
                    li.ListSubItems.Add , , Trim(rsReservas("Turno"))
                    li.ListSubItems.Add , , Trim(rsReservas("Jugador1A"))
                    
                    If rsReservas.Fields("Jugador1B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador1B"))
                    
                    li.ListSubItems.Add , , Trim(rsReservas("Jugador2A"))
                    
                    If rsReservas.Fields("Jugador2B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador2B"))
                End If
                rsReservas.MoveNext
            Loop
        End If
    End If
    
    cn.Close
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
    If MonthView2.Value < MonthView1.Value Then
        MonthView2.Value = MonthView1.Value
    End If
    
    Set rsReservas = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"

    If Combo1 = "Periodo (Todos)" Or Combo1 = "Periodo (Individual)" Then
        If MonthView1.Value > MonthView2.Value Then
            MonthView1.Value = MonthView2.Value
        End If
        
        If Combo1 = "Periodo (Todos)" Then
            query = "SELECT * FROM Reservas"
        Else
            query = "SELECT * FROM Reservas WHERE Jugador1A = '" & Combo2.text & "'"
        End If
        
        rsReservas.Source = "Reservas"
        rsReservas.CursorType = adOpenKeyset
        rsReservas.LockType = adLockOptimistic
        
        rsReservas.Open query & " ORDER BY Fecha DESC, Turno DESC", cn
        
        If (rsReservas.EOF = True And rsReservas.BOF = False) Then Exit Sub
        
        ListView1.ListItems.Clear
        
        Dim li As ListItem
        If rsReservas.EOF = False Then
            rsReservas.MoveFirst
            Do Until rsReservas.EOF
                If rsReservas.Fields("Fecha") >= MonthView1.Value And rsReservas.Fields("Fecha") <= MonthView2.Value Then
                    Set li = ListView1.ListItems.Add(, , Trim(rsReservas("Fecha")))
                    li.ListSubItems.Add , , Mid(Trim(rsReservas("Cancha")), 7, Len(rsReservas("Cancha")))
                    li.ListSubItems.Add , , Trim(rsReservas("Turno"))
                    li.ListSubItems.Add , , Trim(rsReservas("Jugador1A"))
                    
                    If rsReservas.Fields("Jugador1B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador1B"))
                    
                    li.ListSubItems.Add , , Trim(rsReservas("Jugador2A"))
                    
                    If rsReservas.Fields("Jugador2B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador2B"))
                End If
                rsReservas.MoveNext
            Loop
        End If
    End If
    
    cn.Close
End Sub

Private Sub txtBuscar_Change()
    Set rsReservas = New ADODB.Recordset
    Set cn = New ADODB.Connection
    
    Dim query As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    query = "SELECT * FROM Reservas"
    
    rsReservas.Source = "Reservas"
    rsReservas.CursorType = adOpenKeyset
    rsReservas.LockType = adLockOptimistic
    
    If txtBuscar <> "" And Combo1 <> "" Then
        If Combo1 = "Socio reservante" Then
            query = query & " WHERE Jugador1A" & " LIKE '%" & txtBuscar & "%'"
        End If
        
        If Combo1 = "Cancha" Then
            query = query & " WHERE Cancha" & " LIKE '%" & "Cancha N° " & txtBuscar & "%'"
        End If
    End If
    
    rsReservas.Open query & " ORDER BY Fecha DESC, Turno DESC", cn
    
    If (rsReservas.EOF = True And rsReservas.BOF = False) Then Exit Sub
    
    ListView1.ListItems.Clear
    
    Dim li As ListItem
    If rsReservas.EOF = False Then
        rsReservas.MoveFirst
        Do Until rsReservas.EOF
            Set li = ListView1.ListItems.Add(, , Trim(rsReservas("Fecha")))
            li.ListSubItems.Add , , Mid(Trim(rsReservas("Cancha")), 7, Len(rsReservas("Cancha")))
            li.ListSubItems.Add , , Trim(rsReservas("Turno"))
            li.ListSubItems.Add , , Trim(rsReservas("Jugador1A"))
            
            If rsReservas.Fields("Jugador1B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador1B"))
            
            li.ListSubItems.Add , , Trim(rsReservas("Jugador2A"))
            
            If rsReservas.Fields("Jugador2B") <> Empty Then li.ListSubItems.Add , , Trim(rsReservas("Jugador2B"))
            
            rsReservas.MoveNext
        Loop
    End If

    cn.Close
End Sub

Private Sub txtBuscar_LostFocus()
    txtBuscar = UCase(txtBuscar)
End Sub
