VERSION 5.00
Begin VB.Form Reservar 
   BackColor       =   &H006E943D&
   Caption         =   "Reserva de canchas"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
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
   Begin VB.CommandButton btnReservar 
      Caption         =   "Reservar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   18
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H006E943D&
      Caption         =   "Cancha y horario"
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
      Height          =   3495
      Left            =   3600
      TabIndex        =   15
      Top             =   3120
      Width           =   5295
      Begin VB.ListBox listHorarios 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   2175
      End
      Begin VB.ListBox listCanchas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H006E943D&
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   3480
      TabIndex        =   10
      Top             =   480
      Width           =   5535
      Begin VB.CommandButton btnRemoverB 
         Caption         =   "Remover"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3300
         TabIndex        =   14
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton btnRemoverA 
         Caption         =   "Remover"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   550
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ListBox listElegidoB 
         Height          =   1410
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
      Begin VB.ListBox listElegidoA 
         Height          =   1410
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.ListBox listSociosA 
      Height          =   4110
      ItemData        =   "Reservar.frx":0000
      Left            =   480
      List            =   "Reservar.frx":0002
      TabIndex        =   7
      Top             =   480
      Width           =   2895
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
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H006E943D&
      Caption         =   "Modalidad"
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
      Height          =   1695
      Left            =   9480
      TabIndex        =   3
      Top             =   5520
      Width           =   2295
      Begin VB.OptionButton optSingles 
         BackColor       =   &H006E943D&
         Caption         =   "Singles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optDobles 
         BackColor       =   &H006E943D&
         Caption         =   "Dobles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.ListBox listSociosB 
      Height          =   4110
      Left            =   9120
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton btnElegirA 
      Caption         =   "Elegir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton btnElegirB 
      Caption         =   "Elegir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "Editar reserva"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   20
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar reserva"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   21
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   360
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jugador 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   960
      TabIndex        =   9
      Top             =   0
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jugador 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9600
      TabIndex        =   8
      Top             =   0
      Width           =   1620
   End
End
Attribute VB_Name = "Reservar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim countA As Integer
Dim countB As Integer
Dim cancha_ As String
Dim canchaNum As String
Dim a
Dim reservaPasada As Boolean
Dim canchaReservada As String
Dim turnoReservado As String
Dim fechaDeReserva As Date
Dim reservante As String

Private Sub btnCancelar_Click()
    Dim res As Integer
    res = MsgBox("Está seguro de que quiere cancelar su reserva?", vbYesNo, "Confirmar cancelación")
           
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
        rsReservas.Open "SELECT * FROM Reservas WHERE Jugador1A = '" & reservante & "' AND Fecha LIKE '%" & fechaDeReserva & _
                        "%'", cn
                        
        If rsReservas.EOF Then Exit Sub
        
        'Eliminamos la entrada encontrada y actualizamos la DB
        rsReservas.Delete
        rsReservas.Update
        
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE ultimaReserva LIKE '%" & fechaDeReserva & "%'", cn
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
        
        rsHorarios.Source = "Horarios"
        rsHorarios.CursorType = adOpenKeyset
        rsHorarios.LockType = adLockOptimistic
        rsHorarios.Open "SELECT * FROM Horarios WHERE Hora = '" & turnoReservado & "'", cn
        rsHorarios.MoveFirst
        
        With rsHorarios
            .Requery
            
            .Fields(canchaReservada) = 0
                
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With
        
        cn.Close
        
        MsgBox ("Su reserva ha sido cancelada. Puede volver a reservar en el día corriente.")
        
        Menu.Show
        Unload Me
    End If
End Sub

Private Sub btnEditar_Click()
    'Verificamos que se hayan seleccionado todos los datos necesarios
  '|--------------------------------------------------------------------------------------------------|
    If optDobles.Value = True Then
        If listElegidoA.ListCount = 0 Then
            MsgBox ("Por favor seleccione jugadores para el equipo 1")
            listSociosA.SetFocus
            Exit Sub
        ElseIf listElegidoA.ListCount = 1 Then
            MsgBox ("Por favor seleccione el segundo jugador del equipo 1")
            listSociosA.SetFocus
            Exit Sub
        End If
    
    If listElegidoB.ListCount = 0 Then
        MsgBox ("Por favor seleccione jugadores para el equipo 2")
            listSociosB.SetFocus
            Exit Sub
        ElseIf listElegidoB.ListCount = 1 Then
            MsgBox ("Por favor seleccione el segundo jugador del equipo 2")
            listSociosB.SetFocus
            Exit Sub
        End If
    End If

    If listElegidoA.ListCount = 0 Then
        MsgBox ("Por favor seleccione el jugador 1")
        listSociosA.SetFocus
        Exit Sub
    End If
    
    If listElegidoB.ListCount = 0 Then
        MsgBox ("Por favor seleccione el jugador 2")
        listSociosB.SetFocus
        Exit Sub
    End If

    If Not listCanchas.ListIndex >= 0 Then
        MsgBox ("Por favor seleccione una cancha para reservar")
        listCanchas.SetFocus
        Exit Sub
    End If
    
    If Not listHorarios.ListIndex >= 0 Then
        MsgBox ("Por favor seleccione un horario para reservar")
        listHorarios.SetFocus
        Exit Sub
    End If
   '|--------------------------------------------------------------------------------------------------|
   
   'Mensaje de confirmación
   '|--------------------------------------------------------------------------------------------------|
    Dim res As Integer
    
    If optSingles.Value = True Then
        res = MsgBox("Revise los datos ingresados y confirme su reserva:" & Chr(10) & Chr(10) & _
        "Modo de juego: Singles" & Chr(10) & listElegidoA.List(0) & " vs " & listElegidoB.List(0) & Chr(10) & _
        listCanchas.List(listCanchas.ListIndex) & " entre las " & listHorarios.List(listHorarios.ListIndex) & "hs", vbYesNo, _
        "Confirmar reserva")
    Else
        res = MsgBox("Revise los datos ingresados y confirme su reserva:" & Chr(10) & Chr(10) & _
        "Modo de juego: Dobles" & Chr(10) & listElegidoA.List(0) & "/" & listElegidoA.List(1) & _
        " vs " & listElegidoB.List(0) & "/" & listElegidoB.List(1) & Chr(10) & _
        listCanchas.List(listCanchas.ListIndex) & " entre las " & listHorarios.List(listHorarios.ListIndex) & "hs", vbYesNo, _
        "Confirmar reserva")
    End If
    '|--------------------------------------------------------------------------------------------------|
    
    If res = 6 Then 'Me imaginaria que res tomaría como valor 1 pero toma 6 como valor si la respuesta es si, 7 si es no.
        Set cn = New ADODB.Connection
        Set rsReservas = New ADODB.Recordset
        Dim sentencia As String
        
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
        
        rsReservas.Source = "Reservas"
        rsReservas.CursorType = adOpenKeyset
        rsReservas.LockType = adLockOptimistic
        rsReservas.Open "SELECT * FROM Reservas WHERE Jugador1A = '" & reservante & "' AND Fecha LIKE '%" & fechaDeReserva & _
        "%'", cn
        
        'Editamos la reserva
        With rsReservas
            .Requery
            
            .Fields("Turno") = listHorarios.List(listHorarios.ListIndex)
            .Fields("Cancha") = UCase(listCanchas.List(listCanchas.ListIndex))
            .Fields("Fecha") = Date
            .Fields("Jugador1A") = listElegidoA.List(0)
            .Fields("Jugador2A") = listElegidoB.List(0)
            
            If optDobles.Value = True Then
                .Fields("Jugador1B") = listElegidoB.List(1)
                .Fields("Jugador2B") = listElegidoB.List(1)
            End If
            
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With

        Set rsHorarios = New ADODB.Recordset
        
        rsHorarios.Source = "Horarios"
        rsHorarios.CursorType = adOpenKeyset
        rsHorarios.LockType = adLockOptimistic
        rsHorarios.Open "SELECT * FROM Horarios", cn
        rsHorarios.MoveFirst
        
        Do Until rsHorarios.EOF
            If rsHorarios.Fields("Hora") = turnoReservado And rsHorarios.Fields(canchaReservada) = 1 Then
                rsHorarios.Requery
                rsHorarios.Fields(canchaReservada) = 0
                rsHorarios.UpdateBatch
                rsHorarios.Requery
                Exit Do
            End If
            rsHorarios.MoveNext
        Loop
        
        With rsHorarios
            .Requery
            .Find "hora = '" & listHorarios.List(listHorarios.ListIndex) & "'"
            
            .Fields(cancha_) = 1
            
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With
        
        MsgBox ("Reserva editada con éxito.")
        
        If editando = True Then
            editando = False
            ConsultaReservas.Show
        Else
            Menu.Show
        End If
        
        Unload Me
    End If
End Sub

Private Sub btnElegirA_Click()
    'Si no hay nadie seleccionado, le pide al usuario que seleccione un jugador.
    If Not listSociosA.ListIndex >= 0 Then
        MsgBox ("Debe seleccionar un/a jugador/a.")
        Exit Sub
    End If

    listElegidoA.AddItem listSociosA.List(listSociosA.ListIndex)
    countA = listElegidoA.ListCount
    
    If optSingles.Value = True Then
        btnElegirA.Enabled = False
    Else
        If optSingles.Value = False And countA = 2 Then
            btnElegirA.Enabled = False
        End If
    End If
    
    btnRemoverA.Enabled = True
    
    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open "SELECT * FROM Socios WHERE Estado = 'OK' AND isHabilitado = 1", cn
    rsSocios.MoveFirst
    
    listSociosB.Clear
    listSociosA.Clear
    
    'Filtramos los nombres que ya esten elegidos en la lista contraria y que hayamos elegido
    While rsSocios.EOF = False
        If rsSocios.Fields("Estado") = "En deuda" Then rsSocios.MoveNext
        If Not rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = listElegidoA.List(0) Then
            If Not rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = listElegidoA.List(1) Then
                If Not rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = listElegidoB.List(0) Then
                    If Not rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = listElegidoB.List(1) Then
                        listSociosA.AddItem (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido"))
                        listSociosB.AddItem (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido"))
                    End If
                End If
            End If
        End If
        rsSocios.MoveNext
    Wend
End Sub

Private Sub btnElegirB_Click()
    If Not listSociosB.ListIndex >= 0 Then
        MsgBox ("Debe seleccionar un/a jugador/a.")
        Exit Sub
    End If
    
    listElegidoB.AddItem listSociosB.List(listSociosB.ListIndex)
    countB = listElegidoB.ListCount
    
    If optSingles.Value = True Then
        btnElegirB.Enabled = False
    Else
        If optSingles.Value = False And countB = 2 Then
            btnElegirB.Enabled = False
        End If
    End If
    
    btnRemoverB.Enabled = True
    
    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open "SELECT * FROM Socios WHERE Estado = 'OK' AND isHabilitado = 1", cn
    rsSocios.MoveFirst
    
    listSociosB.Clear
    listSociosA.Clear
    
    'Filtramos los nombres que ya esten elegidos en la lista contraria y que hayamos elegido
    While rsSocios.EOF = False
        If rsSocios.Fields("Estado") = "En deuda" Then rsSocios.MoveNext
        If Not rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = listElegidoB.List(0) Then
            If Not rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = listElegidoB.List(1) Then
                If Not rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = listElegidoA.List(0) Then
                    If Not rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") = listElegidoA.List(1) Then
                        listSociosA.AddItem (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido"))
                        listSociosB.AddItem (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido"))
                    End If
                End If
            End If
        End If
        rsSocios.MoveNext
    Wend
End Sub

Private Sub btnRemoverA_Click()
    If countA = 0 Then
        Exit Sub
    End If
    
    If (Not listElegidoA.ListIndex >= 0) And listElegidoA.List(countA - 1) <> nombreUser Then
        listElegidoA.RemoveItem (countA - 1)
    ElseIf listElegidoA.List(listElegidoA.ListIndex) <> nombreUser Then
        listElegidoA.RemoveItem (listElegidoA.ListIndex)
    End If
        
    countA = listElegidoA.ListCount
    btnElegirA.Enabled = True
    
    If countA = 0 Or listElegidoA.List(0) = nombreUser Then
        btnRemoverA.Enabled = False
    End If
    
    listSociosA.Clear
    listSociosB.Clear
    a = cargarSocios()
End Sub

Private Sub btnRemoverB_Click()
    If countB = 0 Then
        Exit Sub
    End If

    If Not listElegidoB.ListIndex >= 0 Then
        listElegidoB.RemoveItem (countB - 1)
    Else: listElegidoB.RemoveItem (listElegidoB.ListIndex)
    End If
    
    countB = listElegidoB.ListCount
    btnElegirB.Enabled = True
    
    If countB = 0 Then btnRemoverB.Enabled = False
    
    listSociosA.Clear
    listSociosB.Clear
    a = cargarSocios()
End Sub

Private Sub btnReservar_Click()
    'Verificamos que se hayan seleccionado todos los datos necesarios
  '|--------------------------------------------------------------------------------------------------|
    If optDobles.Value = True Then
        If listElegidoA.ListCount = 0 Then
            MsgBox ("Por favor seleccione jugadores para el equipo 1")
            listSociosA.SetFocus
            Exit Sub
        ElseIf listElegidoA.ListCount = 1 Then
            MsgBox ("Por favor seleccione el segundo jugador del equipo 1")
            listSociosA.SetFocus
            Exit Sub
        End If
    
    If listElegidoB.ListCount = 0 Then
        MsgBox ("Por favor seleccione jugadores para el equipo 2")
            listSociosB.SetFocus
            Exit Sub
        ElseIf listElegidoB.ListCount = 1 Then
            MsgBox ("Por favor seleccione el segundo jugador del equipo 2")
            listSociosB.SetFocus
            Exit Sub
        End If
    End If

    If listElegidoA.ListCount = 0 Then
        MsgBox ("Por favor seleccione el jugador 1")
        listSociosA.SetFocus
        Exit Sub
    End If
    
    If listElegidoB.ListCount = 0 Then
        MsgBox ("Por favor seleccione el jugador 2")
        listSociosB.SetFocus
        Exit Sub
    End If

    If Not listCanchas.ListIndex >= 0 Then
        MsgBox ("Por favor seleccione una cancha para reservar")
        listCanchas.SetFocus
        Exit Sub
    End If
    
    If Not listHorarios.ListIndex >= 0 Then
        MsgBox ("Por favor seleccione un horario para reservar")
        listHorarios.SetFocus
        Exit Sub
    End If
   '|--------------------------------------------------------------------------------------------------|
   
   'Mensaje de confirmación
   '|--------------------------------------------------------------------------------------------------|
    Dim res As Integer
    
    If optSingles.Value = True Then
        res = MsgBox("Revise los datos ingresados y confirme su reserva:" & Chr(10) & Chr(10) & _
        "Modo de juego: Singles" & Chr(10) & listElegidoA.List(0) & " vs " & listElegidoB.List(0) & Chr(10) & _
        listCanchas.List(listCanchas.ListIndex) & " entre las " & listHorarios.List(listHorarios.ListIndex) & "hs", vbYesNo, _
        "Confirmar reserva")
    Else
        res = MsgBox("Revise los datos ingresados y confirme su reserva:" & Chr(10) & Chr(10) & _
        "Modo de juego: Dobles" & Chr(10) & listElegidoA.List(0) & "/" & listElegidoA.List(1) & _
        " vs " & listElegidoB.List(0) & "/" & listElegidoB.List(1) & Chr(10) & _
        listCanchas.List(listCanchas.ListIndex) & " entre las " & listHorarios.List(listHorarios.ListIndex) & "hs", vbYesNo, _
        "Confirmar reserva")
    End If
    '|--------------------------------------------------------------------------------------------------|
    
    If res = 6 Then 'Me imaginaria que res tomaría como valor 1 pero toma 6 como valor si la respuesta es si, 7 si es no.
        Set cn = New ADODB.Connection
        Set rsReservas = New ADODB.Recordset
        Dim sentencia As String
        
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
        
        'Componemos la sentencia INSERT
        If optSingles.Value = True Then
            sentencia = "INSERT INTO Reservas" & _
                "(Turno, Cancha, Jugador1A, Jugador2A, Fecha) " & _
                " VALUES (" & _
                "'" & listHorarios.List(listHorarios.ListIndex) & "', " & _
                "'" & UCase(listCanchas.List(listCanchas.ListIndex)) & "', " & _
                "'" & listElegidoA.List(0) & "', " & "'" & listElegidoB.List(0) & "', " & _
                "'" & Date & "'" & ")"
        Else
            sentencia = "INSERT INTO Reservas" & _
                "(Turno, Cancha, Jugador1A, Jugador1B, Jugador2A, Jugador2B, Fecha) " & _
                " VALUES (" & _
                "'" & listHorarios.List(listHorarios.ListIndex) & "', " & _
                "'" & UCase(listCanchas.List(listCanchas.ListIndex)) & "', " & _
                "'" & listElegidoA.List(0) & "', " & "'" & listElegidoA.List(1) & "', " & _
                "'" & listElegidoB.List(0) & "', " & "'" & listElegidoB.List(1) & "', " & _
                "'" & Date & "'" & ")"
        End If
            
        'Ejecutamos la sentencia
        cn.Execute sentencia, , adCmdText
        
        Set rsHorarios = New ADODB.Recordset
        
        rsHorarios.Source = "Horarios"
        rsHorarios.CursorType = adOpenKeyset
        rsHorarios.LockType = adLockOptimistic
        rsHorarios.Open "SELECT * FROM Horarios ", cn
        rsHorarios.MoveFirst
        
        With rsHorarios
            .Requery
            .Find "hora = '" & listHorarios.List(listHorarios.ListIndex) & "'"
            
            .Fields(cancha_) = 1
            
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With
        
        Set rsSocios = New ADODB.Recordset
        
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        
        Dim user As String
        If loginRol = 3 Then
            rsSocios.Open "SELECT * FROM Socios WHERE Usuario = '" & loginUser & "'", cn
            With rsSocios
                .Requery
                    
                .Fields("reservaAux") = .Fields("ultimaReserva")
                .Fields("ultimaReserva") = Date
                    
                .UpdateBatch 'Actualizamos la DB
                .Requery
            End With
        Else
            rsSocios.Open "SELECT * FROM Socios", cn
            rsSocios.MoveFirst
            Do Until rsSocios.EOF
                If (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido")) = listElegidoA.List(0) Then
                    user = rsSocios.Fields("Usuario")
                    With rsSocios
                        .Requery
                        .Find "Usuario = '" & user & "'"
                
                        .Fields("reservaAux") = .Fields("ultimaReserva")
                        .Fields("ultimaReserva") = Date
                            
                        .UpdateBatch 'Actualizamos la DB
                        .Requery
                    End With
                    Exit Do
                End If
                rsSocios.MoveNext
            Loop
        End If
            
        'Cerramos la conexión
        cn.Close
        
        Menu.Show
        Unload Me
    End If
End Sub

Private Sub btnVolver_Click()
    If editando = True Then
        editando = False
        ConsultaReservas.Show
    Else
        Menu.Show
    End If
        
    Unload Me
End Sub

Private Sub Form_Load()
    countA = 0
    countB = 0
    reservaPasada = False
    
    Set cn = New ADODB.Connection
    Set rsCanchas = New ADODB.Recordset
    Set rsSocios = New ADODB.Recordset

    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    
    'Nos fijamos si existe una reserva en el día de hoy de parte del usuario que ingresó o del cual se está editando
    'la reserva.
    
    If editando = False Then
        rsSocios.Open "SELECT * FROM Socios WHERE Usuario = '" & loginUser & "' AND ultimaReserva LIKE '%" & Date & "%'", cn
    Else
        If loginRol = 1 Then btnCancelar.Enabled = False
        rsSocios.Open "SELECT * FROM Socios WHERE Usuario = '" & editUser & "' AND ultimaReserva LIKE '%" & Date & "%'", cn
        btnVolver.Caption = "Cancelar"
    End If
    
    'Si no encuentra nada ejecuta normalmente.
    If rsSocios.EOF Then
        GoTo cargarNormal
    'De lo contrario carga los datos de la reserva siendo editada
    Else
        Set rsReservas = New ADODB.Recordset
    
        rsReservas.Source = "Reservas"
        rsReservas.CursorType = adOpenKeyset
        rsReservas.LockType = adLockOptimistic
        
        'Buscamos la reserva en el día de la fecha del socio seleccionado.
        rsReservas.Open "SELECT * FROM Reservas WHERE Jugador1A = '" & _
                         (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido") & "' AND Fecha LIKE '%" & _
                          Date & "%'"), cn
                          
        'Si la reserva ya ocurrió, ejecuta el código normalmente y no permite reservar, editar o eliminar nada.
        If DateDiff("h", Time, Mid(rsReservas.Fields("Turno"), 1, 5)) <= 0 Then
            reservaPasada = True
            cn.Close
            GoTo cargarNormal
        End If
        
        canchaReservada = UCase(Mid(rsReservas.Fields("Cancha"), 1, 6)) & _
                          Mid(Trim(rsReservas.Fields("Cancha")), 11, Len(rsReservas.Fields("Cancha")))
                          
        turnoReservado = rsReservas.Fields("Turno")
        fechaDeReserva = rsReservas.Fields("Fecha")
        reservante = rsReservas.Fields("Jugador1A")
        
        'Cargamos los dos jugadores principales.
        listElegidoA.AddItem (rsReservas.Fields("Jugador1A"))
        listElegidoB.AddItem (rsReservas.Fields("Jugador2A"))
        
        'Si la reserva es dual, cargamos el resto de los jugadores.
        If rsReservas.Fields("Jugador1B") <> "" Then
            listElegidoA.List(1) = (rsReservas.Fields("Jugador1B"))
            listElegidoB.List(1) = (rsReservas.Fields("Jugador2B"))
            optDobles.Value = True
        End If
        
        a = cargarSocios()
        
        Set rsCanchas = New ADODB.Recordset
        
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
        
        rsCanchas.Source = "Canchas"
        rsCanchas.CursorType = adOpenKeyset
        rsCanchas.LockType = adLockOptimistic
        rsCanchas.Open "SELECT * FROM Canchas ", cn
        rsCanchas.MoveFirst
    
        Dim bandera As Boolean
        bandera = True
        'Carga las canchas en el ListBox.
        While rsCanchas.EOF = False
            If rsCanchas.Fields("IsHabilitada") = 1 Then
                If bandera = True Then
                    canchaNum = Mid(rsCanchas.Fields("Cancha"), Len(rsCanchas.Fields("Cancha")), 3)
                    bandera = False
                End If
            
                listCanchas.AddItem "Cancha " & rsCanchas.Fields("Cancha")
            End If
            rsCanchas.MoveNext
        Wend
        
        'Busca la cancha de la reserva siendo editada.
        Dim i As Integer
        For i = 0 To listCanchas.ListCount
            If Trim(UCase(listCanchas.List(i))) = Trim(rsReservas.Fields("Cancha")) Then
                Exit For
            End If
        Next i
        
        listCanchas.Selected(i) = True 'Selecciona la cancha editada y provoca que se carguen los horarios para la misma.
        
        'Revisa los horarios hasta encontrar el de la reserva actual.
        For i = 0 To listHorarios.ListCount
            If Trim(listHorarios.List(i)) = turnoReservado Then
                Exit For
            End If
        Next i
        
        listHorarios.Selected(i) = True 'Selecciona el horario siendo editado.
        
        btnReservar.Visible = False
        btnEditar.Visible = True
        btnCancelar.Visible = True
    End If
    
    btnElegirA.Enabled = False
    
    If countA = 2 Then
        btnRemoverA.Enabled = True
    Else
        btnRemoverA.Enabled = False
    End If
    
    If countB > 0 Then
        btnRemoverB.Enabled = True
        btnElegirB.Enabled = False
    End If
    
    Exit Sub

cargarNormal:
    a = cargarSocios()
        
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsCanchas.Source = "Canchas"
    rsCanchas.CursorType = adOpenKeyset
    rsCanchas.LockType = adLockOptimistic
    rsCanchas.Open "SELECT * FROM Canchas ", cn
    rsCanchas.MoveFirst
    
    While rsCanchas.EOF = False
        If rsCanchas.Fields("IsHabilitada") = 1 Then
            listCanchas.AddItem "Cancha " & rsCanchas.Fields("Cancha")
        End If
        rsCanchas.MoveNext
    Wend
    
    btnRemoverA.Enabled = False
    btnRemoverB.Enabled = False
    
    If loginRol = 1 Or loginRol = 0 Then
        Label3.Caption = "El primer jugador 1 será el" & Chr(10) & "socio el dueño de la reserva."
        Label3.Visible = True
    End If
    
    cn.Close
End Sub

Private Sub listCanchas_Click()
    Set cn = New ADODB.Connection
    Set rsHorarios = New ADODB.Recordset
    Set rsCanchas = New ADODB.Recordset

    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"

    rsCanchas.Source = "Canchas"
    rsCanchas.CursorType = adOpenKeyset
    rsCanchas.LockType = adLockOptimistic
    rsCanchas.Open "SELECT * FROM Canchas ", cn
    rsCanchas.MoveFirst
    
    rsHorarios.Source = "Horarios"
    rsHorarios.CursorType = adOpenKeyset
    rsHorarios.LockType = adLockOptimistic
    rsHorarios.Open "SELECT * FROM Horarios", cn
    rsHorarios.MoveFirst
    
    Dim i As Integer
    
    cancha_ = "Cancha" & listCanchas.ListIndex + canchaNum
    
    listHorarios.Clear
    
    While rsHorarios.EOF = False
        If rsHorarios.Fields("IsHabilitado") = 1 And rsHorarios.Fields(cancha_) = 0 And _
        DateDiff("h", Time, Mid(rsHorarios.Fields("Hora"), 1, 5)) > 0 Then
            listHorarios.AddItem rsHorarios.Fields("Hora")
        Else
            If rsHorarios.Fields("IsHabilitado") = 1 And UCase(cancha_) = canchaReservada And _
            rsHorarios.Fields("Hora") = turnoReservado Then
                listHorarios.AddItem rsHorarios.Fields("Hora")
            End If
        End If
        rsHorarios.MoveNext
    Wend
End Sub

Private Sub listElegidoA_Click()
    If listElegidoA.List(listElegidoA.ListIndex) = nombreUser Then
        btnRemoverA.Enabled = False
    Else: btnRemoverA.Enabled = True
    End If
End Sub

Private Sub listSociosA_Click()
    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset

    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open "SELECT * FROM Socios", cn
    rsSocios.MoveFirst
    
    If loginRol <> 3 Then
        If countA = 0 Then
            rsSocios.MoveFirst
            Do Until rsSocios.EOF
                If (rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido")) = listSociosA.List(listSociosA.ListIndex) Then
                    If rsSocios.Fields("ultimaReserva") = Date Then
                        btnElegirA.Enabled = False
                    Else: btnElegirA.Enabled = True
                    End If
                    Exit Do
                End If
            rsSocios.MoveNext
            Loop
        End If
    End If
End Sub

Private Sub optDobles_Click()
    Label1.Caption = "Equipo 1"
    Label2.Caption = "Equipo 2"
    
    If countA < 2 And countB < 2 Then
        btnElegirA.Enabled = True
        btnElegirB.Enabled = True
    End If
End Sub

Private Sub optSingles_Click()
    Label1.Caption = "Jugador 1"
    Label2.Caption = "Jugador 2"
    
    If listElegidoA.List(0) <> nombreUser Then
        listElegidoA.Clear
        listElegidoB.Clear
        listSociosA.Clear
        listSociosB.Clear
        a = cargarSocios()
    Else
        If listElegidoA.ListCount = 2 Then listElegidoA.RemoveItem (1)
        listElegidoB.Clear
        listSociosA.Clear
        listSociosB.Clear
        a = cargarSocios()
        btnRemoverA.Enabled = False
    End If
    
    countA = listElegidoA.ListCount
    countB = listElegidoB.ListCount
    
    If countA = 1 Then btnElegirA.Enabled = False
    If countB = 1 Then btnElegirB.Enabled = False
End Sub

Function cargarSocios()
    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset
    Dim socio As String

    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    rsSocios.Open "SELECT * FROM Socios WHERE Estado = 'OK' AND isHabilitado = 1", cn
    rsSocios.MoveFirst
    
    While rsSocios.EOF = False
        socio = rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido")
        If rsSocios.Fields("Usuario") = loginUser Then
            If listElegidoA.List(0) <> nombreUser Then
                btnElegirA.Enabled = False
                btnRemoverA.Enabled = True
                
                'Desactiva el botón para reservar si ya se reservó en el día
                If rsSocios.Fields("ultimaReserva") = Date Then
                    'Si la reserva aun no ocurrió permite editarla o cancelarla
                    If loginRol = 3 And reservaPasada = False Then
                        btnReservar.Visible = False
                        btnEditar.Visible = True
                        btnCancelar.Visible = True
                    Else
                        btnReservar.Enabled = False
                        Label3.Caption = "Ha alcanzado el límite de" & Chr(10) & "reservas por hoy. Por favor," & Chr(10) & "vuelva a reservar mañana."
                        Label3.FontSize = 12
                        Label3.Visible = True
                        listElegidoA.AddItem socio
                    End If
                Else
                    listElegidoA.AddItem socio
                End If
            End If
        Else
            If Not rsSocios.Fields("Estado") = "En deuda" Then
                If socio <> listElegidoA.List(1) And socio <> listElegidoB.List(0) And socio <> listElegidoB.List(1) Then
                    listSociosA.AddItem rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido")
                    listSociosB.AddItem rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido")
                End If
            End If
        End If
        rsSocios.MoveNext
    Wend
    
    countA = listElegidoA.ListCount
    countB = listElegidoB.ListCount
    cn.Close
    
    cargarSocios = Null
End Function
