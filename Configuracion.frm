VERSION 5.00
Begin VB.Form Configuracion 
   BackColor       =   &H006E943D&
   Caption         =   "Configuración"
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
   Picture         =   "Configuracion.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   12450
   Begin VB.Frame Frame2 
      BackColor       =   &H006E943D&
      Caption         =   "Canchas"
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   6240
      TabIndex        =   27
      Top             =   720
      Width           =   2535
      Begin VB.CommandButton btnGuardarcanchas 
         Caption         =   "Guardar"
         Height          =   495
         Left            =   720
         TabIndex        =   38
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   9
         Left            =   120
         TabIndex        =   37
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   36
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   7
         Left            =   120
         TabIndex        =   35
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H006E943D&
         Caption         =   "Cancha Nro. 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H006E943D&
      Caption         =   "Horarios"
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   3855
      Begin VB.CommandButton btnGuardar 
         Caption         =   "Guardar"
         Height          =   495
         Left            =   2160
         TabIndex        =   26
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   23
         Left            =   1920
         TabIndex        =   25
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   22
         Left            =   1920
         TabIndex        =   24
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   21
         Left            =   1920
         TabIndex        =   23
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   20
         Left            =   1920
         TabIndex        =   22
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   19
         Left            =   1920
         TabIndex        =   21
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   18
         Left            =   1920
         TabIndex        =   20
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   17
         Left            =   1920
         TabIndex        =   19
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   16
         Left            =   1920
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   15
         Left            =   1920
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   14
         Left            =   1920
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   13
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H006E943D&
         Caption         =   "00:00 - 1:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
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
      Left            =   600
      TabIndex        =   0
      Top             =   6000
      Width           =   2775
   End
End
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnGuardar_Click()
    Set cn = New ADODB.Connection
    Set rsHorarios = New ADODB.Recordset
    Set rsReservas = New ADODB.Recordset
    Set rsSocios = New ADODB.Recordset
    Dim i As Integer
    Dim res As Integer
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    rsHorarios.Source = "Horarios"
    rsHorarios.CursorType = adOpenKeyset
    rsHorarios.LockType = adLockOptimistic
    rsHorarios.Open "SELECT * FROM Horarios", cn
    rsHorarios.MoveFirst
    
    rsReservas.Source = "Reservas"
    rsReservas.CursorType = adOpenKeyset
    rsReservas.LockType = adLockOptimistic
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    
    Dim hora As String
    For i = 0 To 23
        res = Check1(i).Value
        hora = Trim(Check1(i).Caption)
        With rsHorarios
            .Requery
            .Find "id = '" & Trim(i + 1) & "'"
        
            If Not .EOF Then
                If (res) Then
                   !isHabilitado = Trim(1)
                Else
                    If Not (res) Then
                        !isHabilitado = Trim(0)
                        
                        rsReservas.Open "SELECT * FROM Reservas WHERE Fecha LIKE '%" & Date & "%' AND Turno = '" & hora & "'", cn
                        
                        'Si se detectan reservas en el turno en el día de la fecha pregunta si se quieren eliminar.
                        If Not rsReservas.EOF Then
                            res = MsgBox("Puede ser que hayan reservas en el turno " & hora & "." & Chr(10) & Chr(10) & _
                                         "¿Desea cancelar las reservas activas?", vbYesNo, "¿Cancelar reservas?")
                                         
                            If res = 6 Then
                                rsReservas.MoveFirst
                                Do Until rsReservas.EOF
                                    'Elimina todas las reservas activas en el día de la fecha.
                                    If rsReservas.Fields("Turno") = hora Then
                                        Dim cancha As String
                                        cancha = Trim(Mid(rsReservas.Fields("Cancha"), 1, 6) & Mid(rsReservas.Fields("Cancha"), Len(Check2(1).Caption), 7))
                                        
                                        If DateDiff("h", Time, Mid(rsReservas.Fields("Turno"), 1, 5)) > 0 Then
                                            rsReservas.Delete
                                            rsReservas.Update
                                            
                                            'Elimina el registro de reserva para cada socio con reservas activas.
                                            rsSocios.Open "SELECT * FROM Socios WHERE ultimaReserva LIKE '%" & Date & "%'", cn
                                            rsSocios.MoveFirst
                                            
                                            Dim user As String
                                            Do Until rsSocios.EOF
                                                user = rsSocios.Fields("Usuario")
                                            
                                                rsSocios.Requery
                                                rsSocios.Find "Usuario = '" & user & "'"
                                                rsSocios.Fields("ultimaReserva") = rsSocios.Fields("reservaAux")
                                                rsSocios.UpdateBatch
                                                rsSocios.MoveNext
                                            Loop
                                            
                                            rsSocios.Close
                                            
                                            'Reestablece todas las canchas en ese horario a "no reservada".
                                            Do Until rsHorarios.EOF
                                                If rsHorarios.Fields("Hora") = hora Then
                                                    If rsHorarios.Fields(cancha) = 1 Then
                                                        rsHorarios.Requery
                                                        rsHorarios.Fields(cancha) = 0
                                                        rsHorarios.UpdateBatch
                                                    End If
                                                End If
                                                rsHorarios.MoveNext
                                            Loop
                                            
                                        End If
                                    End If
                                    rsReservas.MoveNext
                                Loop
                            End If
                        End If
                        
                        rsReservas.Close
                    End If
                End If
            End If
            
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With
    Next i
    
    MsgBox ("Se ha guardado la configuración de los horarios exitosamente.")
End Sub

Private Sub btnGuardarcanchas_Click()
    Set cn = New ADODB.Connection
    Set rsCanchas = New ADODB.Recordset
    Set rsHorarios = New ADODB.Recordset
    Set rsReservas = New ADODB.Recordset
    Set rsSocios = New ADODB.Recordset
    Dim i As Integer
    Dim res As Integer
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    rsCanchas.Source = "Canchas"
    rsCanchas.CursorType = adOpenKeyset
    rsCanchas.LockType = adLockOptimistic
    rsCanchas.Open "SELECT * FROM Canchas", cn
    rsCanchas.MoveFirst
    
    rsHorarios.Source = "Horarios"
    rsHorarios.CursorType = adOpenKeyset
    rsHorarios.LockType = adLockOptimistic
    
    rsReservas.Source = "Reservas"
    rsReservas.CursorType = adOpenKeyset
    rsReservas.LockType = adLockOptimistic
    
    rsSocios.Source = "Socios"
    rsSocios.CursorType = adOpenKeyset
    rsSocios.LockType = adLockOptimistic
    
    Dim cancha As String
    For i = 0 To 9
        res = Check2(i).Value
        cancha = Trim(Mid(Check2(i).Caption, 1, 6) & Mid(Check2(i).Caption, Len(Check2(1).Caption), 7))
        With rsCanchas
            .Requery
            .Find "id = '" & Trim(i + 1) & "'"
        
            If Not .EOF Then
                If (res) Then
                   !isHabilitada = Trim(1)
                Else
                    If Not (res) Then
                        !isHabilitada = Trim(0)
                        
                        rsReservas.Open "SELECT * FROM Reservas WHERE Fecha LIKE '%" & Date & "%' AND Cancha = '" & _
                                        Trim(Check2(i).Caption) & "'", cn
                        
                        'Si se detectan reservas en la cancha en el día de la fecha pregunta si se quieren eliminar.
                        If Not rsReservas.EOF Then
                            res = MsgBox("Puede ser que hayan reservas activas en la Cancha N°" & i + 1 & "." & Chr(10) & Chr(10) & _
                                         "¿Desea cancelar las reservas activas?", vbYesNo, "¿Cancelar reservas?")
                                         
                            If res = 6 Then
                                rsReservas.MoveFirst
                                Do Until rsReservas.EOF
                                    Dim hora As String
                                    hora = rsReservas.Fields("Turno")
                                    
                                    'Elimina todas las reservas activas en el día de la fecha.
                                    If rsReservas.Fields("Cancha") = UCase(Trim(Check2(i).Caption)) Then
                                        If DateDiff("h", Time, Mid(rsReservas.Fields("Turno"), 1, 5)) > 0 Then
                                            rsReservas.Delete
                                            rsReservas.Update
                                            
                                            'Elimina el registro de reserva para cada socio con reservas activas.
                                            rsSocios.Open "SELECT * FROM Socios WHERE ultimaReserva LIKE '%" & Date & "%'", cn
                                            rsSocios.MoveFirst
                                            
                                            Dim user As String
                                            Do Until rsSocios.EOF
                                                user = rsSocios.Fields("Usuario")
                                            
                                                rsSocios.Requery
                                                rsSocios.Find "Usuario = '" & user & "'"
                                                rsSocios.Fields("ultimaReserva") = rsSocios.Fields("reservaAux")
                                                rsSocios.UpdateBatch
                                                rsSocios.MoveNext
                                            Loop
                                            
                                            rsSocios.Close
                                            
                                            rsHorarios.Open "SELECT * FROM Horarios WHERE Hora = '" & hora & "'", cn
                                            rsHorarios.MoveFirst
                                            
                                            'Reestablece todos los horarios de la cancha a "no reservado".
                                            Do Until rsHorarios.EOF
                                                If rsHorarios.Fields(cancha) = 1 Then
                                                    rsHorarios.Requery
                                                    rsHorarios.Fields(cancha) = 0
                                                    rsHorarios.UpdateBatch
                                                End If
                                                rsHorarios.MoveNext
                                            Loop
                                            
                                            rsHorarios.Close
                                        End If
                                    End If
                                    rsReservas.MoveNext
                                Loop
                            End If
                        End If
                        
                        rsReservas.Close
                    End If
                End If
            End If
        
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With
    Next i
    
    MsgBox ("Se ha guardado la configuración de las canchas exitosamente.")
End Sub

Private Sub btnVolver_Click()
    Menu.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rsHorarios = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsHorarios.Source = "Horarios"
    rsHorarios.CursorType = adOpenKeyset
    rsHorarios.LockType = adLockOptimistic
    rsHorarios.Open "SELECT * FROM Horarios", cn, adOpenDynamic, adLockOptimistic
    rsHorarios.MoveFirst
    
    Dim i As Integer
    i = 0
    Do Until rsHorarios.EOF
        'Cambia el caption de cada uno de los CheckBoxes con el del ID correspondiente
        Check1(i).Caption = rsHorarios.Fields("Hora")
        If rsHorarios.Fields("ID") = i + 1 Then 'Revisa si el ID es el mismo
            If rsHorarios.Fields("isHabilitado") = 1 Then
                Check1(i).Value = 1 'Marca la checkbox si el horario está habilitado, la desmarca si no.
            Else: Check1(i).Value = 0
            End If
        End If
        i = i + 1
        rsHorarios.MoveNext
    Loop
    
    Set rsCanchas = New ADODB.Recordset
    
    rsCanchas.Source = "Canchas"
    rsCanchas.CursorType = adOpenKeyset
    rsCanchas.LockType = adLockOptimistic
    rsCanchas.Open "SELECT * FROM Canchas", cn, adOpenDynamic, adLockOptimistic
    rsCanchas.MoveFirst
    
    i = 0
    Do Until rsCanchas.EOF
        'Cambia el caption de cada uno de los CheckBoxes con el del ID correspondiente
        Check2(i).Caption = "Cancha " & rsCanchas.Fields("Cancha")
        If rsCanchas.Fields("ID") = i + 1 Then 'Revisa si el ID es el mismo
            If rsCanchas.Fields("isHabilitada") = 1 Then
                Check2(i).Value = 1 'Marca la checkbox si la cancha está habilitada, la desmarca si no.
            Else: Check2(i).Value = 0
            End If
        End If
        i = i + 1
        rsCanchas.MoveNext
    Loop
End Sub

Private Sub btnGuardarB_Click()
    Set cn = New ADODB.Connection
    Set rsHorarios = New ADODB.Recordset
    Set rsReservas = New ADODB.Recordset
    Set rsSocios = New ADODB.Recordset
    Dim i As Integer
    Dim res As Integer
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    rsHorarios.Source = "Horarios"
    rsHorarios.CursorType = adOpenKeyset
    rsHorarios.LockType = adLockOptimistic
    rsHorarios.Open "SELECT * FROM Horarios", cn, adOpenDynamic, adLockOptimistic
    rsHorarios.MoveFirst
    
    For i = 0 To 23
        res = Check1(i).Value
        With rsHorarios
            .Requery
            .Find "id = '" & Trim(i + 1) & "'"
        
            If Not .EOF Then
                If (res) Then
                   !isHabilitado = Trim(1)
                Else
                    If Not (res) Then
                        !isHabilitado = Trim(0)
                    End If
                End If
            End If
            
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With
    Next i
End Sub

Private Sub btnGuardarcanchasB_Click()
    Set cn = New ADODB.Connection
    Set rsCanchas = New ADODB.Recordset
    Set rsHorarios = New ADODB.Recordset
    Set rsReservas = New ADODB.Recordset
    Set rsSocios = New ADODB.Recordset
    Dim i As Integer
    Dim res As Integer
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    rsCanchas.Source = "Canchas"
    rsCanchas.CursorType = adOpenKeyset
    rsCanchas.LockType = adLockOptimistic
    rsCanchas.Open "SELECT * FROM Canchas", cn, adOpenDynamic, adLockOptimistic
    rsCanchas.MoveFirst
    
    For i = 0 To 9
        res = Check2(i).Value
        With rsCanchas
            .Requery
            .Find "id = '" & Trim(i + 1) & "'"
        
            If Not .EOF Then
                If (res) Then
                   !isHabilitada = Trim(1)
                Else
                    If Not (res) Then
                        !isHabilitada = Trim(0)
                    End If
                End If
            End If
        
            .UpdateBatch 'Actualizamos la DB
            .Requery
        End With
    Next i
End Sub
