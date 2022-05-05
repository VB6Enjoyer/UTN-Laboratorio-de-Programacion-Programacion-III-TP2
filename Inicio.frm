VERSION 5.00
Begin VB.Form Inicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de reserva de canchas"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12450
   ClipControls    =   0   'False
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
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Inicio.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCerrar 
      BackColor       =   &H006E943D&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4790
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton btnEntrar 
      BackColor       =   &H006E943D&
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4790
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15 de Junio de 2021"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   6840
      TabIndex        =   3
      Top             =   6060
      Width           =   4635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Juan Ignacio Núñez"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1080
      TabIndex        =   2
      Top             =   6060
      Width           =   4500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión de reserva de canchas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Left            =   2445
      TabIndex        =   1
      Top             =   960
      Width           =   7575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mitre Tenis Club"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1125
      Left            =   2470
      TabIndex        =   0
      Top             =   -120
      Width           =   7500
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCerrar_Click()
    Dim respuesta
    respuesta = MsgBox("¿Está seguro de que quiere salir?", vbYesNo, "Salir")
    If respuesta = vbYes Then
        End
    End If
End Sub

Private Sub btnEntrar_Click()
    Login.Show 1, Me
End Sub

Private Sub Form_Load()
    editando = False
    editando2 = False

    'Todo esto lo iba a hacer una función para usarlo en el timer del Menú, pero las funciones o me crashean
    'o no me sirven, no se por qué.

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
    End If
    
    cn.Close
End Sub
