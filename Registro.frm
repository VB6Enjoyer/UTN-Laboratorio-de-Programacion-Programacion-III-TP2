VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Registro 
   BackColor       =   &H006E943D&
   Caption         =   "Registro de socio"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
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
   ScaleHeight     =   5505
   ScaleWidth      =   11610
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
      Height          =   495
      Left            =   6960
      TabIndex        =   43
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton btnVolverA 
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton btnContinuar 
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   32
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox txtCodpostal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   30
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtLocalidad 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   28
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtProvincia 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   26
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtTelefono 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   24
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox txtApartamento 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   22
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtPiso 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      MaxLength       =   2
      TabIndex        =   20
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtNumero 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   18
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtCalle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   16
      Top             =   3360
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   4080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16777217
      CurrentDate     =   44352
   End
   Begin VB.ComboBox cmbDNI 
      Height          =   390
      ItemData        =   "Registro.frx":0000
      Left            =   2160
      List            =   "Registro.frx":0010
      TabIndex        =   12
      Top             =   3420
      Width           =   855
   End
   Begin VB.TextBox txtDNI 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtNacionalidad 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   2640
      Width           =   2655
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
      Height          =   1095
      Left            =   5040
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
      Begin VB.OptionButton optGen 
         BackColor       =   &H006E943D&
         Caption         =   "Mujer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optGen 
         BackColor       =   &H006E943D&
         Caption         =   "Varón"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox txtApellido 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaxLength       =   16
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbRol 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      ItemData        =   "Registro.frx":002E
      Left            =   4200
      List            =   "Registro.frx":003B
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4200
      MaxLength       =   16
      PasswordChar    =   "·"
      TabIndex        =   35
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton btnVolverB 
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   40
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btnRegistrar 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   41
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labObligatorio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Los campos subrayados son obligatorios"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3600
      TabIndex        =   42
      Top             =   720
      Width           =   4290
   End
   Begin VB.Label labRol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rol:"
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
      Left            =   4200
      TabIndex        =   39
      Top             =   3000
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label labPass 
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
      Left            =   4200
      TabIndex        =   36
      Top             =   2040
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label labPostal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código postal:"
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
      Left            =   6720
      TabIndex        =   31
      Top             =   2760
      Width           =   2010
   End
   Begin VB.Label labLocalidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7320
      TabIndex        =   29
      Top             =   2040
      Width           =   1440
   End
   Begin VB.Label labProvincia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7320
      TabIndex        =   27
      Top             =   1320
      Width           =   1410
   End
   Begin VB.Label labTel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
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
      Left            =   720
      TabIndex        =   25
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Label labApt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apartamento:"
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
      Left            =   6960
      TabIndex        =   23
      Top             =   4860
      Width           =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6645
      X2              =   6645
      Y1              =   5520
      Y2              =   1080
   End
   Begin VB.Label labPiso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Piso:"
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
      Left            =   10080
      TabIndex        =   21
      Top             =   4140
      Width           =   720
   End
   Begin VB.Label labNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7560
      TabIndex        =   19
      Top             =   4140
      Width           =   1170
   End
   Begin VB.Label labCalle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calle:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7920
      TabIndex        =   17
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label labBirth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nacimiento:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   360
      TabIndex        =   14
      Top             =   4080
      Width           =   1650
   End
   Begin VB.Label labAyuda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF1111&
      Height          =   285
      Left            =   4980
      MouseIcon       =   "Registro.frx":005A
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label labDNI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documento:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   360
      TabIndex        =   11
      Top             =   3435
      Width           =   1680
   End
   Begin VB.Label labNacionalidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nacionalidad:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2715
      Width           =   1890
   End
   Begin VB.Label labTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Formulario de registro de socio"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   7020
   End
   Begin VB.Label labApellido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label labNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   840
      TabIndex        =   1
      Top             =   1260
      Width           =   1170
   End
   Begin VB.Label labUsuario 
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
      Left            =   4200
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnContinuar_Click()
   'Verificaciones
   '|-------------------------------------------------------------|
    If Len(Trim(txtNombre)) = 0 Then
       MsgBox ("Debe ingresar un nombre.")
       txtNombre.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtApellido)) = 0 Then
       MsgBox ("Debe ingresar un apellido.")
       txtApellido.SetFocus
       Exit Sub
    End If

    If Len(Trim(cmbDNI)) = 0 Then
       MsgBox ("Debe elegir un tipo de documento.")
       cmbDNI.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtDNI)) = 0 Then
       MsgBox ("Debe ingresar un número de documento.")
       txtDNI.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtProvincia)) = 0 Then
       MsgBox ("Debe ingresar una provincia.")
       txtProvincia.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtLocalidad)) = 0 Then
       MsgBox ("Debe ingresar una localidad.")
       txtLocalidad.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtCalle)) = 0 Then
       MsgBox ("Debe ingresar el nombre de una calle.")
       txtCalle.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtNumero)) = 0 Then
       MsgBox ("Debe ingresar el número de la calle.")
       txtNumero.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtNombre)) = 0 Then
       MsgBox ("Debe ingresar un nombre.")
       txtNombre.SetFocus
       Exit Sub
    End If
   '|-------------------------------------------------------------|
    
    If editando = False Then
        Dim ctrl As Control
        
        For Each ctrl In Me
            ctrl.Visible = False
        Next ctrl
        
        labTitulo.Visible = True
        labUsuario.Visible = True
        labPass.Visible = True
        labRol.Visible = True
        
        txtUsuario.Visible = True
        txtPass.Visible = True
        cmbRol.Visible = True
        
        btnVolverB.Visible = True
        btnRegistrar.Visible = True
        
        Check1.Visible = True
        
        labTitulo.Caption = "Formulario de registro de usuario"
        Registro.Caption = "Registro de usuario"
    Else
        Dim genero As String
        Dim user As String
        Dim Nombre As String
   
        If optGen(0) Then genero = "Varón" Else: genero = "Mujer"
       
        Set cn = New ADODB.Connection
        Set rsSocios = New ADODB.Recordset
        Set rsUsuarios = New ADODB.Recordset
        Set rsReservas = New ADODB.Recordset
        
        Dim sentencia As String
        
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
        
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE DNI = " & dniUser & "", cn
        
        user = rsSocios.Fields("Usuario")
        Nombre = rsSocios.Fields("Nombre") & " " & rsSocios.Fields("Apellido")
        
        rsUsuarios.Source = "Usuarios"
        rsUsuarios.CursorType = adOpenKeyset
        rsUsuarios.LockType = adLockOptimistic
        rsUsuarios.Open "SELECT * FROM Usuarios WHERE Usuario = '" & user & "'", cn
        
        rsSocios.Requery
        rsSocios.Fields("Nombre") = UCase(Trim(txtNombre))
        rsSocios.Fields("Apellido") = UCase(Trim(txtApellido))
        rsSocios.Fields("Género") = UCase(genero)
        If Not Len(Trim(txtNacionalidad)) = 0 Then rsSocios.Fields("Nacionalidad") = UCase(Trim(txtNacionalidad))
        rsSocios.Fields("TipoDNI") = UCase(Trim(cmbDNI))
        rsSocios.Fields("DNI") = Trim(txtDNI)
        rsSocios.Fields("Nacimiento") = DTPicker1.Value
        rsSocios.Fields("Provincia") = UCase(Trim(txtProvincia))
        rsSocios.Fields("Localidad") = UCase(Trim(txtLocalidad))
        If Not Len(Trim(txtCodpostal)) = 0 Then rsSocios.Fields("Codigo postal") = Trim(txtCodpostal)
        rsSocios.Fields("Calle") = UCase(Trim(txtCalle))
        rsSocios.Fields("NroCalle") = Trim(txtNumero)
        If Not Len(Trim(txtPiso)) = 0 Then rsSocios.Fields("Piso") = Trim(txtPiso)
        If Not Len(Trim(txtApartamento)) = 0 Then rsSocios.Fields("Apartamento") = Trim(txtApartamento)
        If Not Len(Trim(txtTelefono)) = 0 Then rsSocios.Fields("Teléfono") = Trim(txtTelefono)
        rsSocios.Fields("Estado") = "OK"
        rsSocios.UpdateBatch
        rsSocios.Requery
    
        rsUsuarios.Requery
        rsUsuarios.Fields("Nombre") = UCase(Trim(txtNombre))
        rsUsuarios.Fields("Apellido") = UCase(Trim(txtApellido))
        rsUsuarios.UpdateBatch
        rsSocios.Requery
        
        rsReservas.Source = "Reservas"
        rsReservas.CursorType = adOpenKeyset
        rsReservas.LockType = adLockOptimistic
        rsReservas.Open "SELECT * FROM Reservas", cn
        rsReservas.MoveFirst
        
        rsReservas.Requery
        
        rsReservas.Find "Jugador1A = '" & Nombre & "'"
        If rsReservas.EOF = False Then
            Do Until rsReservas.EOF
                If rsReservas.Fields("Jugador1A") = Nombre Then
                    rsReservas.Fields("Jugador1A") = txtNombre & " " & txtApellido
                End If
                rsReservas.MoveNext
            Loop
            rsReservas.MoveFirst
        End If
        
        rsReservas.Find "Jugador1B = '" & Nombre & "'"
        If rsReservas.EOF = False Then
            Do Until rsReservas.EOF
                If rsReservas.Fields("Jugador1B") = Nombre Then
                    rsReservas.Fields("Jugador1B") = txtNombre & " " & txtApellido
                End If
                rsReservas.MoveNext
            Loop
            rsReservas.MoveFirst
        End If
        
        rsReservas.Find "Jugador2A = '" & Nombre & "'"
        If rsReservas.EOF = False Then
            Do Until rsReservas.EOF
                If rsReservas.Fields("Jugador2A") = Nombre Then
                    rsReservas.Fields("Jugador2A") = txtNombre & " " & txtApellido
                End If
                rsReservas.MoveNext
            Loop
            rsReservas.MoveFirst
        End If
        
        rsReservas.Find "Jugador2B = '" & Nombre & "'"
        If rsReservas.EOF = False Then
            Do Until rsReservas.EOF
                If rsReservas.Fields("Jugador2B") = Nombre Then
                    rsReservas.Fields("Jugador2B") = txtNombre & " " & txtApellido
                End If
                rsReservas.MoveNext
            Loop
            rsReservas.MoveFirst
        End If
        
        rsReservas.UpdateBatch
        rsReservas.Requery
        
        cn.Close
        
Finalizar:
        MsgBox ("Usuario editado con exito!")
        Unload Me
    End If
End Sub

Private Sub btnRegistrar_Click()
    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset
    Set rsUsuarios = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    rsUsuarios.Source = "Usuarios"
    rsUsuarios.CursorType = adOpenKeyset
    rsUsuarios.LockType = adLockOptimistic
    rsUsuarios.Open "SELECT * FROM Usuarios WHERE Usuario = '" & txtUsuario & "'", cn

   'Verificaciones
   '|-------------------------------------------------------------|
    If rsUsuarios.EOF = False And editando2 = False Then
        MsgBox ("Ya existe un usuario con este nombre. Por favor, seleccione otro nombre de usuario.")
        txtUsuario.SetFocus
        Exit Sub
    End If
    
    rsUsuarios.Close
    cn.Close
   
    If Len(Trim(txtUsuario)) = 0 Then
       MsgBox ("Debe ingresar un nombre de usuario.")
       txtUsuario.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtPass)) = 0 Then
       MsgBox ("Debe ingresar una contraseña.")
       txtPass.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(cmbRol)) = 0 Then
       MsgBox ("Debe elegir un rol.")
       cmbRol.SetFocus
       Exit Sub
    End If
   '|-------------------------------------------------------------|
   
   Dim genero As String
   
   If optGen(0) Then genero = "Varón" Else: genero = "Mujer"
   
    Set cn = New ADODB.Connection
    Set rsSocios = New ADODB.Recordset
    Set rsUsuarios = New ADODB.Recordset
    
    Dim sentencia As String
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
    
    If editando2 = False Then
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios ", cn
        rsSocios.MoveLast
        rsSocios.MoveNext
        
        rsUsuarios.Source = "Usuarios"
        rsUsuarios.CursorType = adOpenKeyset
        rsUsuarios.LockType = adLockOptimistic
        rsUsuarios.Open "SELECT * FROM Usuarios ", cn
        rsUsuarios.MoveLast
        rsUsuarios.MoveNext
        
        If rsSocios.EOF = True Then
            rsSocios.AddNew
            rsSocios.Fields("Nombre") = UCase(Trim(txtNombre))
            rsSocios.Fields("Apellido") = UCase(Trim(txtApellido))
            rsSocios.Fields("Género") = UCase(genero)
            If Not Len(Trim(txtNacionalidad)) = 0 Then rsSocios.Fields("Nacionalidad") = UCase(Trim(txtNacionalidad))
            rsSocios.Fields("TipoDNI") = UCase(Trim(cmbDNI))
            rsSocios.Fields("DNI") = Trim(txtDNI)
            rsSocios.Fields("Nacimiento") = DTPicker1.Value
            rsSocios.Fields("Provincia") = UCase(Trim(txtProvincia))
            rsSocios.Fields("Localidad") = UCase(Trim(txtLocalidad))
            If Not Len(Trim(txtCodpostal)) = 0 Then rsSocios.Fields("Codigo postal") = Trim(txtCodpostal)
            rsSocios.Fields("Calle") = UCase(Trim(txtCalle))
            rsSocios.Fields("NroCalle") = Trim(txtNumero)
            If Not Len(Trim(txtPiso)) = 0 Then rsSocios.Fields("Piso") = Trim(txtPiso)
            If Not Len(Trim(txtApartamento)) = 0 Then rsSocios.Fields("Apartamento") = Trim(txtApartamento)
            If Not Len(Trim(txtTelefono)) = 0 Then rsSocios.Fields("Teléfono") = Trim(txtTelefono)
            rsSocios.Fields("Estado") = "OK"
            rsSocios.Fields("Usuario") = UCase(Trim(txtUsuario))
            rsSocios.Update
        End If
    
        If rsUsuarios.EOF = True Then
            rsUsuarios.AddNew
            rsUsuarios.Fields("Usuario") = UCase(Trim(txtUsuario))
            rsUsuarios.Fields("Password") = Hash(Trim(txtPass))
            If cmbRol = "Operador" Then rsUsuarios.Fields("Rol") = 1
            If cmbRol = "Invitado" Then rsUsuarios.Fields("Rol") = 2
            If cmbRol = "Socio" Then rsUsuarios.Fields("Rol") = 3
            rsUsuarios.Fields("Nombre") = UCase(Trim(txtNombre))
            rsUsuarios.Fields("Apellido") = UCase(Trim(txtApellido))
            rsUsuarios.Update
        End If
        
        cn.Close
        
        MsgBox ("Usuario registrado con exito!")
    Else
        rsUsuarios.Source = "Usuarios"
        rsUsuarios.CursorType = adOpenKeyset
        rsUsuarios.LockType = adLockOptimistic
        rsUsuarios.Open "SELECT * FROM Usuarios WHERE Usuario = '" & editUser & "'", cn
        
        With rsUsuarios
            .Requery
            
            .Fields("Usuario") = txtUsuario.text
            .Fields("Password") = Hash(Trim(txtPass))
            If cmbRol = "Operador" Then rsUsuarios.Fields("Rol") = 1
            If cmbRol = "Invitado" Then rsUsuarios.Fields("Rol") = 2
            If cmbRol = "Socio" Then rsUsuarios.Fields("Rol") = 3
            
            .UpdateBatch
            .Requery
        End With
        
        MsgBox ("Usuario editado con exito!")
        editando2 = False
    End If
    
    Unload Me
End Sub

Private Sub btnVolverA_Click()
    If editando = True Then
        Socios.Show
        editando = False
    End If
    Unload Me
End Sub

Private Sub btnVolverB_Click()
    If editando2 = False Then
        Dim ctrl As Control
        
        For Each ctrl In Me
            ctrl.Visible = True
        Next ctrl
        
        labTitulo.Visible = False
        labUsuario.Visible = False
        labPass.Visible = False
        labRol.Visible = False
    
        txtUsuario.Visible = False
        txtPass.Visible = False
        cmbRol.Visible = False
        
        btnVolverB.Visible = False
        btnRegistrar.Visible = False
        
        Check1.Visible = False
        
        labTitulo.Caption = "Formulario de registro de socio"
        Registro.Caption = "Registro de usuario"
    Else
        editando2 = False
        Unload Me
    End If
End Sub

Private Sub Check1_Click()
    If Check1 Then txtPass.PasswordChar = "" Else: txtPass.PasswordChar = "·"
End Sub

Private Sub cmbDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDNI.SetFocus
End Sub

Private Sub cmbRol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnRegistrar.SetFocus Else: KeyAscii = 0
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelefono.SetFocus
End Sub

Private Sub Form_Load()
    cmbRol.Clear
    If loginRol = 0 Then
        cmbRol.List(0) = "Operador"
        cmbRol.List(1) = "Invitado"
        cmbRol.List(2) = "Socio"
    Else
        cmbRol.List(0) = "Invitado"
        cmbRol.List(1) = "Socio"
    End If
    
    If editando = True Then
        labTitulo.Caption = "Formulario de edición de socio"
        Registro.Caption = "Edición de socio"
        btnContinuar.Caption = "Editar"
        
        Set cn = New ADODB.Connection
        Set rsSocios = New ADODB.Recordset
        
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
        
        rsSocios.Source = "Socios"
        rsSocios.CursorType = adOpenKeyset
        rsSocios.LockType = adLockOptimistic
        rsSocios.Open "SELECT * FROM Socios WHERE DNI = " & dniUser & "", cn
        
        If rsSocios.EOF Then
            MsgBox ("Error: No se encontró el usuario que se está editando.")
            Unload Me
        End If
        
        If rsSocios.Fields("Género") = "MUJER" Then
            optGen(1).Value = True
        Else
            optGen(0).Value = True
        End If
        
        If Not Len(rsSocios.Fields("Nombre")) = 0 Then txtNombre = rsSocios.Fields("Nombre")
        If Not Len(rsSocios.Fields("Apellido")) = 0 Then txtApellido = rsSocios.Fields("Apellido")
        If Not Len(rsSocios.Fields("Nacionalidad")) = 0 Then txtNacionalidad = rsSocios.Fields("Nacionalidad")
        If Not Len(rsSocios.Fields("TipoDNI")) = 0 Then cmbDNI.text = rsSocios.Fields("TipoDNI")
        If Not Len(rsSocios.Fields("DNI")) = 0 Then txtDNI = rsSocios.Fields("DNI")
        If Not Len(rsSocios.Fields("Nacimiento")) = 0 Then DTPicker1 = rsSocios.Fields("Nacimiento")
        If Not Len(rsSocios.Fields("Teléfono")) = 0 Then txtTelefono = rsSocios.Fields("Teléfono")
        If Not Len(rsSocios.Fields("Provincia")) = 0 Then txtProvincia = rsSocios.Fields("Provincia")
        If Not Len(rsSocios.Fields("Localidad")) = 0 Then txtLocalidad = rsSocios.Fields("Localidad")
        If Not Len(rsSocios.Fields("Codigo postal")) = 0 Then txtCodpostal = rsSocios.Fields("Codigo postal")
        If Not Len(rsSocios.Fields("Calle")) = 0 Then txtCalle = rsSocios.Fields("Calle")
        If Not Len(rsSocios.Fields("NroCalle")) = 0 Then txtNumero = rsSocios.Fields("NroCalle")
        If Not Len(rsSocios.Fields("Piso")) = 0 Then txtPiso = rsSocios.Fields("Piso")
        If Not Len(rsSocios.Fields("Apartamento")) = 0 Then txtApartamento = rsSocios.Fields("Apartamento")
    End If
    
    If editando2 = True Then
        Dim ctrl As Control
        
        For Each ctrl In Me
            ctrl.Visible = False
        Next ctrl
        
        labTitulo.Visible = True
        labUsuario.Visible = True
        labPass.Visible = True
        labRol.Visible = True
        
        txtUsuario.Visible = True
        txtPass.Visible = True
        cmbRol.Visible = True
        
        btnVolverB.Visible = True
        btnRegistrar.Visible = True
        
        Check1.Visible = True
        
        labTitulo.Caption = "Formulario de edición de usuario"
        Registro.Caption = "Edición de usuario"
        btnRegistrar.Caption = "Editar"
        
        If loginRol <> 0 Then
            txtPass.Enabled = False
            Check1.Enabled = False
        End If
        
        Set cn = New ADODB.Connection
        Set rsUsuarios = New ADODB.Recordset
        
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Tenis.mdb;Persist Security Info=False"
        
        rsUsuarios.Source = "Usuarios"
        rsUsuarios.CursorType = adOpenKeyset
        rsUsuarios.LockType = adLockOptimistic
        rsUsuarios.Open "SELECT * FROM Usuarios WHERE Usuario = '" & editUser & "'", cn
        
        If rsUsuarios.EOF Then
            MsgBox ("No se ha encontrado el usuario indicado.")
            Unload Me
        End If
        
        txtUsuario = rsUsuarios.Fields("Usuario")
        
        Select Case rsUsuarios.Fields("Rol")
            Case 0
                cmbRol.Enabled = False
                cmbRol.text = "Admin"
            Case 1
                cmbRol.text = "Operador"
            Case 2
                cmbRol.text = "Invitado"
            Case 3
                cmbRol.text = "Socio"
        End Select
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labAyuda.Font.Underline = False
End Sub

Private Sub labAyuda_Click()
    MsgBox ("D.N.I. es Documento Nacional de Identidad" & Chr(10) & "L.C. es Libreta Cívica" & Chr(10) & "L.E. es Libreta de Enrolamiento" & Chr(10) & "C.I. es Cédula de Identidad")
End Sub

Private Sub labAyuda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labAyuda.Font.Underline = True
End Sub

Private Sub txtApartamento_LostFocus()
    txtApartamento = UCase(txtApartamento)
End Sub

Private Sub txtApartamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnContinuar.SetFocus
End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNacionalidad.SetFocus
End Sub

Private Sub txtapellido_LostFocus()
    txtApellido = UCase(txtApellido)
End Sub

Private Sub txtCalle_LostFocus()
    txtCalle = UCase(txtCalle)
End Sub

Private Sub txtCalle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNumero.SetFocus
End Sub

Private Sub txtCodpostal_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtCodpostal.text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    
    If KeyAscii = 13 Then txtCalle.SetFocus
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtDNI.text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    
    If KeyAscii = 13 Then DTPicker1.SetFocus
End Sub

Private Sub txtLocalidad_LostFocus()
    txtLocalidad = UCase(txtLocalidad)
End Sub

Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCodpostal.SetFocus
End Sub

Private Sub txtNacionalidad_LostFocus()
    txtNacionalidad = UCase(txtNacionalidad)
End Sub

Private Sub txtNacionalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbDNI.SetFocus
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtApellido.SetFocus
End Sub

Private Sub txtNombre_LostFocus()
    txtNombre = UCase(txtNombre)
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtNumero.text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    
    If KeyAscii = 13 Then txtPiso.SetFocus
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbRol.SetFocus
End Sub

Private Sub txtPiso_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtPiso.text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    
    If KeyAscii = 13 Then txtApartamento.SetFocus
End Sub

Private Sub txtProvincia_LostFocus()
    txtProvincia = UCase(txtProvincia)
End Sub

Private Sub txtProvincia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLocalidad.SetFocus
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtTelefono.text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    
    If KeyAscii = 13 Then txtProvincia.SetFocus
End Sub

Private Sub txtUsuario_LostFocus()
    txtUsuario = UCase(txtUsuario)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPass.SetFocus
End Sub
