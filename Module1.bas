Attribute VB_Name = "Module1"
Option Explicit

Global loginUser As String
Global loginRol As Integer
Global nombreUser As String
Global editUser As String
Global dniUser As Long
Global editando As Boolean
Global editando2 As Boolean

Global rsCanchas As ADODB.Recordset
Global rsHorarios As ADODB.Recordset
Global rsReservas As ADODB.Recordset
Global rsUsuarios As ADODB.Recordset
Global rsSocios As ADODB.Recordset
Global cn As ADODB.Connection

'Epsilon Algorithm, Created by Simon Johnson
'Uses storing encrypted passwd's, or producing message digests.
'

Public Function Hash(ByVal text As String) As String
    Dim a
    Dim i
    
    a = 1
    For i = 1 To Len(text)
        a = Sqr(a * i * Asc(Mid(text, i, 1))) 'Numeric Hash
    Next i
    Rnd (-1)
    Randomize a 'seed PRNG
    
    For i = 1 To 16
        Hash = Hash & Chr(Int(Rnd * 256))
    Next i
End Function


