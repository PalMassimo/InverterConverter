VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_VVMD32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' VARIANT USED AS STRING ARRAYS
' VOGLIAMO IMPLEMENTARE UNA MAPPA LUNGHEZZA ALLELICA -> LUNGHEZZA ALLELICA IN BASE N
' PER FARLO INIZIALIZZIAMO DUE ARRAY DOVE ARRAY_1[i] <-> ARRAY_2[i]

Dim lunghezze_alleliche(24) As Variant
Dim lunghezze_alleliche_base_N(24) As Variant

Public Property Get lunghezza_array() As Integer
    lunghezza_array = 24
End Property


Public Function Converti(lunghezza_allelica_base_N As String) As String

'REMOVE SPACES
lunghezza_allelica_base_N = Replace(lunghezza_allelica_base_N, " ", "")


Dim index As Integer
index = -1 'SE IL VALORE NON è STATO TROVATO

For i = 0 To lunghezza_array() - 1 Step 1
    If lunghezza_allelica_base_N = lunghezze_alleliche_base_N(i) Then
        index = i
    End If
Next i

If index = -1 Then
    Converti = "-"
Else: Converti = lunghezze_alleliche(index)
End If
   

End Function

'DEFAULT METHOD TO INITIALIZE CLASS VARIABLES
Private Sub Class_Initialize()

lunghezze_alleliche(0) = 236
lunghezze_alleliche(1) = 238
lunghezze_alleliche(2) = 239
lunghezze_alleliche(3) = 240
lunghezze_alleliche(4) = 241
lunghezze_alleliche(5) = 245
lunghezze_alleliche(6) = 246
lunghezze_alleliche(7) = 247
lunghezze_alleliche(8) = 249
lunghezze_alleliche(9) = 251
lunghezze_alleliche(10) = 253
lunghezze_alleliche(11) = 255
lunghezze_alleliche(12) = 256
lunghezze_alleliche(13) = 257
lunghezze_alleliche(14) = 259
lunghezze_alleliche(15) = 261
lunghezze_alleliche(16) = 263
lunghezze_alleliche(17) = 265
lunghezze_alleliche(18) = 267
lunghezze_alleliche(19) = 269
lunghezze_alleliche(20) = 271
lunghezze_alleliche(21) = 273
lunghezze_alleliche(22) = 275
lunghezze_alleliche(23) = 293

lunghezze_alleliche_base_N(0) = "N"
lunghezze_alleliche_base_N(1) = "N+2"
lunghezze_alleliche_base_N(2) = "N+3"
lunghezze_alleliche_base_N(3) = "N+4"
lunghezze_alleliche_base_N(4) = "N+5"
lunghezze_alleliche_base_N(5) = "N+9"
lunghezze_alleliche_base_N(6) = "N+10"
lunghezze_alleliche_base_N(7) = "N+11"
lunghezze_alleliche_base_N(8) = "N+13"
lunghezze_alleliche_base_N(9) = "N+15"
lunghezze_alleliche_base_N(10) = "N+17"
lunghezze_alleliche_base_N(11) = "N+19"
lunghezze_alleliche_base_N(12) = "N+20"
lunghezze_alleliche_base_N(13) = "N+21"
lunghezze_alleliche_base_N(14) = "N+23"
lunghezze_alleliche_base_N(15) = "N+25"
lunghezze_alleliche_base_N(16) = "N+27"
lunghezze_alleliche_base_N(17) = "N+29"
lunghezze_alleliche_base_N(18) = "N+31"
lunghezze_alleliche_base_N(19) = "N+33"
lunghezze_alleliche_base_N(20) = "N+35"
lunghezze_alleliche_base_N(21) = "N+37"
lunghezze_alleliche_base_N(22) = "N+39"
lunghezze_alleliche_base_N(23) = "N+57"


End Sub








