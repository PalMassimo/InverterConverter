VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_VVMD28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' VARIANT USED AS STRING ARRAYS
' VOGLIAMO IMPLEMENTARE UNA MAPPA LUNGHEZZA ALLELICA -> LUNGHEZZA ALLELICA IN BASE N
' PER FARLO INIZIALIZZIAMO DUE ARRAY DOVE ARRAY_1[i] <-> ARRAY_2[i]

Dim lunghezze_alleliche(32) As Variant
Dim lunghezze_alleliche_base_N(32) As Variant

Public Property Get lunghezza_array() As Integer
    lunghezza_array = 32
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

lunghezze_alleliche(0) = 219
lunghezze_alleliche(1) = 221
lunghezze_alleliche(2) = 223
lunghezze_alleliche(3) = 227
lunghezze_alleliche(4) = 229
lunghezze_alleliche(5) = 231
lunghezze_alleliche(6) = 233
lunghezze_alleliche(7) = 235
lunghezze_alleliche(8) = 236
lunghezze_alleliche(9) = 237
lunghezze_alleliche(10) = 238
lunghezze_alleliche(11) = 239
lunghezze_alleliche(12) = 241
lunghezze_alleliche(13) = 243
lunghezze_alleliche(14) = 245
lunghezze_alleliche(15) = 247
lunghezze_alleliche(16) = 249
lunghezze_alleliche(17) = 251
lunghezze_alleliche(18) = 253
lunghezze_alleliche(19) = 255
lunghezze_alleliche(20) = 257
lunghezze_alleliche(21) = 259
lunghezze_alleliche(22) = 261
lunghezze_alleliche(23) = 263
lunghezze_alleliche(24) = 265
lunghezze_alleliche(25) = 267
lunghezze_alleliche(26) = 273
lunghezze_alleliche(27) = 275
lunghezze_alleliche(28) = 277
lunghezze_alleliche(29) = 279
lunghezze_alleliche(30) = 281
lunghezze_alleliche(31) = 283

lunghezze_alleliche_base_N(0) = "N"
lunghezze_alleliche_base_N(1) = "N+2"
lunghezze_alleliche_base_N(2) = "N+4"
lunghezze_alleliche_base_N(3) = "N+8"
lunghezze_alleliche_base_N(4) = "N+10"
lunghezze_alleliche_base_N(5) = "N+12"
lunghezze_alleliche_base_N(6) = "N+14"
lunghezze_alleliche_base_N(7) = "N+16"
lunghezze_alleliche_base_N(8) = "N+17"
lunghezze_alleliche_base_N(9) = "N+18"
lunghezze_alleliche_base_N(10) = "N+19"
lunghezze_alleliche_base_N(11) = "N+20"
lunghezze_alleliche_base_N(12) = "N+22"
lunghezze_alleliche_base_N(13) = "N+24"
lunghezze_alleliche_base_N(14) = "N+26"
lunghezze_alleliche_base_N(15) = "N+28"
lunghezze_alleliche_base_N(16) = "N+30"
lunghezze_alleliche_base_N(17) = "N+32"
lunghezze_alleliche_base_N(18) = "N+34"
lunghezze_alleliche_base_N(19) = "N+36"
lunghezze_alleliche_base_N(20) = "N+38"
lunghezze_alleliche_base_N(21) = "N+40"
lunghezze_alleliche_base_N(22) = "N+42"
lunghezze_alleliche_base_N(23) = "N+44"
lunghezze_alleliche_base_N(24) = "N+46"
lunghezze_alleliche_base_N(25) = "N+48"
lunghezze_alleliche_base_N(26) = "N+54"
lunghezze_alleliche_base_N(27) = "N+56"
lunghezze_alleliche_base_N(28) = "N+58"
lunghezze_alleliche_base_N(29) = "N+60"
lunghezze_alleliche_base_N(30) = "N+62"
lunghezze_alleliche_base_N(31) = "N+64"

End Sub







