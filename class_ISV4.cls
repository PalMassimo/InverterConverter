VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_ISV4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' VARIANT USED AS STRING ARRAYS
' VOGLIAMO IMPLEMENTARE UNA MAPPA LUNGHEZZA ALLELICA -> LUNGHEZZA ALLELICA IN BASE N
' PER FARLO INIZIALIZZIAMO DUE ARRAY DOVE ARRAY_1[i] <-> ARRAY_2[i]

Dim lunghezze_alleliche(14) As Variant
Dim lunghezze_alleliche_base_N(14) As Variant

Public Property Get lunghezza_array() As Integer
    lunghezza_array = 14
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

lunghezze_alleliche(0) = 169
lunghezze_alleliche(1) = 173
lunghezze_alleliche(2) = 175
lunghezze_alleliche(3) = 177
lunghezze_alleliche(4) = 179
lunghezze_alleliche(5) = 181
lunghezze_alleliche(6) = 183
lunghezze_alleliche(7) = 185
lunghezze_alleliche(8) = 187
lunghezze_alleliche(9) = 189
lunghezze_alleliche(10) = 191
lunghezze_alleliche(11) = 193
lunghezze_alleliche(12) = 195
lunghezze_alleliche(13) = 197

lunghezze_alleliche_base_N(0) = "N"
lunghezze_alleliche_base_N(1) = "N+4"
lunghezze_alleliche_base_N(2) = "N+6"
lunghezze_alleliche_base_N(3) = "N+8"
lunghezze_alleliche_base_N(4) = "N+10"
lunghezze_alleliche_base_N(5) = "N+12"
lunghezze_alleliche_base_N(6) = "N+14"
lunghezze_alleliche_base_N(7) = "N+16"
lunghezze_alleliche_base_N(8) = "N+18"
lunghezze_alleliche_base_N(9) = "N+20"
lunghezze_alleliche_base_N(10) = "N+22"
lunghezze_alleliche_base_N(11) = "N+24"
lunghezze_alleliche_base_N(12) = "N+26"
lunghezze_alleliche_base_N(13) = "N+28"

End Sub






