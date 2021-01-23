VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_ISV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' VARIANT USED AS STRING ARRAYS
' VOGLIAMO IMPLEMENTARE UNA MAPPA LUNGHEZZA ALLELICA -> LUNGHEZZA ALLELICA IN BASE N
' PER FARLO INIZIALIZZIAMO DUE ARRAY DOVE ARRAY_1[i] <-> ARRAY_2[i]

Dim lunghezze_alleliche(26) As Variant
Dim lunghezze_alleliche_base_N(26) As Variant

Public Property Get lunghezza_array() As Integer
    lunghezza_array = 26
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

If index = 0 Then
    Converti = "-"
Else
    Converti = lunghezze_alleliche(index)
End If
   

End Function

'DEFAULT METHOD TO INITIALIZE CLASS VARIABLES
Private Sub Class_Initialize()

lunghezze_alleliche(0) = 121
lunghezze_alleliche(1) = 125
lunghezze_alleliche(2) = 127
lunghezze_alleliche(3) = 129
lunghezze_alleliche(4) = 131
lunghezze_alleliche(5) = 133
lunghezze_alleliche(6) = 135
lunghezze_alleliche(7) = 137
lunghezze_alleliche(8) = 139
lunghezze_alleliche(9) = 141
lunghezze_alleliche(10) = 143
lunghezze_alleliche(11) = 145
lunghezze_alleliche(12) = 147
lunghezze_alleliche(13) = 151
lunghezze_alleliche(14) = 153
lunghezze_alleliche(15) = 155
lunghezze_alleliche(16) = 157
lunghezze_alleliche(17) = 159
lunghezze_alleliche(18) = 161
lunghezze_alleliche(19) = 163
lunghezze_alleliche(20) = 165
lunghezze_alleliche(21) = 167
lunghezze_alleliche(22) = 169
lunghezze_alleliche(23) = 171
lunghezze_alleliche(24) = 173
lunghezze_alleliche(25) = 175

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
lunghezze_alleliche_base_N(13) = "N+30"
lunghezze_alleliche_base_N(14) = "N+32"
lunghezze_alleliche_base_N(15) = "N+34"
lunghezze_alleliche_base_N(16) = "N+36"
lunghezze_alleliche_base_N(17) = "N+38"
lunghezze_alleliche_base_N(18) = "N+40"
lunghezze_alleliche_base_N(19) = "N+42"
lunghezze_alleliche_base_N(20) = "N+44"
lunghezze_alleliche_base_N(21) = "N+46"
lunghezze_alleliche_base_N(22) = "N+48"
lunghezze_alleliche_base_N(23) = "N+50"
lunghezze_alleliche_base_N(24) = "N+52"
lunghezze_alleliche_base_N(25) = "N+54"

End Sub





