VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_VMCNG4B9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' VARIANT USED AS STRING ARRAYS
' VOGLIAMO IMPLEMENTARE UNA MAPPA LUNGHEZZA ALLELICA -> LUNGHEZZA ALLELICA IN BASE N
' PER FARLO INIZIALIZZIAMO DUE ARRAY DOVE ARRAY_1[i] <-> ARRAY_2[i]

Dim lunghezze_alleliche(25) As Variant
Dim lunghezze_alleliche_base_N(25) As Variant

Public Property Get lunghezza_array() As Integer
    lunghezza_array = 25
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

lunghezze_alleliche(0) = 138
lunghezze_alleliche(1) = 146
lunghezze_alleliche(2) = 148
lunghezze_alleliche(3) = 150
lunghezze_alleliche(4) = 151
lunghezze_alleliche(5) = 152
lunghezze_alleliche(6) = 154
lunghezze_alleliche(7) = 156
lunghezze_alleliche(8) = 158
lunghezze_alleliche(9) = 160
lunghezze_alleliche(10) = 161
lunghezze_alleliche(11) = 162
lunghezze_alleliche(12) = 163
lunghezze_alleliche(13) = 164
lunghezze_alleliche(14) = 166
lunghezze_alleliche(15) = 167
lunghezze_alleliche(16) = 168
lunghezze_alleliche(17) = 170
lunghezze_alleliche(18) = 172
lunghezze_alleliche(19) = 174
lunghezze_alleliche(20) = 176
lunghezze_alleliche(21) = 178
lunghezze_alleliche(22) = 180
lunghezze_alleliche(23) = 184
lunghezze_alleliche(24) = 192

lunghezze_alleliche_base_N(0) = "N"
lunghezze_alleliche_base_N(1) = "N+8"
lunghezze_alleliche_base_N(2) = "N+10"
lunghezze_alleliche_base_N(3) = "N+12"
lunghezze_alleliche_base_N(4) = "N+13"
lunghezze_alleliche_base_N(5) = "N+14"
lunghezze_alleliche_base_N(6) = "N+16"
lunghezze_alleliche_base_N(7) = "N+18"
lunghezze_alleliche_base_N(8) = "N+20"
lunghezze_alleliche_base_N(9) = "N+22"
lunghezze_alleliche_base_N(10) = "N+23"
lunghezze_alleliche_base_N(11) = "N+24"
lunghezze_alleliche_base_N(12) = "N+25"
lunghezze_alleliche_base_N(13) = "N+26"
lunghezze_alleliche_base_N(14) = "N+28"
lunghezze_alleliche_base_N(15) = "N+29"
lunghezze_alleliche_base_N(16) = "N+30"
lunghezze_alleliche_base_N(17) = "N+32"
lunghezze_alleliche_base_N(18) = "N+34"
lunghezze_alleliche_base_N(19) = "N+36"
lunghezze_alleliche_base_N(20) = "N+38"
lunghezze_alleliche_base_N(21) = "N+40"
lunghezze_alleliche_base_N(22) = "N+42"
lunghezze_alleliche_base_N(23) = "N+46"
lunghezze_alleliche_base_N(24) = "N+54"
End Sub






