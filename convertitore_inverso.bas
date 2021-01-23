Attribute VB_Name = "convertitore_inverso"
' I DATI SARANNO SALVATI DIRETTAMENTE NEL FOGLIO DI OUTPUT
' NON INIZIALIZZIAMO QUINDI ALCUNA VARIABILE A RIGUARDO

' TABELLA PRESA DA INPUT
Dim tabella_input As Range

' CLASSE INIZIALIZZATA DI VOLTA IN VOLTA
Dim marcatore



Sub convertitore_inverso()

'CREA UN NUOVO FOGLIO EXCEL PRENDENDO IN INPUT DALL'UTENTE IL NOME DELLO STESSO
Dim nome_foglio As String
nome_foglio = InputBox("Inserisci nome del nuovo foglio Excel", "Convertitore da base N a numerica", "nome_file")
Sheets.Add().Name = nome_foglio



' PRENDIAMO ORA L'INTERA TABELLA CHE DOVE ESSERE OPPORTUNATAMENTE TRATTATA: LA PRIMA RIGA CONTIENE LE INTESTAZIONE DEI MARCATORI
' LA PRIMA COLONNA INVECE I NOMI DELLE VARIETà DEI VINI

Set tabella_input = Application.InputBox(Prompt:="seleziona le celle su cui applicare il filtro", Title:="Convertitore", _
                                             Default:="seleziona anche le intestazioni", Type:=8)


' PER PRIMA COSA tabella_output DEVE AVERE LA PRIMA COLONNA UGUALE, PERCHE' LE VARIETA' NON CAMBIANO
' STESSA COSA PER QUANTO RIGUARDA LA PRIMA RIGA
' COPIAMO QUINDI LA PRIMA COLONNA DI tabella_input IN tabella_output

'COPIAMO DUNQUE L'INTESTAZIONE DELLE RIGHE...
For i = 1 To tabella_input.Rows.Count Step 1
        Worksheets(nome_foglio).Cells(i, 1) = tabella_input(i, 1)
Next i

'...E DELLE COLONNE
For j = 1 To tabella_input.Columns.Count Step 1
        Worksheets(nome_foglio).Cells(1, j) = tabella_input(1, j)
Next j




' ORA LA PARTE CROCCANTE: DOBBIAMO DIVIDERE tabella_input IN COPPIE DI COLONNE ESTRAENDO IL NOME DEL MARCATORE E LA COPPIA STESSA
' PARTIAMO DA j=2 PERCHE' LA PRIMA COLONNA SONO IL NOME DELLE VARIETA' CHE ABBIAMO GIA' COPIATO

Dim nome_marcatore As String

For j = 2 To tabella_input.Columns.Count Step 2
    
    ' GET NOME DEL MARCATORE
    If tabella_input.Cells(1, j) <> "" Then
        nome_marcatore = tabella_input.Cells(1, j)
    Else
        nome_marcatore = tabella_input.Cells(1, j + 1)
    End If
    
    ' INIZIALIZZIAMO LA CLASSE CORRETTA
    Select Case UCase(nome_marcatore)
    
    Case "ISV2"
    Set marcatore = New class_ISV2
    
    Case "ISV4"
    Set marcatore = New class_ISV4
    
    Case "VMCNG4B9"
    Set marcatore = New class_VMCNG4B9
    
    Case "VRZAG62"
    Set marcatore = New class_VrZAG62
    
    Case "VRZAG79"
    Set marcatore = New class_VrZAG79
    
    Case "VVMD25"
    Set marcatore = New class_VVMD25
    
    Case "VVMD27"
    Set marcatore = New class_VVMD27
    
    Case "VVMD28"
    Set marcatore = New class_VVMD28
    
    Case "VVMD32"
    Set marcatore = New class_VVMD32
    
    Case "VVMD5"
    Set marcatore = New class_VVMD5
    
    Case "VVMD7"
    Set marcatore = New class_VVMD7
    
    Case "VVS2"
    Set marcatore = New class_VVS2
    
    Case Else
    MsgBox ("incontrato marcatore sconosciuto: " + nome_marcatore)
    Exit Sub
    End Select
    
        For i = 2 To tabella_input.Rows.Count Step 1

            ' APPLICHIAMO LA CONVERSIONE
            ' DUE PERCHE' j VIENE INCREMENTATA DI DUE VALORI ALLA VOLTA
        
            Worksheets(nome_foglio).Cells(i, j) = marcatore.Converti(tabella_input(i, j))
            Worksheets(nome_foglio).Cells(i, j + 1) = marcatore.Converti(tabella_input(i, j + 1))
            

    Next i
Next j





End Sub

