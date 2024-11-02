Attribute VB_Name = "verificaPR"
Option Explicit

Sub verificaFattura()
    Dim newInvoice As Worksheet
    Dim volAccisabile As Double
    Dim count As Integer
    
    'Creo il foglio dove salvare i calcoli di verifica
    'Set newInvoice = ThisWorkbook.Worksheets.Add
    Set newInvoice = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    newInvoice.Move Before:=ThisWorkbook.Sheets("HOME")

    On Error Resume Next
    newInvoice.name = Mid(Worksheets("xmlFattura").Range("B2").Value, 4, 6)
    On Error GoTo 0
    
    'Imposto i titoli della sezione alta
    With newInvoice
        .Range("A1").Value = "PRODOTTO"
        .Range("B1").Value = "CL"
        .Range("C1").Value = "BT CRT"
        .Range("D1").Value = "CASSE"
        .Range("E1").Value = "BOTT."
        .Range("F1").Value = "IN FATTURA"
        .Range("G1").Value = "NET BT."
        .Range("H1").Value = "IF/CS"
        .Range("I1").Value = "IF/BT"
        .Range("J1").Value = "CS/BT"
        .Range("K1").Value = "ACCISA"
        .Range("L1").Value = "CONTRAS."
        With .Range("A1:L1").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("A1:L1").Font.Bold = True
        .Range("A1:L1").Font.Italic = True
        .Range("A1:L1").EntireColumn.HorizontalAlignment = xlCenter
        .Range("A1:L1").Interior.Color = RGB(219, 238, 243)
        .Columns("A").ColumnWidth = 28
        .Columns("B").ColumnWidth = 3.29
        .Columns("C").ColumnWidth = 8.71
        .Columns("D").ColumnWidth = 10.14
        .Columns("E").ColumnWidth = 8.71
        .Columns("F").ColumnWidth = 14.57
        .Columns("G").ColumnWidth = 10.71
        .Columns("H").ColumnWidth = 10.71
        .Columns("I").ColumnWidth = 10.71
        .Columns("J").ColumnWidth = 10.71
        .Columns("K").ColumnWidth = 11.14
        .Columns("L").ColumnWidth = 11.14
        .Columns("M").ColumnWidth = 9
        .Columns("N").ColumnWidth = 9
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = 11
        .Columns("Q").ColumnWidth = 13
    End With
    
    Dim tbl As ListObject
    Dim riga As ListRow
    Dim firstFreeRow As Integer
    
    firstFreeRow = 2
    count = 0
    
    Set tbl = Worksheets("xmlFattura").ListObjects("fattura")
    
    For Each riga In tbl.ListRows
        If riga.Range(1, 3) <> "" Then
            newInvoice.Range("A" & firstFreeRow).Value = EstrapolaNOME(riga.Range(1, 4))
            newInvoice.Range("B" & firstFreeRow).Value = EstrapolaCL(riga.Range(1, 4))
            newInvoice.Range("C" & firstFreeRow).Value = EstrapolaBT(riga.Range(1, 4))
            newInvoice.Range("D" & firstFreeRow).Value = riga.Range(1, 5)
            newInvoice.Range("E" & firstFreeRow).FormulaR1C1 = "=RC[-1]*RC[-2]"
            newInvoice.Range("F" & firstFreeRow).Value = riga.Range(1, 7)
            newInvoice.Range("G" & firstFreeRow).FormulaR1C1 = "=RC[-1]/RC[-3]/RC[-4]"
            volAccisabile = EstrapolaVol(riga.Range(1, 3))
            If volAccisabile <> 0 Then
            'IL VALORE NELLA COLONNA H RISULTA ERRATO. RIVEDERE I CALCOLI O IL VOL%
                newInvoice.Range("H" & firstFreeRow).Value = Round((EstrapolaVol(riga.Range(1, 3)) * 10.3552 * newInvoice.Range("B" & firstFreeRow).Value / 100 + 0.047) * newInvoice.Range("C" & firstFreeRow).Value, 2)
                'newInvoice.Range("H" & firstFreeRow).Value = (EstrapolaVol(riga.Range(1, 3)) * 10.3552 * newInvoice.Range("B" & firstFreeRow).Value / 100 + 0.047) * newInvoice.Range("C" & firstFreeRow).Value
                
                newInvoice.Range("I" & firstFreeRow).FormulaR1C1 = "=RC[-1]/RC[-6]-0.047"
                newInvoice.Range("J" & firstFreeRow).Value = 0.047
            Else
                newInvoice.Range("H" & firstFreeRow).Value = ""
                newInvoice.Range("I" & firstFreeRow).Value = ""
                newInvoice.Range("J" & firstFreeRow).Value = ""
            End If
            newInvoice.Range("K" & firstFreeRow).FormulaR1C1 = "=RC[-2]*RC[-6]"
            newInvoice.Range("L" & firstFreeRow).FormulaR1C1 = "=RC[-2]*RC[-7]"
            firstFreeRow = firstFreeRow + 1
        End If
        count = count + 1
    Next riga
    
    newInvoice.Range("F2:L" & firstFreeRow).NumberFormat = "€#,##0.00;€-#,##0.00;€               -;@"
    newInvoice.Range("J2:J" & firstFreeRow).NumberFormat = "€#,##0.000;€-#,##0.000;€               -;@"
    
    'Calcolo i totali
    newInvoice.Range("F" & firstFreeRow).Formula = "=SUM(F2:F" & firstFreeRow - 1 & ")"
    newInvoice.Range("K" & firstFreeRow).Formula = "=SUM(K2:K" & firstFreeRow - 1 & ")"
    newInvoice.Range("L" & firstFreeRow).Formula = "=SUM(L2:L" & firstFreeRow - 1 & ")"
    With newInvoice.Range("F" & firstFreeRow)
        .Font.Bold = True
        .NumberFormat = "€#,##0.00"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
    End With
    With newInvoice.Range("K" & firstFreeRow)
        .Font.Bold = True
        .NumberFormat = "€#,##0.00"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
    End With
    With newInvoice.Range("L" & firstFreeRow)
        .Font.Bold = True
        .NumberFormat = "€#,##0.00"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
    End With
    
    firstFreeRow = firstFreeRow + 2

    'Imposto i titoli della sezione bassa
    With newInvoice
        .Range("A" & firstFreeRow).Value = "PRODOTTO"
        .Range("B" & firstFreeRow).Value = "CL"
        .Range("C" & firstFreeRow).Value = "3%  F.P. BORDER TRIM."
        .Range("D" & firstFreeRow).Value = "3%  F.P. BORDER TRIM."
        .Range("E" & firstFreeRow).Value = "3%  F.P. BORDER TRIM."
        .Range("F" & firstFreeRow).Value = "3%  F.P. BORDER TRIM."
        .Range("G" & firstFreeRow).Value = "5% F.P. DOMESTIC TRIM."
        .Range("H" & firstFreeRow).Value = "5% F.P. DOMESTIC TRIM."
        .Range("I" & firstFreeRow).Value = "6% F.P. DOMESTIC TRIM."
        .Range("J" & firstFreeRow).Value = "6% F.P. DOMESTIC TRIM."
        .Range("K" & firstFreeRow).Value = "PROMO"
        .Range("L" & firstFreeRow).Value = "PROMO"
        .Range("M" & firstFreeRow).Value = "NETTO"
        .Range("N" & firstFreeRow).Value = "PREZZI PATTUITI"
        .Range("O" & firstFreeRow).Value = "DIFFER."
        .Range("P" & firstFreeRow).Value = "VALORE PROMO/NC"
        With .Range("A" & firstFreeRow, "P" & firstFreeRow).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("A" & firstFreeRow, "P" & firstFreeRow).Font.Bold = True
        .Range("A" & firstFreeRow, "P" & firstFreeRow).Font.Italic = True
        .Range("A" & firstFreeRow, "P" & firstFreeRow).EntireColumn.HorizontalAlignment = xlCenter
        .Range("A" & firstFreeRow, "P" & firstFreeRow).Interior.Color = RGB(219, 238, 243)
        .Range("A" & firstFreeRow, "P" & firstFreeRow).EntireRow.RowHeight = 43
        .Range("A" & firstFreeRow, "P" & firstFreeRow).EntireRow.VerticalAlignment = xlCenter
        .Range("A" & firstFreeRow, "P" & firstFreeRow).WrapText = True
        firstFreeRow = firstFreeRow + 1
    End With
    
    Dim start2ndtbl As Integer
    
    start2ndtbl = firstFreeRow
    
    'COMPILO LA SECONDA TABELLA PER IL CALCOLO DEL PREZZO CONCORDATO
    With newInvoice
        Dim i As Integer
        Dim sconto As String
        Dim promo As Double
        For i = 1 To count
            .Range("A" & firstFreeRow).Value = .Range("A" & i + 1).Value
            .Range("B" & firstFreeRow).Value = .Range("B" & i + 1).Value
'            MsgBox Worksheets("xmlFattura").Range("C" & i + 1).Value
            sconto = EstrapolaTipo(Worksheets("xmlFattura").Range("C" & i + 1).Value)
            If sconto = "BORDER" Then
                .Range("C" & firstFreeRow).Formula = "=(G" & i + 1 & "-I" & i + 1 & "-J" & i + 1 & ")*3%"
                .Range("D" & firstFreeRow).Formula = "=E" & i + 1 & "*C" & firstFreeRow & ""
                .Range("E" & firstFreeRow).Formula = "=(G" & i + 1 & "-I" & i + 1 & "-J" & i + 1 & ")*3%"
                .Range("F" & firstFreeRow).Formula = "=E" & i + 1 & "*E" & firstFreeRow & ""
            ElseIf sconto = "DOMESTIC" Then
                .Range("G" & firstFreeRow).Formula = "=(G" & i + 1 & "-I" & i + 1 & "-J" & i + 1 & ")*5%"
                .Range("H" & firstFreeRow).Formula = "=E" & i + 1 & "*G" & firstFreeRow & ""
                .Range("I" & firstFreeRow).Formula = "=(G" & i + 1 & "-I" & i + 1 & "-J" & i + 1 & ")*6%"
                .Range("J" & firstFreeRow).Formula = "=E" & i + 1 & "*I" & firstFreeRow & ""
            Else
                MsgBox "Errore Inatteso"
                Exit Sub
            End If
            promo = EstrapolaPromo(Worksheets("xmlFattura").Range("C" & i + 1).Value)
            If promo <> 0 Then
                .Range("K" & firstFreeRow).Value = promo
            End If
            .Range("L" & firstFreeRow).Formula = "=E" & i + 1 & "*K" & firstFreeRow & ""
            .Range("M" & firstFreeRow).Formula = "=G" & i + 1 & "-C" & firstFreeRow & "-E" & firstFreeRow & "-G" & firstFreeRow & "-I" & firstFreeRow & "-K" & firstFreeRow & ""
            .Range("N" & firstFreeRow).FormulaR1C1 = "=RC[-1]"
            .Range("O" & firstFreeRow).FormulaR1C1 = "=RC[-1]-RC[-2]"
            .Range("P" & firstFreeRow).Formula = "=O" & firstFreeRow & "*E" & i + 1 & ""
            .Range("Q" & firstFreeRow).Formula = "=M" & firstFreeRow & "*E" & i + 1 & ""
            .Range("C" & firstFreeRow & ":Q" & firstFreeRow).NumberFormat = "€#,##0.00"
            .Range("L" & firstFreeRow).NumberFormat = "€#,##0.00;€-#,##0.00;€               -;@"
            .Range("O" & firstFreeRow).NumberFormat = "€#,##0.00;€-#,##0.00;€               -;@"
            .Range("P" & firstFreeRow).NumberFormat = "€#,##0.00;€-#,##0.00;€               -;@"

            .Range("L" & firstFreeRow).Font.Color = RGB(255, 0, 0)
            .Range("L" & firstFreeRow).Font.Bold = True
            .Range("N" & firstFreeRow).Font.Bold = True
            .Range("N" & firstFreeRow).Interior.Color = RGB(255, 255, 0)
            firstFreeRow = firstFreeRow + 1
        Next i
    End With
    
    'imposto i totali
    With newInvoice
        .Range("D" & firstFreeRow).Value = 1
        .Range("D" & firstFreeRow).Formula = "=SUM(D" & start2ndtbl & ":D" & firstFreeRow - 1 & ")"
        .Range("F" & firstFreeRow).Formula = "=SUM(F" & start2ndtbl & ":F" & firstFreeRow - 1 & ")"
        .Range("H" & firstFreeRow).Formula = "=SUM(H" & start2ndtbl & ":H" & firstFreeRow - 1 & ")"
        .Range("J" & firstFreeRow).Formula = "=SUM(J" & start2ndtbl & ":J" & firstFreeRow - 1 & ")"
        .Range("L" & firstFreeRow).Formula = "=SUM(L" & start2ndtbl & ":L" & firstFreeRow - 1 & ")"
        .Range("P" & firstFreeRow).Formula = "=SUM(P" & start2ndtbl & ":P" & firstFreeRow - 1 & ")"
        .Range("Q" & firstFreeRow).Formula = "=SUM(Q" & start2ndtbl & ":Q" & firstFreeRow - 1 & ")+D" & firstFreeRow & "+F" & firstFreeRow & "+H" & firstFreeRow & "+J" & firstFreeRow & "+L" & firstFreeRow & ""
        .Range("A" & firstFreeRow).EntireRow.Font.Bold = True
        .Range("A" & firstFreeRow, "P" & firstFreeRow).Font.Color = RGB(255, 0, 0)
        With newInvoice.Range("D" & firstFreeRow)
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        End With
        With newInvoice.Range("F" & firstFreeRow)
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        End With
        With newInvoice.Range("H" & firstFreeRow)
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        End With
        With newInvoice.Range("J" & firstFreeRow)
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        End With
        With newInvoice.Range("L" & firstFreeRow)
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        End With
        With newInvoice.Range("P" & firstFreeRow)
            .Font.Bold = True
            .Interior.Color = RGB(255, 255, 0)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        End With
        With newInvoice.Range("Q" & firstFreeRow)
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        End With
    End With
    Application.Calculation = xlCalculationAutomatic
End Sub

Function EstrapolaNOME(descrizione As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim name As String

    ' Crea un oggetto Regex
    Set regex = CreateObject("VBScript.RegExp")
    'regex.Pattern = "(?:\d{3}CL|CL\d{3})X\d+\s*([A-Za-z\s]+)"
    regex.Pattern = "(?:\d{3}CL|CL\d{3})X\d+\s*([A-Za-z0-9\s]+)"
    regex.Global = False

    ' Esegui la ricerca
    Set matches = regex.Execute(descrizione)
    
    ' Se non ci sono corrispondenze, prova il secondo pattern per due cifre
    If matches.count = 0 Then
        regex.Pattern = "CL\d{2}X\d+\s*([A-Za-z\s]+)"
        Set matches = regex.Execute(descrizione)
    End If

    ' Inizializza la variabile centilitri
    name = ""

    ' Itera sulle corrispondenze trovate
    For Each match In matches
        name = match.SubMatches(0)
        Exit For ' Esce dopo la prima corrispondenza trovata
    Next match

    ' Restituisce i centilitri
    EstrapolaNOME = name
End Function

Function EstrapolaCL(descrizione As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim centilitri As String

    ' Crea un oggetto Regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(?:CL(\d{3})|(\d{3})CL)"
    regex.Global = False

    ' Esegui la ricerca
    Set matches = regex.Execute(descrizione)
    
    ' Se non ci sono corrispondenze, prova il secondo pattern per due cifre
    If matches.count = 0 Then
        regex.Pattern = "CL(\d{2})"
        Set matches = regex.Execute(descrizione)
    End If

    ' Inizializza la variabile centilitri
    centilitri = ""

    ' Itera sulle corrispondenze trovate
    For Each match In matches
        ' Verifica quale submatch ha catturato il valore
        If match.SubMatches(0) <> "" Then
            centilitri = match.SubMatches(0)
        ElseIf match.SubMatches.count > 1 And match.SubMatches(1) <> "" Then
            centilitri = match.SubMatches(1)
        End If
        Exit For ' Esce dopo la prima corrispondenza trovata
    Next match

    ' Restituisce i centilitri
    EstrapolaCL = centilitri
End Function

Function EstrapolaBT(descrizione As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim bottles As String

    ' Crea un oggetto Regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(?:CL\d{3}X(\d+)|\d{3}CLX(\d+))"
    regex.Global = False

    ' Esegui la ricerca
    Set matches = regex.Execute(descrizione)
    
    ' Se non ci sono corrispondenze, prova il secondo pattern per due cifre
    If matches.count = 0 Then
        regex.Pattern = "CL\d{2}X(\d+)"
        Set matches = regex.Execute(descrizione)
    End If

    ' Inizializza la variabile bottiglie
    bottles = ""

    ' Itera sulle corrispondenze trovate
    For Each match In matches
        ' Verifica quale submatch ha catturato il valore
        If match.SubMatches(0) <> "" Then
            bottles = match.SubMatches(0)
        ElseIf match.SubMatches.count > 1 And match.SubMatches(1) <> "" Then
            bottles = match.SubMatches(1)
        End If
        Exit For ' Esce dopo la prima corrispondenza trovata
    Next match

    ' Restituisce il numero di bottiglie
    EstrapolaBT = bottles
End Function

Function EstrapolaVol(codProdotto As String) As Double
    Dim c As Range
    Dim vol As Double

    On Error GoTo ErrorHandler ' Imposta il gestore degli errori

    With Worksheets("prodotti")
        Set c = .Range("A:A").Find(codProdotto, LookIn:=xlValues, lookat:=xlWhole)
        If Not c Is Nothing Then
            vol = c.Offset(0, 3).Value
        Else
            Err.Raise vbObjectError + 513, "EstrapolaVol", "Prodotto non riconosciuto"
        End If
    End With

    EstrapolaVol = vol
    Exit Function ' Esce dalla funzione se tutto è andato bene

ErrorHandler:
    ' Gestisce l'errore e mostra il messaggio di errore
    MsgBox "Prodotto non riconosciuto, verificare nella scheda PRODOTTI.", vbExclamation, "Errore"
    EstrapolaVol = 0 ' Imposta il valore di ritorno a 0 o un valore predefinito
End Function

Function EstrapolaTipo(codProdotto As String) As String
    Dim c As Range
    Dim tipo As String

    On Error GoTo ErrorHandler ' Imposta il gestore degli errori

    With Worksheets("prodotti")
        Set c = .Range("A:A").Find(codProdotto, LookIn:=xlValues, lookat:=xlWhole)
        If Not c Is Nothing Then
            tipo = c.Offset(0, 4).Value
        Else
            Err.Raise vbObjectError + 513, "EstrapolaTipo", "Prodotto non riconosciuto"
        End If
    End With

    EstrapolaTipo = tipo
    Exit Function ' Esce dalla funzione se tutto è andato bene

ErrorHandler:
    ' Gestisce l'errore e mostra il messaggio di errore
    MsgBox "Prodotto non riconosciuto, verificare nella scheda PRODOTTI.", vbExclamation, "Errore"
    EstrapolaTipo = 0 ' Imposta il valore di ritorno a 0 o un valore predefinito
End Function

Function EstrapolaPromo(codProdotto As String) As Double
    Dim c As Range
    Dim promo As Double

    On Error GoTo ErrorHandler ' Imposta il gestore degli errori

    With Worksheets("prodotti")
        Set c = .Range("A:A").Find(codProdotto, LookIn:=xlValues, lookat:=xlWhole)
        If Not c Is Nothing Then
            promo = c.Offset(0, 5).Value
        Else
            Err.Raise vbObjectError + 513, "EstrapolaPromo", "Prodotto non riconosciuto"
        End If
    End With

    EstrapolaPromo = promo
    Exit Function ' Esce dalla funzione se tutto è andato bene

ErrorHandler:
    ' Gestisce l'errore e mostra il messaggio di errore
    MsgBox "Prodotto non riconosciuto, verificare nella scheda PRODOTTI.", vbExclamation, "Errore"
    EstrapolaPromo = 0 ' Imposta il valore di ritorno a 0 o un valore predefinito
End Function

