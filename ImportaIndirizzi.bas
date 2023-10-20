Attribute VB_Name = "ImportaIndirizzi"
Const FILE_PATH As String = "C:\NAS\Public\Fischetti\concordati\TECNOSERVICE SRL\Relazioni e SP\Tecnoservice.prn"
Const StopLine = 9999

Option Explicit

Sub ProcessTextFile()
    Dim ws As Worksheet
    Dim mCron, mCred As String, mDom As String, mCodFisc As String
    Dim currentLine As String, labelSection As String, dataSection As String
    Dim lineNum As Long, fileLineNum As Long
    Dim flagE1 As Boolean, flagDom As Boolean
    Dim validLabels As Collection
    Dim mPecCred As String, mPecDom As String
    Dim flagPecCred As Boolean, flagPecDom As Boolean

    Const ColCron As Long = 1
    Const ColCred As Long = 2
    Const ColDomic As Long = 3
    Const ColPecCred As Long = 4
    Const ColPecDom As Long = 5
    Const ColCodFisc As Long = 6
    Const LabelDataSeparator As Long = 24  ' Posizione del separatore tra etichette e dati
    Const widthCol As Long = 36 ' Larghezza colonna in unità di carattere, circa 250 pixel



    Set validLabels = New Collection
    validLabels.Add "00-0"
    validLabels.Add "PEC Creditore"
    validLabels.Add "PEC Domiciliatario"
    validLabels.Add "Codice fiscale"
    
    
    ' === Section 1: Initialization ===
        Set ws = ThisWorkbook.Sheets.Add
        ws.Cells(1, ColCron).Value = "Cron"
        ws.Cells(1, ColCred).Value = "Creditore"
        ws.Cells(1, ColDomic).Value = "Domiciliatario"
        ws.Cells(1, ColPecCred).Value = "PEC Cred"
        ws.Cells(1, ColPecDom).Value = "PEC Domic"
        ws.Cells(1, ColCodFisc).Value = "Codice Fiscale"
        
        ws.Columns("A:F").ColumnWidth = widthCol
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(255, 255, 0)
        ws.Rows("2:2").Select
        ActiveWindow.FreezePanes = True
        
        
        lineNum = 2
        flagE1 = False
        flagDom = False
        fileLineNum = 0

    ' === Section 2: File Reading ===
    Dim fileNum As Long
    fileNum = FreeFile
    Open FILE_PATH For Input As #fileNum
    
    Do Until EOF(fileNum)
        fileLineNum = fileLineNum + 1
        If fileLineNum = StopLine Then Stop
        
        Line Input #fileNum, currentLine
        labelSection = Left(currentLine, LabelDataSeparator)
        dataSection = Mid(currentLine, LabelDataSeparator + 1)
        
        ' === Section 3: Label Handling ===
        ' E1: "00-0"
        If Left(labelSection, 4) = "00-0" Then
            Dim nextFourChars As String
            nextFourChars = Mid(labelSection, 5, 4)
            If IsNumeric(nextFourChars) Then
                mCron = Left(labelSection, 4) & nextFourChars ' Set mCron
                ' Only reset mCred if it's a new "00-0" block
                If mCred = "" Then
                    mCred = dataSection
                Else
                    mCred = mCred & " " & dataSection
                End If
                mDom = ""
                flagE1 = True
                flagDom = False
            End If
        End If
            ' E4: "Codice Fiscale"
        If Left(labelSection, 14) = "Codice fiscale" Then
            mCodFisc = dataSection
            ' Scrive tutte le variabili in Excel utilizzando le costanti per le colonne
            ws.Cells(lineNum, ColCron).Value = mCron  ' Nuova colonna
            ws.Cells(lineNum, ColCred).Value = StripBlank(mCred)
            ws.Cells(lineNum, ColDomic).Value = StripBlank(mDom)
            ws.Cells(lineNum, ColPecCred).Value = mPecCred  ' Nessuna rimozione di spazi per l'email
            ws.Cells(lineNum, ColPecDom).Value = mPecDom  ' Nessuna rimozione di spazi per l'email
            ws.Cells(lineNum, ColCodFisc).Value = "'" & StripBlank(mCodFisc)
            mCron = ""
            mCred = ""
            mDom = ""
            mPecCred = ""
            mPecDom = ""
            
            lineNum = lineNum + 1  ' Passa alla riga successiva
        End If
' === Section 4: Data Handling ===
        If flagE1 Then
            If Left(dataSection, 3) <> "c/o" And Not IsInCollection(labelSection, validLabels) And dataSection <> "" Then
                If mCred <> "" Then mCred = mCred & " "
                mCred = mCred & dataSection
            ElseIf dataSection = "" Then
                flagE1 = False
                ' Handle the case where dataSection is empty and mCred should be considered complete
                ' Code to handle this specific case can be added here
            End If
        End If
        
        If Left(dataSection, 3) = "c/o" Then
            mDom = Mid(dataSection, 4)
            flagDom = True
            flagE1 = False
        ElseIf flagDom Then
            If Not IsInCollection(labelSection, validLabels) And dataSection <> "" Then
                If mDom <> "" Then mDom = mDom & " "
                mDom = mDom & dataSection
            ElseIf dataSection = "" Then
                flagDom = False
            End If
        End If
        
        ' Handle "PEC Creditore"
        If Left(labelSection, 13) = "PEC Creditore" Then
            mPecCred = dataSection
            flagPecCred = True
        ElseIf flagPecCred Then
            If IsInCollection(labelSection, validLabels) Or dataSection = "" Then
                flagPecCred = False
            Else
                mPecCred = mPecCred & dataSection
'                ElseIf dataSection = "" Or IsInCollection(labelSection, validLabels) Then
'                flagPecCred = False
            End If
        End If
        
        ' Handle "PEC Domiciliatario"
        If Left(labelSection, 18) = "PEC Domiciliatario" Then
            mPecDom = dataSection
            flagPecDom = True
        ElseIf flagPecDom And Not IsInCollection(labelSection, validLabels) And dataSection <> "" Then
            mPecDom = mPecDom & dataSection
            ElseIf dataSection = "" Or IsInCollection(labelSection, validLabels) Then
            flagPecDom = False
        End If


Loop
    
    ' === Section 5: Cleanup ===
    Close #fileNum
End Sub


' Funzione per rimuovere spazi extra, lasciando un solo spazio tra le parole
Function StripBlank(str As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\s+"
    StripBlank = Trim(regex.Replace(str, " "))
End Function


Function IsInCollection(item As String, col As Collection) As Boolean
    Dim elem As Variant
    IsInCollection = False
    On Error Resume Next
    For Each elem In col
        If Left(item, Len(elem)) = elem Then
            IsInCollection = True
            Exit Function
        End If
    Next elem
End Function

