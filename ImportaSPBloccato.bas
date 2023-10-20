Attribute VB_Name = "ImportaSPBloccato"
Option Explicit

Sub ImportDataFromProtectedFile()
    ' Parametri configurabili (P)
    Dim inputFilePath As String
    Dim inputFileName As String
    Dim inputSheetName As String
    Dim startRow As Long
    Dim endRow As Long
    
    inputFilePath = "C:\NAS\Public\Fischetti\temp\" ' Directory dove si trova il file di input
    inputFileName = "sp.xlsx" ' Nome del file di input
    inputSheetName = "Sheet1" ' Nome del foglio nel file di input
    startRow = 50 ' Riga iniziale per la scansione della colonna più a destra
    endRow = 80 ' Riga finale per la scansione della colonna più a destra
    
    ' Altre variabili
    Dim srcWorkbook As Workbook
    Dim srcWorksheet As Worksheet
    Dim destWorkbook As Workbook
    Dim destWorksheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim tempLastCol As Long
    
    ' Imposta il workbook corrente come destinazione
    Set destWorkbook = ThisWorkbook
    
    ' Apri il file di input protetto (solo lettura)
    Set srcWorkbook = Workbooks.Open(inputFilePath & inputFileName, ReadOnly:=True)
    
    ' Accedi al foglio di lavoro nel file di input
    Set srcWorksheet = srcWorkbook.Sheets(inputSheetName)
    
    ' Crea un nuovo foglio nel file corrente
    Set destWorksheet = destWorkbook.Sheets.Add(After:=destWorkbook.Sheets(destWorkbook.Sheets.Count))
    
    ' Trova l'ultima riga con dati nel foglio di input
    lastRow = srcWorksheet.Cells(srcWorksheet.Rows.Count, 1).End(xlUp).Row
    
    ' Inizializza lastCol a 1
    lastCol = 1
    
    ' Trova la colonna più a destra che contiene dati tra le righe startRow e endRow
    For r = startRow To endRow
        tempLastCol = srcWorksheet.Cells(r, srcWorksheet.Columns.Count).End(xlToLeft).Column
        If tempLastCol > lastCol Then
            lastCol = tempLastCol
        End If
    Next r
    
    ' Copia i dati dal foglio di input al nuovo foglio
    srcWorksheet.Range(srcWorksheet.Cells(1, 1), srcWorksheet.Cells(lastRow, lastCol)).Copy destWorksheet.Cells(1, 1)
    
    ' Chiudi il file di input
    srcWorkbook.Close False
End Sub

