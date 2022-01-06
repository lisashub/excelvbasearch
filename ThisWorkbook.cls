VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Test()

    Dim searchObject        As New Search
    Dim resultsCollection   As New Collection
    Dim resultsObject       As result
    Dim resultsSheet        As Worksheet
    Dim procedureName       As String
    Dim currentRow          As Long
    Dim searchTerm          As String
    
    
    '************************************************************************************************
    'Collects filename from user
    '************************************************************************************************
    searchObject.FileName = InputBox("Enter the full path of the local file to check (e.g. C:\test\testworkbook.xlsm:)")

    '************************************************************************************************
    'Checks integrity of provided file path
    '************************************************************************************************
    If searchObject.FileName = vbNullString Then
        MsgBox ("No filename provided. Exiting program.")
        Exit Sub
    End If
    
    If Len(Dir(searchObject.FileName)) = 0 Then
        MsgBox ("Sorry, file does not exist. Please check provided path and try again." & vbNewLine & "Exiting program.")
        Exit Sub
    End If
    
    '************************************************************************************************
    'Adds search terms to search term array
    '************************************************************************************************
    Do
        searchTerm = InputBox("Please enter a term you would like to search for. Leave blank when finished.")
        If Not searchTerm = "" Then
            searchObject.addTerm searchTerm
        End If
    Loop Until searchTerm = ""
    
    '************************************************************************************************
    'Calls procedure to search specified file for search terms
    '************************************************************************************************
    Set resultsCollection = searchObject.CaptureSearchResults

    '************************************************************************************************
    'Populates worksheet with search results
    '************************************************************************************************
    If Not resultsCollection Is Nothing Then
        Range("A4").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection).ClearContents
        Range(Selection, Selection).ClearFormats
        
        currentRow = 4
        
        With Sheet1
            .Cells(currentRow, 1).value = "Search Term "
            .Cells(currentRow, 2).value = "Module/Object Name "
            .Cells(currentRow, 3).value = "Subroutine/Function Name "
            .Cells(currentRow, 4).value = "Line Text "
            .Cells(currentRow, 5).value = "File Name "
            currentRow = currentRow + 1
        End With
        
        For Each resultsObject In resultsCollection
        
            If resultsObject.procedureName = vbNullString Then
                procedureName = "(None)"
            Else:
                procedureName = resultsObject.procedureName
            End If
        
            With Sheet1
                .Cells(currentRow, 1).value = resultsObject.searchTerm
                .Cells(currentRow, 2).value = resultsObject.ModuleName
                .Cells(currentRow, 3).value = procedureName
                .Cells(currentRow, 4).value = resultsObject.ProcedureResultLine
                .Cells(currentRow, 5).value = searchObject.FileName
                currentRow = currentRow + 1
            End With
        Next
        
        ActiveSheet.Range("A4:E" & (currentRow - 1)).Select
        Range(Selection, Selection).Columns.AutoFit
        ActiveSheet.Range("A5:E" & (currentRow - 1)).Sort Key1:=Range("A5"), Order1:=xlAscending
    
        ActiveSheet.Range("A4:E4").Font.Bold = True
        With ActiveSheet.Range("A4:E4").Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        
        ActiveSheet.Range("A5:E" & (currentRow - 1)).Font.Bold = False
        
    End If
    
    '************************************************************************************************
    'Clears generated objects in memory
    '************************************************************************************************
    Set searchObject = Nothing
    Set resultsCollection = Nothing
    Set oResult = Nothing
    
    MsgBox ("Search complete.")
    
    ThisWorkbook.VBProject.VBE.MainWindow.Visible = True
    
End Sub

