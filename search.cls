VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private searchTerms()
Private rFileName            As String
Private termCounter          As Integer
Private excelObject          As Excel.Application
Private excelWorkbook        As Workbook
Private msgBoxResult         As String
Private msgBoxResult1        As String
' Creates a class object with termCounter set to zero
'**************************************************************************************
Private Sub Class_Initialize()
    termCounter = 0
End Sub
' Sets the user-specified filename
'**************************************************************************************
Property Let fileName(ByVal value As String)
    rFileName = value
End Property
' Retrieves the user-specified filename
'**************************************************************************************
Property Get fileName() As String
    fileName = rFileName
End Property
' Clears word counter and search term array.
'**************************************************************************************
Public Sub clearTerms()
    termCounter = 0
    ReDim searchTerms(0 To termCounter)
End Sub
' Adds a term to the search term array; increases term counter value by 1.
' **************************************************************************************
Public Sub addTerm(ByVal value As String)
    ReDim Preserve searchTerms(0 To termCounter)
    searchTerms(termCounter) = value
    termCounter = termCounter + 1
End Sub
' Creates a file system and file object, activate/captures results from search function,
' sets status bar to custom message, and returns collection of search results to client.
' **************************************************************************************
Public Function captureSearchResults() As Collection

On Error GoTo ERR_HANDLER:

    Dim fso          As New FileSystemObject
    Dim file         As file
    Dim blnCheckFile As Boolean
    
    Set captureSearchResults = Nothing
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(rFileName)
    Application.StatusBar = "Checking File: " & file.Path & " : {" & Now() & "}"
    Set captureSearchResults = generateResultsCollection(file)

EXIT_HERE:
    Application.StatusBar = Empty
    Set fso = Nothing
    Set file = Nothing
    Exit Function

ERR_HANDLER:
    Debug.Print Err.Description
    GoTo EXIT_HERE
    
End Function


' Creates a file system and file object, captures results from search subroutine,
' sets status bar to custom message, and returns collection of search results.
' **************************************************************************************
Private Function generateResultsCollection(ByVal oFile As file) As Collection
    
On Error GoTo ERR_HANDLER:

    Dim VBComponent             As VBIDE.VBComponent
    Dim VBCodeModule            As VBIDE.CodeModule
    Dim VBComponentList         As VBIDE.VBComponents
    Dim ModuleName              As String
    Dim procedureType           As VBIDE.vbext_ProcKind
    Dim procedureName           As String
    Dim ProcedureResultLine     As String
    Dim ProcedureTotalLines     As Long
    Dim procedureStartColumn    As Long
    Dim procedureEndColumn      As Long
    Dim termFound               As Boolean
    Dim resultsCollection       As New Collection
    Dim resultObject            As result
    Dim resultCounter           As Long
    Dim moduleStartLine         As Long
    Dim moduleEndLine           As Long
    Dim moduleStartColumn       As Long
    Dim moduleEndColumn         As Long

    Set generateResultsCollection = Nothing
    Set excelObject = CreateObject("Excel.Application")
    Set excelWorkbook = excelObject.Workbooks.Open(fileName:=fileName, UpdateLinks:=False, ReadOnly:=True, local:=True)
    Set VBComponentList = excelWorkbook.VBProject.VBComponents
    
    If VBComponentList Is Nothing Then
        GoTo NO_COMPONENTS
    End If
    
    Application.VBE.MainWindow.Visible = False
    
    For Each VBComponent In VBComponentList
            DoEvents
            Select Case VBComponent.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document

                Set VBCodeModule = VBComponent.CodeModule
                ModuleName = VBCodeModule.Name
                With VBCodeModule
                
                    For resultCounter = LBound(searchTerms) To UBound(searchTerms)
                        moduleStartLine = 1
                        moduleEndLine = .CountOfLines
                        moduleStartColumn = 1
                        moduleEndColumn = 255
                    
                        termFound = .Find(Target:=searchTerms(resultCounter), StartLine:=moduleStartLine, StartColumn:=moduleStartColumn, _
                                         EndLine:=moduleEndLine, EndColumn:=moduleEndColumn, wholeword:=True, MatchCase:=False, patternsearch:=False)
                        Do Until Not termFound
                            DoEvents
                            procedureName = .ProcOfLine(moduleStartLine, procedureType)
            
                            Set resultObject = New result
                            resultObject.searchTerm = searchTerms(resultCounter)
                            resultObject.procedureName = procedureName
                            resultObject.ModuleName = ModuleName
                            resultObject.ProcedureResultLine = Trim(VBComponent.CodeModule.Lines(moduleStartLine, 1))
                            resultsCollection.Add resultObject
                            moduleEndLine = .CountOfLines
                            moduleStartColumn = moduleEndColumn + 1
                            moduleEndColumn = 255
                            termFound = .Find(Target:=searchTerms(resultCounter), StartLine:=moduleStartLine, StartColumn:=moduleStartColumn, _
                                             EndLine:=moduleEndLine, EndColumn:=moduleEndColumn, wholeword:=True, MatchCase:=False, patternsearch:=False)
                        Loop
                                    
                    Next resultCounter
                End With
            End Select
    Next VBComponent
    Application.VBE.MainWindow.Visible = True
    Set generateResultsCollection = resultsCollection

NO_COMPONENTS:
    GoTo EXIT_HERE

EXIT_HERE:
    Application.VBE.MainWindow.Visible = True
    Set VBComponent = Nothing
    Set VBCodeModule = Nothing
    Set VBComponentList = Nothing
    Set resultsCollection = Nothing
    Set resultObject = Nothing
    excelWorkbook.Close False
    Set excelWorkbook = Nothing
    Exit Function

ERR_HANDLER:
    Debug.Print Err.Description
    GoTo EXIT_HERE
    
End Function
