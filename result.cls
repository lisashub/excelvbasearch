VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is a class to hold result details; an object is created from this class for each search result.
'************************************************************************************************

'Below are the names and data types for each class variable
Private rModuleName              As String
Private rModuleType              As String
Private rProcedureName           As String
Private rProcedureResultLine     As String
Private rSearchTerm              As String
'Remaining code defines how each variable is set and retrieved
'************************************************************************************************
Property Let ModuleName(ByVal value As String)
    rModuleName = value
End Property
Property Get ModuleName() As String
    ModuleName = rModuleName '!!!suspect this line and others like it are optional; try to test w/o!!!
End Property
'************************************************************************************************
Property Let ModuleType(ByVal value As String)
    rModuleType = value
End Property
Property Get ModuleType() As String
    ModuleType = rModuleType
End Property
'************************************************************************************************
Property Let procedureName(ByVal value As String)
    rProcedureName = value
End Property
Property Get procedureName() As String
    procedureName = rProcedureName
End Property
'************************************************************************************************
Property Let ProcedureResultLine(ByVal value As String)
    rProcedureResultLine = value
End Property
Property Get ProcedureResultLine() As String
    ProcedureResultLine = rProcedureResultLine
End Property
'************************************************************************************************
Property Let searchTerm(ByVal value As String)
    rSearchTerm = value
End Property
Property Get searchTerm() As String
    searchTerm = rSearchTerm
End Property

