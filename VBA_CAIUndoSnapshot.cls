'''''' 模块: VBA/CAIUndoSnapshot  类型: CAIUndoSnapshot.cls  行数: 38
Attribute VB_Name = "CAIUndoSnapshot"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public SnapshotBook As Object
Public WorkbookNames As Collection
Public WorksheetNames As Collection
Public Addresses As Collection
Public SnapshotSheets As Collection
Public FormulaValues As Collection
Public ClearFilters As Collection

Private Sub Class_Initialize()
    Set WorkbookNames = New Collection
    Set WorksheetNames = New Collection
    Set Addresses = New Collection
    Set SnapshotSheets = New Collection
    Set FormulaValues = New Collection
    Set ClearFilters = New Collection
End Sub

Public Property Get EntryCount() As Long
    EntryCount = WorkbookNames.count
End Property

Public Sub CloseSnapshot()
    On Error Resume Next
    If Not SnapshotBook Is Nothing Then SnapshotBook.Close False
    Set SnapshotBook = Nothing
    On Error GoTo 0
End Sub

