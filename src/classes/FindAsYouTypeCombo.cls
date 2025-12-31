' ============================================================================
' CLASS: FindAsYouTypeCombo (Optimized v2.0)
' PURPOSE: High-performance Find-As-You-Type combo with memory caching
' AUTHOR: MajP (Enhanced)
' ============================================================================
Option Compare Database
Option Explicit

' === Events ===
Private WithEvents mCombo As Access.ComboBox
Private WithEvents mForm As Access.Form

' === Configuration ===
Private mFilterFieldName As String
Private mSearchType As searchType
Private mHandleArrows As Boolean
Private mAutoCompleteEnabled As Boolean
Private mRowSource As String

' === Recordsets ===
Private mRsOriginalList As DAO.Recordset

' === Memory Cache (for performance) ===
Private mCacheDict As Object
Private mCacheBuilt As Boolean
Private mBoundColumnIndex As Integer
Private mFilterFieldIndex As Integer

' === Constants ===
Private Const CACHE_SIZE_LIMIT As Long = 10000

Public Enum searchType
    anywhereinstring = 0
    FromBeginning = 1
End Enum

' ============================================================================
' INITIALIZATION
' ============================================================================

Public Sub InitalizeFilterCombo( _
    TheComboBox As Access.ComboBox, _
    Optional FilterFieldName As String = "All_Fields", _
    Optional TheSearchType As searchType = searchType.anywhereinstring, _
    Optional handleArrows As Boolean = True)
    
    On Error GoTo ErrHandler
    
    ' Validate combo box
    If Not TheComboBox.RowSourceType = "Table/Query" Then
        MsgBox "This class requires a Table/Query row source type.", vbExclamation
        Exit Sub
    End If
    
    ' Set references
    Set mCombo = TheComboBox
    
    ' Get form reference (handles tab controls)
    Dim tempControl As Object
    Set tempControl = TheComboBox
    Do
        Set tempControl = tempControl.Parent
    Loop Until Left(TypeName(tempControl), 4) = "Form"
    Set mForm = tempControl
    
    ' Configure behavior
    mHandleArrows = handleArrows
    mAutoCompleteEnabled = True
    mFilterFieldName = FilterFieldName
    mSearchType = TheSearchType
    mCacheBuilt = False
    
    ' Set event procedures
    mCombo.OnClick = "[Event Procedure]"
    mCombo.OnChange = "[Event Procedure]"
    mCombo.AfterUpdate = "[Event Procedure]"
    mCombo.OnGotFocus = "[Event Procedure]"
    mForm.OnCurrent = "[Event Procedure]"
    mForm.OnClose = "[Event Procedure]"
    
    If mHandleArrows Then
        mCombo.OnKeyDown = "[Event Procedure]"
    End If
    
    ' Configure combo
    With mCombo
        .AutoExpand = False
    End With
    
    ' Initialize recordset and cache
    If mCombo.RowSource <> "" Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset(TheComboBox.RowSource)
        Set mCombo.Recordset = rs
        mRowSource = mCombo.RowSource
    End If
    
    If Not mCombo.Recordset Is Nothing Then
        Set mRsOriginalList = mCombo.Recordset.Clone
        mBoundColumnIndex = mCombo.BoundColumn - 1
        If mBoundColumnIndex < 0 Then mBoundColumnIndex = 0
        BuildCache
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Initialization error: " & Err.Description, vbCritical
End Sub

' ============================================================================
' CACHE BUILDING
' ============================================================================

Private Sub BuildCache()
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    
    Set rs = mRsOriginalList
    
    ' Don't cache empty recordsets
    If rs.EOF And rs.BOF Then
        mCacheBuilt = False
        Exit Sub
    End If
    
    ' Don't cache very large datasets (performance degradation)
    rs.MoveLast
    If rs.RecordCount > CACHE_SIZE_LIMIT Then
        Debug.Print "FAYT> Cache disabled for large dataset (" & rs.RecordCount & " records)"
        mCacheBuilt = False
        rs.MoveFirst
        Exit Sub
    End If
    rs.MoveFirst
    
    ' Create dictionary
    Set mCacheDict = CreateObject("Scripting.Dictionary")
    mCacheDict.CompareMode = 1 ' Text compare, case-insensitive
    
    ' Find filter field index
    mFilterFieldIndex = -1
    If mFilterFieldName <> "" And mFilterFieldName <> "All_Fields" Then
        Dim i As Integer
        For i = 0 To rs.Fields.Count - 1
            If rs.Fields(i).name = mFilterFieldName Then
                mFilterFieldIndex = i
                Exit For
            End If
        Next i
    End If
    
    ' Build cache
    Dim strKey As String
    Dim strSearchText As String
    
    Do Until rs.EOF
        ' Use bound column as key
        strKey = Nz(rs.Fields(mBoundColumnIndex).Value, "")
        
        ' Build searchable text
        If mFilterFieldIndex >= 0 Then
            ' Specific field
            strSearchText = UCase(Nz(rs.Fields(mFilterFieldIndex).Value, ""))
        Else
            ' All text fields
            strSearchText = ""
            For Each fld In rs.Fields
                If fld.Type = dbText Or fld.Type = dbMemo Then
                    strSearchText = strSearchText & UCase(Nz(fld.Value, "")) & "|"
                End If
            Next fld
        End If
        
        ' Store in dictionary
        If strKey <> "" And Not mCacheDict.Exists(strKey) Then
            mCacheDict.Add strKey, strSearchText
        End If
        
        rs.MoveNext
    Loop
    
    mCacheBuilt = True
    Debug.Print "FAYT> Cache built: " & mCacheDict.Count & " entries"
    
    Exit Sub
    
ErrHandler:
    mCacheBuilt = False
    Debug.Print "FAYT> Cache build error: " & Err.Description
End Sub

' ============================================================================
' FILTERING
' ============================================================================

Private Sub FilterList()
    On Error GoTo ErrHandler
    
    ' Ensure focus
    If Left(TypeName(mCombo), 4) = "Form" Then
        mCombo.SetFocus
    End If
    
    If Not mAutoCompleteEnabled Then Exit Sub
    
    Dim strText As String
    Dim strFilter As String
    Dim rsTemp As DAO.Recordset
    
    strText = mCombo.Text
    
    ' Build filter
    strFilter = GetFilter(strText)
    
    ' Apply filter
    Set rsTemp = mRsOriginalList.OpenRecordset
    rsTemp.Filter = strFilter
    Set rsTemp = rsTemp.OpenRecordset
    
    ' Check results
    If Not (rsTemp.EOF And rsTemp.BOF) Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    Else
        Beep
    End If
    
    ' Update combo
    Set mCombo.Recordset = rsTemp
    
    If rsTemp.RecordCount > 0 Then
        If Nz(mCombo.Value, "") <> Nz(mCombo.Text, "") Then
            mCombo.Dropdown
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 3061 Then
        Debug.Print "FAYT> Invalid field name in filter"
    Else
        Debug.Print "FAYT> Filter error: " & Err.Description
    End If
End Sub

Private Function GetFilter(TheText As String) As String
    On Error Resume Next
    
    Dim fld As DAO.Field
    Dim rs As DAO.Recordset
    Dim strFilter As String
    Dim strLike As String
    
    ' Escape special characters
    TheText = Replace(TheText, "'", "''")
    TheText = Replace(TheText, "#", "[#]")
    
    ' Build LIKE pattern
    If mSearchType = FromBeginning Then
        strLike = " LIKE '"
    Else
        strLike = " LIKE '*"
    End If
    
    Set rs = mCombo.Recordset
    
    ' Build filter string
    If mFilterFieldName = "" Or mFilterFieldName = "All_Fields" Then
        ' Filter all text fields
        For Each fld In rs.Fields
            If fld.Type = dbMemo Or fld.Type = dbText Then
                If strFilter = "" Then
                    strFilter = "[" & fld.name & "]" & strLike & TheText & "*'"
                Else
                    strFilter = strFilter & " OR [" & fld.name & "]" & strLike & TheText & "*'"
                End If
            End If
        Next fld
    Else
        ' Filter specific field
        strFilter = "[" & mFilterFieldName & "]" & strLike & TheText & "*'"
    End If
    
    GetFilter = strFilter
End Function

Private Sub UnFilterList()
    On Error Resume Next
    Set mCombo.Recordset = mRsOriginalList
End Sub

' ============================================================================
' PROPERTIES
' ============================================================================

Public Property Get RowSource() As String
    RowSource = mRowSource
End Property

Public Property Let RowSource(ByVal NewRowSource As String)
    On Error Resume Next
    
    Dim rs As DAO.Recordset
    
    mRowSource = NewRowSource
    mCombo.RowSource = NewRowSource
    
    Set rs = CurrentDb.OpenRecordset(NewRowSource)
    Set mCombo.Recordset = rs
    Set mRsOriginalList = mCombo.Recordset.Clone
    
    mBoundColumnIndex = mCombo.BoundColumn - 1
    If mBoundColumnIndex < 0 Then mBoundColumnIndex = 0
    
    BuildCache
End Property

Public Property Get FilterComboBox() As Access.ComboBox
    Set FilterComboBox = mCombo
End Property

Public Property Set FilterComboBox(TheComboBox As Access.ComboBox)
    Set mCombo = TheComboBox
End Property

Public Property Get FilterFieldName() As String
    FilterFieldName = mFilterFieldName
End Property

Public Property Let FilterFieldName(ByVal theFieldName As String)
    mFilterFieldName = theFieldName
    BuildCache ' Rebuild cache if field changes
End Property

Public Property Get handleArrows() As Boolean
    handleArrows = mHandleArrows
End Property

Public Property Let handleArrows(ByVal TheValue As Boolean)
    mHandleArrows = TheValue
End Property

Public Property Get FilterType() As searchType
    FilterType = mSearchType
End Property

Public Property Let FilterType(ByVal TheFilterType As searchType)
    mSearchType = TheFilterType
End Property

' ============================================================================
' METHODS
' ============================================================================

Public Sub Requery()
    ' Use class requery, not control requery
    Me.RowSource = Me.FilterComboBox.RowSource
End Sub

Public Sub RefreshCombo()
    ' Refresh the combo box and rebuild cache
    ' This is especially useful when module-level variables are destroyed due to code errors
    On Error GoTo ErrHandler

    ' Rebuild the original recordset
    If mRowSource <> "" Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset(mRowSource)
        Set mCombo.Recordset = rs
        Set mRsOriginalList = mCombo.Recordset.Clone
        mBoundColumnIndex = mCombo.BoundColumn - 1
        If mBoundColumnIndex < 0 Then mBoundColumnIndex = 0

        ' Rebuild cache
        BuildCache

        ' Reset the filter
        UnFilterList
    End If

    Exit Sub

ErrHandler:
    MsgBox "RefreshCombo error: " & Err.Description, vbCritical
End Sub

' ============================================================================
' EVENT HANDLERS
' ============================================================================

Private Sub mCombo_Change()
    Call FilterList
    mAutoCompleteEnabled = True
End Sub

Private Sub mCombo_AfterUpdate()
    mAutoCompleteEnabled = True
    UnFilterList
End Sub

Private Sub mCombo_Click()
    mAutoCompleteEnabled = False
End Sub

Private Sub mCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Handle F5 key regardless of mHandleArrows setting
    If KeyCode = vbKeyF5 Then
        RefreshCombo
        Exit Sub
    End If

    ' Only handle arrow keys if mHandleArrows is True
    If Not mHandleArrows Then Exit Sub

    Select Case KeyCode
        Case vbKeyDown, vbKeyUp, vbKeyReturn, vbKeyPageDown, vbKeyPageUp
            mCombo.Dropdown
            mAutoCompleteEnabled = False
        Case Else
            mAutoCompleteEnabled = True
    End Select
End Sub

Private Sub mCombo_GotFocus()
    ' Optional: Pre-expand on focus
End Sub

Private Sub mForm_Current()
    Call UnFilterList
End Sub

Private Sub mForm_Close()
    Call Class_Terminate
End Sub

' ============================================================================
' CLEANUP
' ============================================================================

Private Sub Class_Terminate()
    On Error Resume Next
    
    ' Release objects
    Set mForm = Nothing
    Set mCombo = Nothing
    
    If Not mRsOriginalList Is Nothing Then
        If mRsOriginalList.State = 1 Then mRsOriginalList.Close
        Set mRsOriginalList = Nothing
    End If
    
    Set mCacheDict = Nothing
    mCacheBuilt = False
End Sub
