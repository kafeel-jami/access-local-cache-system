' ============================================================================
' MODULE: modConnection (Enhanced v2.0)
' PURPOSE: Optimized backend connection management with connection pooling
' SUPPORTS: Access (ACE/JET) and SQL Server backends
' ============================================================================
Option Compare Database
Option Explicit

' === Backend Type Enumeration ===
Public Enum eBackendType
    Backend_Access = 1
    Backend_SQLServer = 2
End Enum

' === Global Configuration ===
Public gBackendType As eBackendType
Private gADOConnection As ADODB.Connection
Private gConnectionAttempts As Long
Private gLastConnectionError As String

' === Constants ===
Private Const DB_PASSWORD As String = "XXXXXX"
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

' ============================================================================
' GetConnection - Returns connection (creates if needed)
' ============================================================================
Public Function GetConnection() As ADODB.Connection
    On Error GoTo ErrHandler
    
    EnsureConnection
    Set GetConnection = gADOConnection
    Exit Function
    
ErrHandler:
    gLastConnectionError = "GetConnection: " & Err.Description
    Err.Raise Err.Number, "modConnection.GetConnection", Err.Description
End Function

' ============================================================================
' EnsureConnection - Ensures connection is open and valid
' ============================================================================
Private Sub EnsureConnection()
    On Error GoTo ErrHandler
    
    ' Check if connection exists and is open
    If Not gADOConnection Is Nothing Then
        If gADOConnection.State = adStateOpen Then
            ' Test connection validity
            If TestConnection() Then Exit Sub
        End If
    End If
    
    ' Connection needs to be (re)established
    OpenConnection
    Exit Sub
    
ErrHandler:
    gLastConnectionError = "EnsureConnection: " & Err.Description
    Err.Raise Err.Number, "modConnection.EnsureConnection", Err.Description
End Sub

' ============================================================================
' OpenConnection - Opens new connection with retry logic
' ============================================================================
Private Sub OpenConnection()
    On Error GoTo ErrHandler
    
    Dim attempt As Long
    Dim connected As Boolean
    
    ' Close existing connection if any
    If Not gADOConnection Is Nothing Then
        If gADOConnection.State = adStateOpen Then gADOConnection.Close
    End If
    
    ' Create new connection object
    Set gADOConnection = New ADODB.Connection
    
    ' Retry loop
    For attempt = 1 To MAX_RETRY_ATTEMPTS
        On Error Resume Next
        
        gADOConnection.ConnectionString = GetConnectionString()
        gADOConnection.ConnectionTimeout = 30
        gADOConnection.CommandTimeout = 60
        gADOConnection.Open
        
        If Err.Number = 0 And gADOConnection.State = adStateOpen Then
            connected = True
            gConnectionAttempts = attempt
            Exit For
        Else
            gLastConnectionError = "Attempt " & attempt & ": " & Err.Description
            If attempt < MAX_RETRY_ATTEMPTS Then
                Sleep RETRY_DELAY_MS
            End If
        End If
        
        On Error GoTo ErrHandler
    Next attempt
    
    If Not connected Then
        Err.Raise vbObjectError + 1000, "modConnection.OpenConnection", _
                  "Failed to connect after " & MAX_RETRY_ATTEMPTS & " attempts. " & _
                  "Last error: " & gLastConnectionError
    End If
    
    Exit Sub
    
ErrHandler:
    gLastConnectionError = Err.Description
    Err.Raise Err.Number, "modConnection.OpenConnection", Err.Description
End Sub

' ============================================================================
' GetConnectionString - Returns appropriate connection string
' ============================================================================
Private Function GetConnectionString() As String
    On Error GoTo ErrHandler
    
    Select Case gBackendType
        Case Backend_SQLServer
            GetConnectionString = _
                "Provider=SQLOLEDB;" & _
                "Data Source=YOUR_SERVER;" & _
                "Initial Catalog=YOUR_DATABASE;" & _
                "Integrated Security=SSPI;" & _
                "Persist Security Info=False;" & _
                "Connection Timeout=30;"
                
        Case Backend_Access
            GetConnectionString = _
                "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & DATABASE_LOCATION & ";" & _
                "Jet OLEDB:Database Password=" & DB_PASSWORD & ";" & _
                "Mode=Share Deny None;" & _
                "Persist Security Info=False;"
                
        Case Else
            Err.Raise vbObjectError + 1001, , "Invalid backend type specified"
    End Select
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "modConnection.GetConnectionString", Err.Description
End Function

' ============================================================================
' TestConnection - Validates connection is working
' ============================================================================
Private Function TestConnection() As Boolean
    On Error Resume Next
    
    Dim testRS As ADODB.Recordset
    Set testRS = New ADODB.Recordset
    
    testRS.Open "SELECT 1", gADOConnection, adOpenForwardOnly, adLockReadOnly
    TestConnection = (Err.Number = 0 And Not testRS Is Nothing)
    
    If Not testRS Is Nothing Then
        If testRS.State = adStateOpen Then testRS.Close
        Set testRS = Nothing
    End If
End Function

' ============================================================================
' ExecuteRS - Returns connected recordset
' ============================================================================
Public Function ExecuteRS(ByVal SqlText As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    
    EnsureConnection
    
    Dim rs As New ADODB.Recordset
    rs.Open SqlText, gADOConnection, adOpenStatic, adLockReadOnly
    
    Set ExecuteRS = rs
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "modConnection.ExecuteRS", _
              "SQL: " & Left(SqlText, 100) & vbCrLf & "Error: " & Err.Description
End Function

' ============================================================================
' ExecuteRS_Disconnected - Returns disconnected recordset (for caching)
' ============================================================================
Public Function ExecuteRS_Disconnected(ByVal SqlText As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    
    EnsureConnection
    
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open SqlText, gADOConnection, adOpenStatic, adLockReadOnly
    
    ' Disconnect
    Set rs.ActiveConnection = Nothing
    
    Set ExecuteRS_Disconnected = rs
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "modConnection.ExecuteRS_Disconnected", _
              "SQL: " & Left(SqlText, 100) & vbCrLf & "Error: " & Err.Description
End Function

' ============================================================================
' ExecuteNonQuery - Executes action query
' ============================================================================
Public Function ExecuteNonQuery(ByVal SqlText As String) As Long
    On Error GoTo ErrHandler
    
    EnsureConnection
    
    Dim recordsAffected As Long
    gADOConnection.Execute SqlText, recordsAffected, adExecuteNoRecords
    
    ExecuteNonQuery = recordsAffected
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "modConnection.ExecuteNonQuery", _
              "SQL: " & Left(SqlText, 100) & vbCrLf & "Error: " & Err.Description
End Function

' ============================================================================
' ExecuteInTransaction - Executes with transaction support
' ============================================================================
Public Sub ExecuteInTransaction(ByVal SqlText As String)
    On Error GoTo ErrHandler
    
    EnsureConnection
    
    gADOConnection.BeginTrans
    gADOConnection.Execute SqlText, , adExecuteNoRecords
    gADOConnection.CommitTrans
    
    Exit Sub
    
ErrHandler:
    If Not gADOConnection Is Nothing Then
        If gADOConnection.State = adStateOpen Then
            gADOConnection.RollbackTrans
        End If
    End If
    Err.Raise Err.Number, "modConnection.ExecuteInTransaction", _
              "Transaction rolled back. " & Err.Description
End Sub

' ============================================================================
' ExecuteScalar - Returns single value
' ============================================================================
Public Function ExecuteScalar(ByVal SqlText As String) As Variant
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteRS(SqlText)
    
    If Not rs.EOF Then
        ExecuteScalar = rs.Fields(0).Value
    Else
        ExecuteScalar = Null
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    ExecuteScalar = Null
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Function

' ============================================================================
' CloseGlobalConnection - Cleanup
' ============================================================================
Public Sub CloseGlobalConnection()
    On Error Resume Next
    
    If Not gADOConnection Is Nothing Then
        If gADOConnection.State = adStateOpen Then
            gADOConnection.Close
        End If
        Set gADOConnection = Nothing
    End If
    
    gConnectionAttempts = 0
    gLastConnectionError = ""
End Sub

' ============================================================================
' InitializeConnection - Manual initialization
' ============================================================================
Public Sub InitializeConnection()
    EnsureConnection
End Sub

' ============================================================================
' GetConnectionInfo - Diagnostic information
' ============================================================================
Public Function GetConnectionInfo() As String
    On Error Resume Next
    
    Dim info As String
    info = "Backend: " & IIf(gBackendType = Backend_SQLServer, "SQL Server", "Access") & vbCrLf
    
    If Not gADOConnection Is Nothing Then
        info = info & "State: " & IIf(gADOConnection.State = adStateOpen, "Open", "Closed") & vbCrLf
        info = info & "Provider: " & gADOConnection.Provider & vbCrLf
    Else
        info = info & "State: Not initialized" & vbCrLf
    End If
    
    info = info & "Attempts: " & gConnectionAttempts & vbCrLf
    If gLastConnectionError <> "" Then
        info = info & "Last Error: " & gLastConnectionError
    End If
    
    GetConnectionInfo = info
End Function

' ============================================================================
' Sleep - Wait function for retry logic
' ============================================================================
Private Sub Sleep(ByVal milliseconds As Long)
    Dim endTime As Double
    endTime = Timer + (milliseconds / 1000)
    Do While Timer < endTime
        DoEvents
    Loop
End Sub
