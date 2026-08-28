Attribute VB_Name = "DoppioCache"
''
' Doppio Cache Module
' Caching system for API responses to reduce redundant calls
'
' @module DoppioCache
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
'
' CROSS-PLATFORM: Works on both Mac and Windows
' - Uses Collection for cross-platform compatibility
' - No Scripting.Dictionary (Windows-only)
''
Option Explicit

' =============================================================================
' PRIVATE STATE
' =============================================================================

' Cache storage - use Collection for cross-platform compatibility
Private m_Cache As Collection
Private m_CacheKeys As Collection  ' Parallel collection for key tracking
Private m_Initialized As Boolean

' Legacy record cache
Private m_RecordCache As Collection

' Cache settings
Private Const MAX_CACHE_SIZE As Long = 100
Private Const DEFAULT_TTL_SECONDS As Long = 300  ' 5 minutes

' Cache entry structure indices
' Each entry is stored as: Array(Key, Data, Timestamp, TTL)
Private Const ENTRY_KEY As Integer = 0
Private Const ENTRY_DATA As Integer = 1
Private Const ENTRY_TIMESTAMP As Integer = 2
Private Const ENTRY_TTL As Integer = 3

' =============================================================================
' INITIALIZATION
' =============================================================================

''
' Initialize the cache
''
Public Sub Cache_InitializeCache()
    Set m_Cache = New Collection
    Set m_CacheKeys = New Collection
    m_Initialized = True
End Sub

''
' Ensure cache is initialized
''
Private Sub EnsureInitialized()
    If Not m_Initialized Or m_Cache Is Nothing Then
        Cache_InitializeCache
    End If
End Sub

''
' Check if key exists in cache (cross-platform)
''
Private Function CacheKeyExists(cacheKey As String) As Boolean
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To m_CacheKeys.count
        If m_CacheKeys(i) = cacheKey Then
            CacheKeyExists = True
            Exit Function
        End If
    Next i
    CacheKeyExists = False
    On Error GoTo 0
End Function

''
' Get index of key in cache
''
Private Function GetCacheKeyIndex(cacheKey As String) As Long
    Dim i As Long
    
    For i = 1 To m_CacheKeys.count
        If m_CacheKeys(i) = cacheKey Then
            GetCacheKeyIndex = i
            Exit Function
        End If
    Next i
    GetCacheKeyIndex = 0
End Function

''
' Remove item from cache by key
''
Private Sub RemoveCacheItem(cacheKey As String)
    Dim idx As Long
    idx = GetCacheKeyIndex(cacheKey)
    
    If idx > 0 Then
        m_Cache.Remove idx
        m_CacheKeys.Remove idx
    End If
End Sub

' =============================================================================
' PUBLIC API
' =============================================================================

''
' Try to get a response from cache
' @param cacheKey - Key to look up
' @param response - Output: The cached response if found
' @return Boolean - True if found and not expired
''
Public Function Cache_TryGetFromCache(cacheKey As String, ByRef response As apiResponse) As Boolean
    Dim entry As Variant
    Dim cacheTimestamp As Date
    Dim ttl As Long
    Dim idx As Long
    Dim json As Object
    
    EnsureInitialized
    
    idx = GetCacheKeyIndex(cacheKey)
    If idx = 0 Then
        Cache_TryGetFromCache = False
        Exit Function
    End If
    
    entry = m_Cache(idx)
    cacheTimestamp = entry(ENTRY_TIMESTAMP)
    ttl = entry(ENTRY_TTL)
    
    ' Check if expired
    If DateDiff("s", cacheTimestamp, Now) > ttl Then
        ' Remove expired entry
        RemoveCacheItem cacheKey
        Cache_TryGetFromCache = False
        Exit Function
    End If
    
    ' Return cached data
    response.data = entry(ENTRY_DATA)
    response.success = True
    response.status = ApiStatus_Success
    
    #If DEBUG_MODE Then
        Debug.Print "Cache_TryGetFromCache: HIT for " & Left(cacheKey, 50)
        Debug.Print "Cache_TryGetFromCache: Data length = " & Len(response.data)
    #End If
    
    ' Parse the cached JSON if it's MI data
    If Left(response.data, 1) = "{" Then
        On Error Resume Next
        Set json = JsonConverter.ParseJson(response.data)
        
        If Err.Number <> 0 Then
            #If DEBUG_MODE Then
                Debug.Print "Cache_TryGetFromCache: JSON parse error - " & Err.description
            #End If
            Err.Clear
        ElseIf Not json Is Nothing Then
            If json.exists("results") Then
                Set response.results = json.item("results")
                If Not response.results Is Nothing Then
                    If response.results.count > 0 Then
                        Set response.records = response.results(1).item("records")
                        If Not response.records Is Nothing Then
                            response.recordCount = response.records.count
                            #If DEBUG_MODE Then
                                Debug.Print "Cache_TryGetFromCache: Parsed " & response.recordCount & " records"
                            #End If
                        End If
                    End If
                End If
            Else
                Set response.results = json
                #If DEBUG_MODE Then
                    Debug.Print "Cache_TryGetFromCache: Set Results to json object (no 'results' key)"
                #End If
            End If
        End If
        On Error GoTo 0
    End If
    
    Cache_TryGetFromCache = True
End Function

''
' Store a response in cache
' @param cacheKey - Key to store under
' @param response - Response to cache
' @param Optional ttlSeconds - Time to live in seconds
''
Public Sub Cache_StoreInCache(cacheKey As String, response As apiResponse, _
                        Optional ttlSeconds As Long = 0)
    Dim entry(0 To 3) As Variant
    
    EnsureInitialized
    
    If ttlSeconds = 0 Then ttlSeconds = DEFAULT_TTL_SECONDS
    
    ' Check cache size and evict if necessary
    If m_Cache.count >= MAX_CACHE_SIZE Then
        EvictOldestEntries 10  ' Remove 10 oldest entries
    End If
    
    ' Create cache entry
    entry(ENTRY_KEY) = cacheKey
    entry(ENTRY_DATA) = response.data
    entry(ENTRY_TIMESTAMP) = Now
    entry(ENTRY_TTL) = ttlSeconds
    
    ' Remove if exists, then add
    If CacheKeyExists(cacheKey) Then
        RemoveCacheItem cacheKey
    End If
    
    m_Cache.Add entry
    m_CacheKeys.Add cacheKey
    
    #If DEBUG_MODE Then
        Debug.Print "Cache STORE: " & cacheKey
    #End If
End Sub

''
' Store raw data in cache (for compatibility with existing code)
' @param cacheKey - Key to store under
' @param data - Raw data string
' @param Optional ttlSeconds - Time to live in seconds
''
Public Sub Cache_StoreDataInCache(cacheKey As String, data As String, _
                            Optional ttlSeconds As Long = 0)
    Dim response As apiResponse
    response.data = data
    response.success = True
    response.status = ApiStatus_Success
    
    Cache_StoreInCache cacheKey, response, ttlSeconds
End Sub

''
' Check if a key exists in cache (and is not expired)
' @param cacheKey - Key to check
' @return Boolean - True if exists and valid
''
Public Function Cache_ExistsInCache(cacheKey As String) As Boolean
    Dim entry As Variant
    Dim cacheTimestamp As Date
    Dim ttl As Long
    Dim idx As Long
    
    EnsureInitialized
    
    idx = GetCacheKeyIndex(cacheKey)
    If idx = 0 Then
        Cache_ExistsInCache = False
        Exit Function
    End If
    
    entry = m_Cache(idx)
    cacheTimestamp = entry(ENTRY_TIMESTAMP)
    ttl = entry(ENTRY_TTL)
    
    ' Check if expired
    If DateDiff("s", cacheTimestamp, Now) > ttl Then
        RemoveCacheItem cacheKey
        Cache_ExistsInCache = False
    Else
        Cache_ExistsInCache = True
    End If
End Function

''
' Remove a specific entry from cache
' @param cacheKey - Key to remove
''
Public Sub Cache_RemoveFromCache(cacheKey As String)
    EnsureInitialized
    RemoveCacheItem cacheKey
End Sub

''
' Clear all entries from cache
''
Public Sub Cache_ClearCache()
    EnsureInitialized
    Set m_Cache = New Collection
    Set m_CacheKeys = New Collection
    
    ' Also clear the Cache sheet if it exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cache")
    If Not ws Is Nothing Then
        ws.Rows("2:" & ws.Rows.count).ClearContents
    End If
    On Error GoTo 0
    
    #If DEBUG_MODE Then
        Debug.Print "Cache CLEARED"
    #End If
End Sub

''
' Get the number of entries in cache
' @return Long - Number of entries
''
Public Function Cache_CacheCount() As Long
    EnsureInitialized
    Cache_CacheCount = m_Cache.count
End Function

''
' Display cache contents on the Cache sheet
''
Public Sub Cache_DisplayCache()
    Dim ws As Worksheet
    Dim entry As Variant
    Dim i As Long
    Dim row As Long
    
    EnsureInitialized
    
    ' Get or create Cache sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cache")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Cache"
    End If
    On Error GoTo 0
    
    ws.Visible = xlSheetVisible
    ws.Cells.ClearContents
    
    ' Headers
    ws.Cells(1, 1).value = "Key"
    ws.Cells(1, 2).value = "Timestamp"
    ws.Cells(1, 3).value = "TTL (sec)"
    ws.Cells(1, 4).value = "Data Preview"
    ws.Range("A1:D1").Font.Bold = True
    
    ' Data
    row = 2
    
    For i = 1 To m_Cache.count
        entry = m_Cache(i)
        ws.Cells(row, 1).value = entry(ENTRY_KEY)
        ws.Cells(row, 2).value = entry(ENTRY_TIMESTAMP)
        ws.Cells(row, 3).value = entry(ENTRY_TTL)
        ws.Cells(row, 4).value = Left(entry(ENTRY_DATA), 200)
        row = row + 1
    Next i
    
    ws.columns.AutoFit
    ws.Activate
End Sub

' =============================================================================
' PRIVATE HELPERS
' =============================================================================

''
' Evict oldest entries from cache
' @param count - Number of entries to remove
''
Private Sub EvictOldestEntries(count As Long)
    Dim timestamps() As Date
    Dim indices() As Long
    Dim i As Long
    Dim j As Long
    Dim tempIdx As Long
    Dim tempDate As Date
    Dim entry As Variant
    Dim removeCount As Long
    
    If m_Cache.count = 0 Then Exit Sub
    
    ReDim timestamps(1 To m_Cache.count)
    ReDim indices(1 To m_Cache.count)
    
    ' Extract timestamps and indices
    For i = 1 To m_Cache.count
        entry = m_Cache(i)
        timestamps(i) = entry(ENTRY_TIMESTAMP)
        indices(i) = i
    Next i
    
    ' Simple bubble sort by timestamp (oldest first)
    For i = 1 To UBound(timestamps) - 1
        For j = i + 1 To UBound(timestamps)
            If timestamps(i) > timestamps(j) Then
                tempDate = timestamps(i)
                timestamps(i) = timestamps(j)
                timestamps(j) = tempDate
                
                tempIdx = indices(i)
                indices(i) = indices(j)
                indices(j) = tempIdx
            End If
        Next j
    Next i
    
    ' Remove oldest entries (remove from end to preserve indices)
    removeCount = WorksheetFunction.Min(count, m_Cache.count)
    
    ' Sort indices in descending order for safe removal
    For i = 1 To removeCount - 1
        For j = i + 1 To removeCount
            If indices(i) < indices(j) Then
                tempIdx = indices(i)
                indices(i) = indices(j)
                indices(j) = tempIdx
            End If
        Next j
    Next i
    
    ' Remove entries
    For i = 1 To removeCount
        If indices(i) <= m_Cache.count Then
            m_Cache.Remove indices(i)
            m_CacheKeys.Remove indices(i)
        End If
    Next i
End Sub

' =============================================================================
' PERSISTENCE (OPTIONAL)
' =============================================================================

''
' Save cache to the Cache sheet for persistence across sessions
''
Public Sub Cache_SaveCacheToSheet()
    Dim ws As Worksheet
    Dim entry As Variant
    Dim i As Long
    Dim row As Long
    
    EnsureInitialized
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cache")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ws.Rows("2:" & ws.Rows.count).ClearContents
    
    row = 2
    
    For i = 1 To m_Cache.count
        entry = m_Cache(i)
        ws.Cells(row, 1).value = entry(ENTRY_KEY)
        ws.Cells(row, 2).value = entry(ENTRY_TIMESTAMP)
        ws.Cells(row, 3).value = entry(ENTRY_TTL)
        ws.Cells(row, 4).value = entry(ENTRY_DATA)
        row = row + 1
    Next i
End Sub

''
' Load cache from the Cache sheet
''
Public Sub Cache_LoadCacheFromSheet()
    Dim ws As Worksheet
    Dim row As Long
    Dim entry(0 To 3) As Variant
    Dim cacheKey As String
    Dim cacheTimestamp As Date
    Dim ttl As Long
    
    EnsureInitialized
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cache")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    
    row = 2
    Do While ws.Cells(row, 1).value <> ""
        cacheKey = ws.Cells(row, 1).value
        
        On Error Resume Next
        cacheTimestamp = ws.Cells(row, 2).value
        ttl = CLng(ws.Cells(row, 3).value)
        On Error GoTo 0
        
        ' Only load if not expired
        If DateDiff("s", cacheTimestamp, Now) <= ttl Then
            entry(ENTRY_KEY) = cacheKey
            entry(ENTRY_DATA) = ws.Cells(row, 4).value
            entry(ENTRY_TIMESTAMP) = cacheTimestamp
            entry(ENTRY_TTL) = ttl
            
            m_Cache.Add entry
            m_CacheKeys.Add cacheKey
        End If
        
        row = row + 1
    Loop
    
    #If DEBUG_MODE Then
        Debug.Print "Loaded " & m_Cache.count & " entries from cache sheet"
    #End If
End Sub

' =============================================================================
' LEGACY RECORDCACHE COMPATIBILITY
' =============================================================================
' These functions maintain compatibility with the original RecordCache code

''
' Initialize the record cache (legacy)
''
Public Sub Cache_RecordCache_Initialize()
    If m_RecordCache Is Nothing Then
        Set m_RecordCache = New Collection
    End If
    EnsureInitialized
End Sub

''
' Store records in cache (legacy)
' @param cacheKey - Key to store under
''
Public Sub Cache_RecordCache_Store(cacheKey As String)
    ' This function stores current response records
    ' In the new architecture, this is handled by Cache_StoreInCache
    ' Keep for compatibility
End Sub

''
' Retrieve records from cache (legacy)
' @param cacheKey - Key to look up
' @param found - Output: True if found
''
Public Sub Cache_RecordCache_Retreive(cacheKey As String, ByRef found As Boolean)
    Dim response As apiResponse
    found = Cache_TryGetFromCache(cacheKey, response)
End Sub

''
' Find position of key in cache (legacy)
' @param cacheKey - Key to find
' @return Long - Position or 0 if not found
''
Public Function Cache_RecordCache_Find(cacheKey As String) As Long
    If Cache_ExistsInCache(cacheKey) Then
        Cache_RecordCache_Find = 1
    Else
        Cache_RecordCache_Find = 0
    End If
End Function

''
' Reset cache (legacy)
''
Public Sub Cache_RecordCache_Reset()
    Cache_ClearCache
    Set m_RecordCache = New Collection
End Sub

''
' Display cache (legacy)
''
Public Sub Cache_RecordCache_Display()
    Cache_DisplayCache
End Sub

''
' Load cache (legacy)
''
Public Sub Cache_RecordCache_Load()
    Cache_LoadCacheFromSheet
End Sub

''
' Remove Bearer tokens from cache
' Clears any cached authentication tokens
''
Public Sub Cache_RecordCache_RemoveBearerTokens()
    Dim i As Long
    Dim pair As Variant
    Dim tokenParts As Variant
    
    Cache_RecordCache_Initialize
    
    For i = m_RecordCache.count To 1 Step -1
        pair = Split(m_RecordCache.item(i), "|")
        If UBound(pair) < 1 Then GoTo ContinueLoop
        
        tokenParts = Split(pair(1), "~")
        If UBound(tokenParts) >= 6 Then
            If Trim(tokenParts(0)) = "Bearer" Then
                m_RecordCache.Remove i
            End If
        End If
ContinueLoop:
    Next i
End Sub

''
' Dump cache to Cache sheet
''
Public Sub Cache_RecordCache_Dump()
    Dim cacheSheet As Worksheet
    Dim i As Long
    Dim pair As Variant
    Dim dataArr() As Variant
    Dim buttonNames, buttonActions, buttonPositions As Variant
    Dim exists As Boolean
    Dim shp As Shape
    Dim btn As Object
    Dim recordCount As Long
    
    ' Optimize Performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Initialize cache if needed
    Cache_RecordCache_Initialize

    ' Check if "Cache" sheet exists, else create it
    On Error Resume Next
    Set cacheSheet = ThisWorkbook.Sheets("Cache")
    If cacheSheet Is Nothing Then
        Set cacheSheet = ThisWorkbook.Sheets.Add
        cacheSheet.Name = "Cache"
        cacheSheet.Tab.Color = vbBlack
    End If
    On Error GoTo 0
    
    ' If "Cache" is active, clear and set headers
    If ActiveSheet.Name = "Cache" Then
        cacheSheet.Cells.Clear
        cacheSheet.Rows(1).RowHeight = 97

        ' Write headers
        cacheSheet.Cells(1, 1).value = "Key (URL)"
        cacheSheet.Cells(1, 2).value = "Value (Serialized JSON)"
        cacheSheet.Cells(1, 1).Font.Bold = True
        cacheSheet.Cells(1, 2).Font.Bold = True
        
        ' Set column widths
        cacheSheet.columns(1).ColumnWidth = 70
        cacheSheet.columns(2).ColumnWidth = 70
        
        ' Define button properties - use new Cache_ prefixed routines
        buttonNames = Array("Dump Cache", "Load", "Reset", "Close")
        buttonActions = Array("DoppioCache.Cache_RecordCache_Dump", "DoppioCache.Cache_RecordCache_Load", "DoppioCache.Cache_RecordCache_Reset", "DoppioCache.Cache_RecordCache_Close")
        buttonPositions = Array(28, 112, 197, 281)

        ' Loop through buttons and add if not exists
        For i = LBound(buttonNames) To UBound(buttonNames)
            exists = False
            
            ' Check if button exists
            For Each shp In cacheSheet.Shapes
                If shp.Type = msoFormControl Then
                    If LCase(shp.Name) = LCase(buttonNames(i)) Then
                        exists = True
                        Exit For
                    End If
                End If
            Next shp
            
            ' Add button only if it does not exist
            If Not exists Then
                Set btn = cacheSheet.Buttons.Add(buttonPositions(i), 26, 71, 30)
                btn.Caption = buttonNames(i)
                btn.OnAction = buttonActions(i)
            End If
        Next i
    End If

    ' Exit if recordCache is empty
    If m_RecordCache Is Nothing Then GoTo Cleanup
    recordCount = m_RecordCache.count
    If recordCount = 0 Then GoTo Cleanup
    
    ' Prepare data array for faster writing
    ReDim dataArr(1 To recordCount, 1 To 2)
    For i = 1 To recordCount
        pair = Split(m_RecordCache.item(i), "|")
        dataArr(i, 1) = pair(0) ' Cache Key
        If UBound(pair) >= 1 Then
            dataArr(i, 2) = pair(1) ' Cached Value (JSON)
        End If
    Next i

    ' Write data to sheet in bulk
    cacheSheet.Range("A2").Resize(recordCount, 2).value = dataArr

    ' Freeze Top Row if Active
    If ActiveSheet Is cacheSheet Then
        cacheSheet.Rows(2).Select
        ActiveWindow.FreezePanes = True
    End If

    ' Enable AutoFilter
    If Not cacheSheet.AutoFilterMode Then cacheSheet.Rows(1).AutoFilter

    ' Sort Data
    With cacheSheet.Sort
        .SortFields.Clear
        .SortFields.Add key:=cacheSheet.columns(1), Order:=xlAscending
        .SetRange cacheSheet.Range("A1:B" & recordCount + 1)
        .Header = xlYes
        .Apply
    End With

Cleanup:
    ' Restore Application Settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

''
' Close/hide the Cache sheet
''
Public Sub Cache_RecordCache_Close()
    On Error Resume Next
    Sheets("Cache").Visible = False
    On Error GoTo 0
End Sub

' =============================================================================
' LEGACY WRAPPERS - For backward compatibility with existing code
' =============================================================================

Public Sub RecordCache_Initialize()
    Cache_RecordCache_Initialize
End Sub

Public Sub RecordCache_Store(url As String)
    Cache_RecordCache_Store url
End Sub

Public Sub RecordCache_Retreive(url As String, ByRef found As Boolean)
    Cache_RecordCache_Retreive url, found
End Sub

Public Function RecordCache_Find(cacheKey As String) As Long
    RecordCache_Find = Cache_RecordCache_Find(cacheKey)
End Function

Public Sub RecordCache_Reset()
    Cache_RecordCache_Reset
End Sub

Public Sub RecordCache_Display()
    Cache_RecordCache_Display
End Sub

Public Sub RecordCache_Load()
    Cache_RecordCache_Load
End Sub

Public Sub RecordCache_RemoveBearerTokens()
    Cache_RecordCache_RemoveBearerTokens
End Sub

Public Sub RecordCache_Dump()
    Cache_RecordCache_Dump
End Sub

Public Sub RecordCache_Close()
    Cache_RecordCache_Close
End Sub


