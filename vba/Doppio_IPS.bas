Attribute VB_Name = "Doppio_IPS"
''
' Doppio Group - Doppio_IPS Module
' (c) Eric Pronovost - eric@doppiogroup.com
'
' IPS (Interactive Program Services) swagger discovery, layout loading,
' and transaction execution.
'
' Endpoints:
'   Transactions: GET /M3/ips/service/ionapi-doc/?pageSize=1000&search={api}
'   Layout:       GET /M3/ips/service/{api}/
'   Execute:      POST /M3/ips/service/{api}  (SOAP/XML)
'
Option Explicit

' =============================================================================
' CONSTANTS
' =============================================================================

Private Const IPS_SWAGGER_PATH As String = "/M3/ips/service/ionapi-doc/"
Private Const IPS_SERVICE_PATH As String = "/M3/ips/service/"

' =============================================================================
' IPS LAYOUT (Web Services)
' =============================================================================

Public Sub ProcessLayoutIPS(ws As Worksheet, mandatory As Boolean)
    Dim url As String
    
    On Error GoTo ErrorHandler
    
    ' Rename sheet
    If Left(ws.name, 5) = "Sheet" Then UI_RenameSheet ""
    
    m_b_Webservice = False
    url = ws.Cells(2, 1).value & "/"
    
    If m_s_LoadedUrl <> url Then
        Doppio_IPS.GetLayoutWS ws.Range("API").value, ws
        m_s_LoadedUrl = url
    End If

    
    AutoFit_ColumnsAndRows True, mandatory
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ProcessLayoutIPS: ERROR - " & Err.description
End Sub

' =============================================================================
' PUBLIC - TRANSACTIONS (Swagger Discovery)
' =============================================================================

''
' Load available IPS transactions for a given API into the Transactions sheet.
' Calls: GET /M3/ips/service/ionapi-doc/?pageSize=1000&search={api}
'
' @param api - API name (e.g. "PCS260WS")
' @param ws  - Active worksheet
'
Public Sub GetTransactionsIPS(api As String, ws As Worksheet)
    Dim response As apiResponse
    Dim cacheKey As String
    
    On Error GoTo ErrorHandler
    
    RecordCache_Initialize
    
    cacheKey = "IPS:ionapi-doc:" & api
    
    response = TryGetFromCache(cacheKey)
    
    If Not response.success Then
        response = ExecuteSwaggerGet(IPS_SWAGGER_PATH, "?pageSize=1000&search=" & api)
        
        If response.success Then
            Cache_StoreDataInCache cacheKey, response.data
        End If
    End If
    
    If response.success Then
        OutputSwaggerTransactions response
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Doppio_IPS.GetTransactionsIPS: ERROR - " & Err.description
End Sub


' =============================================================================
' PUBLIC - LAYOUT (Swagger Discovery)
' =============================================================================

''
' Load field layout for an IPS web service method.
' Calls: GET /M3/ips/service/{api}/
' Caches the ENTIRE response (all methods), then filters by transaction name.
'
' @param api - API name (e.g. "PCS260WS")
' @param ws  - Active worksheet (reads Transaction named range)
'
Public Sub GetLayoutWS(api As String, ws As Worksheet)
    Dim response As apiResponse
    Dim cacheKey As String
    Dim methodName As String
    Dim methodData As Object
    Dim inputItem As Object
    Dim outputItem As Object
    Dim aliasName As String
    Dim fieldName As String
    Dim dataType As String
    Dim fieldLength As Long
    Dim isRequired As Boolean
    Dim subPrograms As Object
    Dim subProgram As Object
    Dim subProgramName As String
    Dim subItem As Object

    On Error GoTo ErrorHandler

    cacheKey = "IPS:layout:" & api

    response = TryGetFromCache(cacheKey)

    If Not response.success Then
        response = ExecuteSwaggerGet(IPS_SERVICE_PATH & api & "/", "")

        If response.success Then
            Cache_StoreDataInCache cacheKey, response.data
        End If
    End If

    If Not response.success Then
        Debug.Print "GetLayoutWS: Failed to load layout for " & api
        Exit Sub
    End If

    ' Initialize column collections
    m_obj_ColumnNames.Initialize
    m_obj_ColumnDescriptions.Initialize
    m_obj_ColumnTypes.Initialize
    m_obj_ColumnConditions.Initialize
    m_obj_ColumnDirections.Initialize

    methodName = ws.Range("Transaction").value

    If response.results Is Nothing Then Exit Sub

    For Each methodData In response.results.item("methods")
        If methodData.item("name") = methodName Then

            ' --- Main program: visible input fields ---
            For Each inputItem In methodData.item("program").item("visibleInput")
                aliasName = inputItem.item("alias")
                fieldName = inputItem.item("name")
                dataType = ConvertDataType(inputItem.item("datatype"))
                fieldLength = inputItem.item("fieldLength")
                isRequired = inputItem.item("required")

                m_obj_ColumnNames.Add aliasName
                m_obj_ColumnDescriptions.Add fieldName & vbCrLf & dataType & fieldLength
                m_obj_ColumnTypes.Add dataType
                m_obj_ColumnConditions.Add isRequired
                m_obj_ColumnDirections.Add "I"
            Next inputItem

            ' --- Main program: output fields ---
            For Each outputItem In methodData.item("program").item("outputs")
                aliasName = outputItem.item("alias")
                fieldName = outputItem.item("name")
                dataType = ConvertDataType(outputItem.item("datatype"))
                fieldLength = outputItem.item("fieldLength")

                m_obj_ColumnNames.Add aliasName
                m_obj_ColumnDescriptions.Add fieldName & vbCrLf & dataType & fieldLength
                m_obj_ColumnTypes.Add dataType
                m_obj_ColumnConditions.Add False
                m_obj_ColumnDirections.Add "O"
            Next outputItem

            ' --- Sub-programs (programs[] nested under the main program) ---
            ' Fields are prefixed "SubProgramName:Alias" so the SOAP body builder
            ' routes them into the correct nested XML element.
            On Error Resume Next
            Set subPrograms = methodData.item("program").item("programs")
            On Error GoTo ErrorHandler

            If Not subPrograms Is Nothing Then
                For Each subProgram In subPrograms
                    subProgramName = subProgram.item("name")

                    ' Sub-program visible input fields
                    On Error Resume Next
                    Set subItem = Nothing
                    For Each subItem In subProgram.item("visibleInput")
                        On Error GoTo ErrorHandler
                        aliasName = subProgramName & ":" & subItem.item("alias")
                        fieldName = subItem.item("name")
                        dataType = ConvertDataType(subItem.item("datatype"))
                        fieldLength = subItem.item("fieldLength")
                        isRequired = subItem.item("required")

                        m_obj_ColumnNames.Add aliasName
                        m_obj_ColumnDescriptions.Add fieldName & vbCrLf & dataType & fieldLength
                        m_obj_ColumnTypes.Add dataType
                        m_obj_ColumnConditions.Add isRequired
                        m_obj_ColumnDirections.Add "I"
                        On Error Resume Next
                    Next subItem
                    On Error GoTo ErrorHandler

                Next subProgram
            End If

            Exit For
        End If
    Next methodData

    Exit Sub

ErrorHandler:
    Debug.Print "Doppio_IPS.GetLayoutWS: ERROR - " & Err.description
End Sub


' =============================================================================
' PUBLIC - PROCESS IPS TRANSACTIONS
' =============================================================================

''
' Process IPS (Web Service) transactions row by row.
' Each row generates a SOAP envelope and is sent as an individual POST.
' IPS does not support batching like MI does.
'
' @param ws         - Active worksheet
' @param startTime  - Timer value for elapsed time display
' @param currentRow - ByRef row counter (shared with Doppio_Process)
'
Public Sub ProcessIPSTransactions(ws As Worksheet, startTime As Single, ByRef currentRow As Long)
    Dim strServiceName As String
    Dim strMethod As String
    Dim strProgram As String
    Dim lastColumn As Integer
    Dim inputColumns() As Integer
    Dim inputFields() As String
    Dim inputValues() As String
    Dim soapBody As String
    Dim i As Long, j As Long
    Dim settings As ApiSettings
    Dim httpResp As httpResponse
    Dim cancelled As Boolean
    
    On Error GoTo ErrorHandler
    
    Application.EnableCancelKey = xlErrorHandler
    cancelled = False
    
    settings = Config_ApiSettings
    
    strServiceName = ws.Range("API").value
    strMethod = ws.Range("Transaction").value
    
    If strServiceName = "" Or strMethod = "" Then
        MsgBox "Please enter API and Transaction.", vbExclamation
        Exit Sub
    End If
    
    ' Get actual program name from cached layout (e.g. "PPS040")
    strProgram = GetProgramName(strServiceName, strMethod)
    
    lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    ' Initialize field arrays
    ReDim inputColumns(1 To 400)
    ReDim inputFields(1 To 400)
    ReDim inputValues(1 To 400)
    
    j = 0
    For i = 1 To lastColumn
        If ws.Cells(currentRow - 1, i + 1) <> "" Then
            j = j + 1
            If j > 400 Then Exit For
            inputColumns(j) = i + 1
            inputFields(j) = ws.Cells(currentRow - 1, i + 1)
        End If
    Next i
    
    ws.Calculate
    
    If IsEmpty(ws.Range("B9").value) Then
        ws.Range("B9").value = "?"
    End If
    
    ' Process each row (one SOAP call per row - IPS does not support batching)
    While ws.Cells(currentRow, 2).value <> ""
        DoEvents  ' Allow ESC detection
        
        ' Collect values from current row
        For i = 1 To lastColumn
            If inputColumns(i) <> 0 Then
                Dim cellValue As Variant
                cellValue = ws.Cells(currentRow, inputColumns(i))
                
                If isError(cellValue) Then cellValue = ""
                If cellValue = "?" And currentRow = 9 And i = 1 Then cellValue = ""
                
                inputValues(i) = CStr(cellValue)
            End If
        Next i
        
        ' Build SOAP envelope
        soapBody = BuildSoapEnvelope(strServiceName, strMethod, strProgram, _
                                      inputColumns, inputFields, inputValues, _
                                      m_s_Company, m_s_Division)
        Debug.Print "Doppio_IPS.ProcessIPSTransactions: " & soapBody
            
        UI_ShowPleaseWait "Please Wait... Calling IPS (row " & _
            currentRow - 8 & ") - Press ESC to cancel"
        
        ' Execute SOAP POST
        httpResp = ExecuteIpsPost(strServiceName, soapBody, settings.MaxTimeout)
        
        ' Parse XML response and write status
        ParseSoapResponse httpResp, ws, currentRow, strServiceName, strMethod
        
        currentRow = currentRow + 1
    Wend
    
Cleanup:
    Application.EnableCancelKey = xlInterrupt
    If cancelled Then
        UI_KillPleaseWait
        ws.Range("A" & currentRow).value = "CANCELLED"
        ws.Range("A" & currentRow).Font.Color = COLOR_ERROR
        MsgBox "Process cancelled by user at row " & currentRow & "." & vbCrLf & _
               "Rows processed before cancellation are valid.", vbExclamation, "Cancelled"
    End If
    Exit Sub
    
ErrorHandler:
    If Err.Number = 18 Then
        cancelled = True
        Resume Cleanup
    End If
    
    Application.EnableCancelKey = xlInterrupt
    UI_KillPleaseWait
    MsgBox "Error processing IPS transactions: " & Err.description, vbCritical
End Sub


' =============================================================================
' PRIVATE - SWAGGER HTTP (JSON discovery endpoints)
' =============================================================================

''
' Execute a GET against an IPS/Swagger JSON endpoint.
' Swagger discovery endpoints return JSON (not XML like IPS execution).
'
Private Function ExecuteSwaggerGet(path As String, queryString As String) As apiResponse
    Dim config As httpConfig
    Dim httpResp As httpResponse
    Dim response As apiResponse
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    config.url = m_s_MainUrl & path & queryString
    config.method = HttpMethod_GET
    config.contentType = "application/json; charset=UTF-8"
    config.AcceptType = "application/json; charset=UTF-8"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = 30
    config.body = ""
    
    Debug.Print "Doppio_IPS.ExecuteSwaggerGet: " & config.url
    
    httpResp = ExecuteRequest(config)
    
    response.success = httpResp.success
    response.data = httpResp.body
    response.errorMessage = httpResp.errorMessage
    
    If httpResp.success And Len(httpResp.body) > 0 Then
        Set json = ParseJson(httpResp.body)
        If Not json Is Nothing Then
            Set response.results = json
        End If
    End If
    
    ExecuteSwaggerGet = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.errorMessage = Err.description
    ExecuteSwaggerGet = response
End Function


' =============================================================================
' PRIVATE - SOAP HTTP (XML execution endpoints)
' =============================================================================

''
' Execute a SOAP POST to IPS.
' URL: /M3/ips/service/{API}
'
Private Function ExecuteIpsPost(apiName As String, soapBody As String, _
                                 timeoutSeconds As Integer) As httpResponse
    Dim config As httpConfig
    
    config.url = m_s_MainUrl & IPS_SERVICE_PATH & apiName
    config.method = HttpMethod_POST
    config.contentType = "application/xml"
    config.AcceptType = "application/xml; charset=UTF-8"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = timeoutSeconds
    config.body = soapBody
    
    Debug.Print "Doppio_IPS.ExecuteIpsPost: " & config.url
    
    ExecuteIpsPost = ExecuteRequest(config)
End Function


' =============================================================================
' PRIVATE - SOAP ENVELOPE BUILDER
' =============================================================================

''
' Build a complete SOAP envelope for an IPS call.
'
' Structure:
'   <?xml version="1.0" encoding="UTF-8"?>
'   <SOAP-ENV:Envelope xmlns:...>
'     <SOAP-ENV:Header>
'       <cred:lws><cred:company>780</cred:company><cred:division>AAA</cred:division></cred:lws>
'     </SOAP-ENV:Header>
'     <SOAP-ENV:Body>
'       <chg:Add><chg:PCS260><chg:FACI>100</chg:FACI>...</chg:PCS260></chg:Add>
'     </SOAP-ENV:Body>
'   </SOAP-ENV:Envelope>
'
Private Function BuildSoapEnvelope(webService As String, method As String, _
                                    program As String, _
                                    inputColumns() As Integer, _
                                    inputFields() As String, _
                                    inputValues() As String, _
                                    cono As String, divi As String) As String
    Dim xml As String
    Dim nsUri As String
    
    ' Namespace URI depends on multi-tenant vs single-tenant
    If m_b_Multitenant Then
        nsUri = "http://schemas.infor.com/ips/" & webService & "/" & method
    Else
        nsUri = "http://your.company.net/" & webService & "/" & method
    End If

    ' Namespace prefix derived from method name (e.g. "Add" -> "add")
    Dim nsPrefix As String
    nsPrefix = LCase(method)

    ' XML declaration + Envelope open
    xml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>" & _
          "<SOAP-ENV:Envelope" & _
          " xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""" & _
          " xmlns:" & nsPrefix & "=""" & nsUri & """" & _
          " xmlns:cred=""http://lawson.com/ws/credentials"">"

    ' Header (company/division)
    xml = xml & "<SOAP-ENV:Header>" & _
                "<cred:lws>" & _
                "<cred:company>" & cono & "</cred:company>" & _
                "<cred:division>" & divi & "</cred:division>" & _
                "</cred:lws>" & _
                "</SOAP-ENV:Header>"

    ' Body
    xml = xml & "<SOAP-ENV:Body>"
    xml = xml & BuildSoapBody(method, program, nsPrefix, inputColumns, inputFields, inputValues)
    xml = xml & "</SOAP-ENV:Body>"
    xml = xml & "</SOAP-ENV:Envelope>"
    
    BuildSoapEnvelope = xml
End Function


''
' Build the SOAP body content with field values.
' Supports nested related programs via "ProgramName:FieldName" syntax in headers.
'
Private Function BuildSoapBody(method As String, program As String, nsPrefix As String, _
                                inputColumns() As Integer, _
                                inputFields() As String, _
                                inputValues() As String) As String
    Dim xml As String
    Dim i As Long
    Dim inputField As String
    Dim currentLevel As String
    Dim level2 As String, level3 As String, level4 As String
    Dim p As String
    p = nsPrefix & ":"

    xml = "<" & p & method & ">"

    For i = 1 To UBound(inputValues)
        If inputValues(i) <> "" And inputColumns(i) <> 0 Then
            inputField = inputFields(i)

            ' Open primary program level on first field
            If currentLevel = "" Then
                currentLevel = program
                xml = xml & "<" & p & program & ">"
            End If

            ' Handle related programs (field format: "RelatedProgram:FieldName")
            If InStr(inputField, ":") > 0 Then
                Dim relProgram As String
                relProgram = Left(inputField, InStr(inputField, ":") - 1)

                If relProgram <> currentLevel Then
                    currentLevel = relProgram
                    If level2 = "" Then
                        level2 = currentLevel
                    ElseIf level3 = "" Then
                        level3 = currentLevel
                    ElseIf level4 = "" Then
                        level4 = currentLevel
                    End If
                    xml = xml & "<" & p & currentLevel & ">"
                End If

                inputField = Mid(inputField, InStr(inputField, ":") + 1)
            End If

            xml = xml & "<" & p & inputField & ">" & EscapeXml(inputValues(i)) & "</" & p & inputField & ">"
        End If
    Next i

    ' Close nested levels
    If level4 <> "" Then xml = xml & "</" & p & level4 & ">"
    If level3 <> "" Then xml = xml & "</" & p & level3 & ">"
    If level2 <> "" Then xml = xml & "</" & p & level2 & ">"
    xml = xml & "</" & p & program & ">"
    xml = xml & "</" & p & method & ">"

    BuildSoapBody = xml
End Function


' =============================================================================
' PRIVATE - RESPONSE PARSING
' =============================================================================

''
' Parse SOAP XML response and write status + output values to the worksheet.
'
Private Sub ParseSoapResponse(httpResp As httpResponse, ws As Worksheet, _
                               rowNum As Long, apiName As String, method As String)
    Dim xmlDoc As Object
    Dim faultNode As Object
    Dim resultNode As Object
    Dim childNode As Object
    Dim ns As String
    Dim nodeName As String
    Dim key As String
    Dim parts() As String
    Dim columnIndex As Long
    Dim i As Long
    Dim lastCol As Long
    
    On Error Resume Next
    
    ' Check HTTP-level failure
    If Not httpResp.success Or Len(httpResp.body) = 0 Then
        ws.Cells(rowNum, 1).value = "NOK " & httpResp.errorMessage
        ws.Cells(rowNum, 1).Font.Color = COLOR_ERROR
        Exit Sub
    End If
    
    ' Parse XML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.LoadXML httpResp.body
    
    If xmlDoc.parseError <> 0 Then
        ws.Cells(rowNum, 1).value = "NOK XML parse error"
        ws.Cells(rowNum, 1).Font.Color = COLOR_ERROR
        Exit Sub
    End If
    
    ' Check for SOAP fault
    Set faultNode = xmlDoc.SelectSingleNode("//faultstring")
    If Not faultNode Is Nothing Then
        ws.Cells(rowNum, 1).value = "NOK " & faultNode.text
        ws.Cells(rowNum, 1).Font.Color = COLOR_ERROR
        Exit Sub
    End If
    
    ' Extract result node � use actual program name from layout (e.g. "PPS040")
    ' Namespace prefix mirrors BuildSoapEnvelope: LCase(method)
    nodeName = GetProgramName(apiName, method)
    Dim nsPrefix As String
    nsPrefix = LCase(method)

    If m_b_Multitenant Then
        ns = "xmlns:" & nsPrefix & "='http://schemas.infor.com/ips/" & apiName & "/" & method & "'"
    Else
        ns = "xmlns:" & nsPrefix & "='http://your.company.net/" & apiName & "/" & method & "'"
    End If

    xmlDoc.SetProperty "SelectionNamespaces", ns
    xmlDoc.SetProperty "SelectionLanguage", "XPath"
    Set resultNode = xmlDoc.SelectSingleNode("//" & nsPrefix & ":" & nodeName)
    
    ' Default to OK
    ws.Cells(rowNum, 1).value = "OK"
    ws.Cells(rowNum, 1).Font.Color = COLOR_SUCCESS

    ' Extract output field values into matching columns
    If Not resultNode Is Nothing Then
        lastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column

        ' Pre-clear all output columns so fields absent from the response are blanked
        ' Output columns are identified by their header cell color (RGB 128,128,128)
        Dim clearCol As Long
        For clearCol = 2 To lastCol
            If ws.Cells(8, clearCol).Interior.Color = RGB(128, 128, 128) Then
                ws.Cells(rowNum, clearCol).value = ""
            End If
        Next clearCol

        For Each childNode In resultNode.ChildNodes
            key = childNode.nodeName

            ' Strip namespace prefix (e.g. "chg:ITNO" -> "ITNO")
            parts = Split(key, ":")
            If UBound(parts) > 0 Then key = parts(UBound(parts))

            ' Check if this child is a sub-program container (has element children,
            ' not a text value). If so, iterate its children using "Program:Field" keys.
            Dim firstChild As Object
            Set firstChild = Nothing
            On Error Resume Next
            Set firstChild = childNode.ChildNodes(0)
            On Error GoTo 0

            If Not firstChild Is Nothing Then
            If childNode.text <> firstChild.text Then
                ' Sub-program container � walk its children with prefixed key
                Dim subChildNode As Object
                Dim subKey As String
                Dim subParts() As String
                For Each subChildNode In childNode.ChildNodes
                    subKey = subChildNode.nodeName
                    subParts = Split(subKey, ":")
                    If UBound(subParts) > 0 Then subKey = subParts(UBound(subParts))

                    columnIndex = 0
                    For i = 2 To lastCol
                        If ws.Cells(8, i).value = key & ":" & subKey Then
                            columnIndex = i
                            Exit For
                        End If
                    Next i

                    If columnIndex > 0 Then
                        ws.Cells(rowNum, columnIndex).value = subChildNode.text
                    End If
                Next subChildNode
            Else
                ' Regular leaf field � match directly against column header
                columnIndex = 0
                For i = 2 To lastCol
                    If ws.Cells(8, i).value = key Then
                        columnIndex = i
                        Exit For
                    End If
                Next i

                If columnIndex > 0 Then
                    ws.Cells(rowNum, columnIndex).value = childNode.text
                End If
            End If
            End If
        Next childNode
    End If
    
    On Error GoTo 0
End Sub


' =============================================================================
' PRIVATE - CACHE
' =============================================================================

Private Function TryGetFromCache(cacheKey As String) As apiResponse
    Dim response As apiResponse
    Dim found As Boolean
    
    found = Cache_TryGetFromCache(cacheKey, response)
    
    If Not found Then
        response.success = False
    End If
    
    TryGetFromCache = response
End Function


' =============================================================================
' PRIVATE - OUTPUT
' =============================================================================

''
' Write swagger collection to Transactions sheet.
' Parses entity format "PCS260WS#Add" into Col A (API) and Col B (method).
'
Private Sub OutputSwaggerTransactions(response As apiResponse)
    Dim ws As Worksheet
    Dim swaggerArray As Object
    Dim swagger As Object
    Dim entity As String
    Dim entityParts() As String
    Dim rowNum As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Transactions")
    ws.UsedRange.Clear
    
    If response.results Is Nothing Then Exit Sub
    
    Set swaggerArray = response.results.item("swaggerCollection").item("swagger")
    
    rowNum = 3
    For Each swagger In swaggerArray
        entity = swagger.item("entity")
        entityParts = Split(entity, "#")
        
        ws.Cells(rowNum, 1).value = entityParts(0)
        If UBound(entityParts) >= 1 Then
            ws.Cells(rowNum, 2).value = entityParts(1)
        End If
        ws.Cells(rowNum, 3).value = "S"
        rowNum = rowNum + 1
    Next swagger
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Doppio_IPS.OutputSwaggerTransactions: ERROR - " & Err.description
End Sub


' =============================================================================
' PRIVATE - UTILITIES
' =============================================================================

''
' Convert IPS data types to short codes for column display
'
Private Function ConvertDataType(ipsType As String) As String
    Select Case ipsType
        Case "STRING":  ConvertDataType = "A"
        Case "DECIMAL": ConvertDataType = "N"
        Case "INTEGER": ConvertDataType = "N"
        Case "DATE":    ConvertDataType = "D"
        Case Else:      ConvertDataType = ipsType
    End Select
End Function


''
' Get the actual program name from the cached IPS layout JSON.
' e.g. for MPD_PPS044_Add / Add -> returns "PPS040"
' Falls back to DeriveProgram if layout is not cached or method not found.
'
Private Function GetProgramName(api As String, methodName As String) As String
    Dim response As apiResponse
    Dim cacheKey As String
    Dim methodData As Object

    On Error GoTo Fallback

    cacheKey = "IPS:layout:" & api
    response = TryGetFromCache(cacheKey)

    If response.success And Not response.results Is Nothing Then
        For Each methodData In response.results.item("methods")
            If methodData.item("name") = methodName Then
                GetProgramName = methodData.item("program").item("name")
                Exit Function
            End If
        Next methodData
    End If

Fallback:
    GetProgramName = DeriveProgram(api)
End Function


''
' Derive program name from API name.
' "PCS260WS" -> "PCS260" (if 3rd char is "S", take 6 chars, else 5)
'
Private Function DeriveProgram(apiName As String) As String
    If Len(apiName) >= 3 Then
        If Mid(apiName, 3, 1) = "S" Then
            DeriveProgram = Left(apiName, 6)
        Else
            DeriveProgram = Left(apiName, 5)
        End If
    Else
        DeriveProgram = apiName
    End If
End Function


''
' Escape special XML characters in values
'
Private Function EscapeXml(value As String) As String
    Dim result As String
    result = value
    result = Replace(result, "&", "&amp;")
    result = Replace(result, "<", "&lt;")
    result = Replace(result, ">", "&gt;")
    result = Replace(result, """", "&quot;")
    result = Replace(result, "'", "&apos;")
    EscapeXml = result
End Function


''
' Build parallel M3-code / IPS-alias arrays from the layout that was most
' recently loaded by GetLayoutWS.
'
' Each entry in m_obj_ColumnDescriptions has the form "W1ITNO<vbCrLf>A15".
' The M3 field code is the 4 characters of that name starting at position 3
' (e.g. "W1ITNO" -> "ITNO", "WPLEA1" -> "LEA1").
'
' For sub-program fields GetLayoutWS already prefixes the alias with the
' program name ("PPS044:SupplyLeadTime"), so the returned ipsAliases array
' contains ready-to-use row-8 values including any colon notation.
'
' Returns the number of entries placed in the arrays (0 = no layout loaded).
'
Public Function BuildM3ToAliasMap(m3Codes() As String, ipsAliases() As String) As Long
    Dim k As Long
    Dim total As Long
    Dim desc As String
    Dim ipsFieldName As String

    total = m_obj_ColumnNames.count()

    If total = 0 Then
        BuildM3ToAliasMap = 0
        Exit Function
    End If

    ReDim m3Codes(1 To total)
    ReDim ipsAliases(1 To total)

    For k = 1 To total
        desc = m_obj_ColumnDescriptions.item(k)

        ' Strip the type/length suffix after the line-break
        If InStr(desc, vbCrLf) > 0 Then
            ipsFieldName = Left(desc, InStr(desc, vbCrLf) - 1)
        Else
            ipsFieldName = desc
        End If

        ' Extract the 4-char M3 code (chars 3-6, e.g. "W1ITNO" -> "ITNO")
        If Len(ipsFieldName) >= 6 Then
            m3Codes(k) = Mid(ipsFieldName, 3, 4)
        Else
            m3Codes(k) = ""
        End If

        ipsAliases(k) = m_obj_ColumnNames.item(k)
    Next k

    BuildM3ToAliasMap = total
End Function


