Attribute VB_Name = "Doppio_PostMan"
Sub Postman_Build()
    Dim ws As Worksheet
    Dim tenant As TenantConfig
    Dim settings As ApiSettings
    Dim json As Object
    Dim program As String
    Dim transaction As String
    Dim requestUrl As String
    Dim requestBody As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    
    ' Load configuration
    Tenant_Information
    
    ' Settings come from Doppio_Config (loaded from Settings sheet at workbook open)
    settings = Config_ApiSettings
    
    ' Tenant info comes from Doppio globals (loaded by Tenant_Information)
    ' Build a TenantConfig from them
    tenant.tenantId = ti
    tenant.clientId = ci
    tenant.clientSecret = cs
    tenant.IonUrl = iu
    tenant.SsoUrl = pu
    tenant.tokenEndpoint = ot
    tenant.ServiceAccountKey = saak
    tenant.ServiceAccountSecret = sask
    
    ' Read program and transaction from the worksheet
    program = ws.Range("API").value
    transaction = ws.Range("Transaction").value
    
    ' Build the request body from worksheet fields (row 8 = names, row 9 = values)
    requestBody = Postman_BuildBody(ws, program, transaction)
    
    ' Build the full API URL from settings
    requestUrl = Postman_BuildUrl(m_s_MainUrl, m_s_MiPath, settings, m_s_M3user, m_s_Company, m_s_Division)
    
    ' Build Postman collection
    Set json = ParseJson("{}")
    json.Add "info", Postman_BuildInfo(tenant.tenantId)
    
    Dim itemArray As Object
    Set itemArray = ParseJson("[]")
    itemArray.Add Postman_BuildRequestItem(program & " " & transaction, requestUrl, requestBody)
    json.Add "item", itemArray
    
    json.Add "auth", Postman_BuildOAuth2(tenant)
    json.Add "event", ParseJson("[]")
    
    ' Show result
    SampleRESTPopup ConvertToJson(json), "Collection Import"
    
    Exit Sub

ErrorHandler:
    Debug.Print "Postman_Build: ERROR - " & Err.description
    MsgBox "Error building Postman collection: " & Err.description, vbExclamation
End Sub


' =============================================================================
' Body & URL Builders
' =============================================================================

Private Function Postman_BuildBody(ws As Worksheet, program As String, transaction As String) As String
    Dim lastCol As Long
    Dim i As Long
    Dim fieldName As String
    Dim fieldValue As String
    Dim recordPairs() As String
    Dim selectedCols() As String
    Dim pairCount As Long
    Dim colCount As Long
    
    lastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    pairCount = 0
    colCount = 0
    
    ' Collect field names and non-blank values from worksheet rows 8/9
    For i = 2 To lastCol
        fieldName = Trim(ws.Cells(8, i).value)
        If fieldName = "" Then GoTo NextField
        
        ' Add to selectedColumns
        ReDim Preserve selectedCols(colCount)
        selectedCols(colCount) = """" & fieldName & """"
        colCount = colCount + 1
        
        ' Add non-blank field-value pairs to record
        fieldValue = Trim(ws.Cells(9, i).value)
        If fieldValue <> "" Then
            fieldValue = Replace(fieldValue, "\", "\\")
            fieldValue = Replace(fieldValue, """", "\""")
            ReDim Preserve recordPairs(pairCount)
            recordPairs(pairCount) = """" & fieldName & """: """ & fieldValue & """"
            pairCount = pairCount + 1
        End If
NextField:
    Next i
    
    ' Build JSON body string
    Dim jsonRecord As String
    Dim jsonSelected As String
    
    If pairCount > 0 Then
        jsonRecord = "{" & Join(recordPairs, ", ") & "}"
    Else
        jsonRecord = "{}"
    End If
    
    If colCount > 0 Then
        jsonSelected = "[" & Join(selectedCols, ", ") & "]"
    Else
        jsonSelected = "[]"
    End If
    
    Postman_BuildBody = "{""program"":""" & program & """," & _
                        """transactions"":[{" & _
                        """transaction"":""" & transaction & """," & _
                        """record"":" & jsonRecord & "," & _
                        """selectedColumns"":" & jsonSelected & _
                        "}]}"
End Function


Private Function Postman_BuildUrl(mainUrl As String, miPath As String, _
                                   settings As ApiSettings, _
                                   m3user As String, company As String, division As String) As String
    Dim url As String
    url = mainUrl & miPath & "?maxrecs=" & settings.MaxRecords & "&extendedresult=true"
    
    If m3user <> "" Then url = url & "&m3user=" & m3user
    
    If settings.righttrim Then
        url = url & "&righttrim=true"
    Else
        url = url & "&righttrim=false"
    End If
    
    If company <> "" Then url = url & "&cono=" & company
    If division <> "" Then url = url & "&divi=" & division
    
    Postman_BuildUrl = url
End Function


' =============================================================================
' Postman JSON Structure Builders
' =============================================================================

Private Function Postman_BuildInfo(tenantId As String) As Object
    Dim info As Object
    Set info = ParseJson("{}")
    info.Add "_postman_id", "4cc39a68-a241-4334-ba6a-0abff2c7e49b"
    info.Add "name", tenantId
    info.Add "schema", "https://schema.getpostman.com/json/collection/v2.1.0/collection.json"
    info.Add "_exporter_id", "4894281"
    Set Postman_BuildInfo = info
End Function


Private Function Postman_BuildRequestItem(itemName As String, url As String, body As String) As Object
    Dim item As Object
    Set item = ParseJson("{}")
    item.Add "name", itemName
    item.Add "request", Postman_BuildRequest(url, body)
    item.Add "response", ParseJson("[]")
    Set Postman_BuildRequestItem = item
End Function


Private Function Postman_BuildRequest(url As String, body As String) As Object
    Dim request As Object
    Set request = ParseJson("{}")
    
    request.Add "method", "POST"
    
    ' Header
    Dim headerArray As Object, headerObj As Object
    Set headerArray = ParseJson("[]")
    Set headerObj = ParseJson("{}")
    headerObj.Add "key", "Content-Type"
    headerObj.Add "value", "application/json"
    headerArray.Add headerObj
    request.Add "header", headerArray
    
    ' Body
    Dim bodyObj As Object, optionsObj As Object, rawObj As Object
    Set bodyObj = ParseJson("{}")
    bodyObj.Add "mode", "raw"
    bodyObj.Add "raw", body
    Set optionsObj = ParseJson("{}")
    Set rawObj = ParseJson("{}")
    rawObj.Add "language", "json"
    optionsObj.Add "raw", rawObj
    bodyObj.Add "options", optionsObj
    request.Add "body", bodyObj
    
    ' URL
    request.Add "url", Postman_ParseURL(url)
    
    ' Description
    request.Add "description", "this is the description"
    
    Set Postman_BuildRequest = request
End Function


Private Function Postman_BuildOAuth2(tenant As TenantConfig) As Object
    Dim auth As Object
    Set auth = ParseJson("{}")
    auth.Add "type", "oauth2"
    
    Dim oauth2 As Object
    Set oauth2 = ParseJson("[]")
    oauth2.Add Postman_CreateOAuthField("password", tenant.ServiceAccountSecret)
    oauth2.Add Postman_CreateOAuthField("username", tenant.ServiceAccountKey)
    oauth2.Add Postman_CreateOAuthField("accessTokenUrl", tenant.SsoUrl & tenant.tokenEndpoint)
    oauth2.Add Postman_CreateOAuthField("grant_type", "password_credentials")
    oauth2.Add Postman_CreateOAuthField("clientSecret", tenant.clientSecret)
    oauth2.Add Postman_CreateOAuthField("clientId", tenant.clientId)
    oauth2.Add Postman_CreateOAuthField("tokenName", tenant.tenantId)
    oauth2.Add Postman_CreateOAuthField("addTokenTo", "header")
    auth.Add "oauth2", oauth2
    
    Set Postman_BuildOAuth2 = auth
End Function


Private Function Postman_CreateOAuthField(key As String, value As String) As Object
    Dim field As Object
    Set field = ParseJson("{}")
    field.Add "key", key
    field.Add "value", value
    field.Add "type", "string"
    Set Postman_CreateOAuthField = field
End Function


' =============================================================================
' URL Parser (returns the url object directly, not wrapped)
' =============================================================================

Private Function Postman_ParseURL(fullUrl As String) As Object
    Dim protocol As String, host As String, path As String
    Dim queryString As String, remainingUrl As String, pathAndQuery As String
    Dim hostParts() As String, pathParts() As String
    Dim pathStart As Long, queryStart As Long
    Dim i As Long
    
    ' Split URL into components
    protocol = Left(fullUrl, InStr(fullUrl, "://") - 1)
    remainingUrl = Mid(fullUrl, InStr(fullUrl, "://") + 3)
    pathStart = InStr(remainingUrl, "/")
    host = Left(remainingUrl, pathStart - 1)
    hostParts = Split(host, ".")
    
    pathAndQuery = Mid(remainingUrl, pathStart + 1)
    queryStart = InStr(pathAndQuery, "?")
    If queryStart > 0 Then
        path = Left(pathAndQuery, queryStart - 1)
        queryString = Mid(pathAndQuery, queryStart + 1)
    Else
        path = pathAndQuery
        queryString = ""
    End If
    pathParts = Split(path, "/")
    
    ' Build URL object
    Dim urlObj As Object
    Set urlObj = ParseJson("{}")
    urlObj.Add "raw", fullUrl
    urlObj.Add "protocol", protocol
    
    Dim hostJson As Object
    Set hostJson = ParseJson("[]")
    For i = LBound(hostParts) To UBound(hostParts)
        hostJson.Add hostParts(i)
    Next i
    urlObj.Add "host", hostJson
    
    Dim pathJson As Object
    Set pathJson = ParseJson("[]")
    For i = LBound(pathParts) To UBound(pathParts)
        pathJson.Add pathParts(i)
    Next i
    urlObj.Add "path", pathJson
    
    Dim queryJson As Object
    Set queryJson = ParseJson("[]")
    If queryString <> "" Then
        Dim params() As String
        params = Split(queryString, "&")
        For i = LBound(params) To UBound(params)
            If InStr(params(i), "=") > 0 Then
                Dim param As Object
                Set param = ParseJson("{}")
                param.Add "key", Split(params(i), "=")(0)
                param.Add "value", Split(params(i), "=")(1)
                queryJson.Add param
            End If
        Next i
    End If
    urlObj.Add "query", queryJson
    
    Set Postman_ParseURL = urlObj
End Function

