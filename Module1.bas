Attribute VB_Name = "Module1"

Sub ToUpperCase()
Attribute ToUpperCase.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ToUpperCase Macro
'

'
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=UPPER(RC[-1])"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C28"), Type:=xlFillDefault
    Range("C2:C28").Select
    Columns("B:B").Select
    Columns("C:C").EntireColumn.AutoFit
    Range("B2:C28").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("H26").Select
    ActiveWorkbook.Save
    Range("E8").Select
    Sheets("Sheet1").Select
End Sub
Sub Format_Requirement()
Attribute Format_Requirement.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Format_Requirement Macro
'

'
    Range("A1").Select
    Selection.Cut
    Range("A1").Select
    ActiveWindow.DisplayGridlines = False
    Range("A1").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("E1").Select
    ActiveSheet.Paste
    Columns("E:E").ColumnWidth = 107.33
    Rows("1:1").RowHeight = 408
    Rows("1:1").EntireRow.AutoFit
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "REQUIREMENT"
    Range("E1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    Range("E2").Select
    ActiveWorkbook.Save
End Sub
Sub MakeGeminiCall()
    
End Sub

' Add a reference to Microsoft XML, v6.0 (or later) in VBA editor:
'   Tools > References > Microsoft XML, v6.0

Option Explicit

' Make sure to have VBA-JSON installed (JsonConverter.bas module imported)

' ---- Configuration (Replace with your values) ----
Const PROJECT_ID As String = "your-gcp-project-id" ' Your Google Cloud project ID
Const LOCATION As String = "us-central1"       ' The location of your Vertex AI endpoint
Const MODEL_ID As String = "text-bison@001"    ' The model you want to use (e.g., text-bison@001)  Check your model ID.
'---------------------------------------------

' API endpoint for text generation
Const API_ENDPOINT As String = "https://us-central1-aiplatform.googleapis.com/v1/projects/" & PROJECT_ID & "/locations/" & LOCATION & "/publishers/google/models/" & MODEL_ID & ":predict"


' Function to generate a JWT token (ONLY FOR SERVICE ACCOUNT AUTH - DO NOT USE IF USING ADC)
' This function is for example purposes only and needs proper implementation
' Function GetJwtToken(ByVal serviceAccountKeyPath As String) As String
'   ' ... JWT token generation logic using the service account key ...
' End Function


' Function to get an access token using ADC
Function GetAccessToken() As String
    Dim objHTTP As Object
    Dim strURL As String
    Dim strResponse As String
    Dim objJson As Object

    Set objHTTP = CreateObject("MSXML2.XMLHTTP60") ' Or XMLHTTP, depending on your version

    strURL = "http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token"

    With objHTTP
        .Open "GET", strURL, False ' Synchronous request
        .setRequestHeader "Metadata-Flavor", "Google"
        .Send

        strResponse = .responseText

        If .Status = 200 Then
            Set objJson = JsonConverter.ParseJson(strResponse)
            GetAccessToken = objJson("access_token")
        Else
            MsgBox "Error getting access token: " & .Status & " - " & .responseText
            GetAccessToken = ""
        End If
    End With

    Set objHTTP = Nothing
    Set objJson = Nothing
End Function


' Main function to call Vertex AI
Sub CallVertexAI(ByVal prompt As String)
    Dim objHTTP As Object
    Dim strURL As String
    Dim strRequestBody As String
    Dim strResponse As String
    Dim objJson As Object
    Dim accessToken As String
    Dim results As Variant
    Dim i As Long

    ' 1. Get Access Token (ADC or JWT)
    accessToken = GetAccessToken() ' ADC method
    'accessToken = GetJwtToken("path/to/your/service_account_key.json") ' Service Account method (DO NOT STORE IN CODE)

    If accessToken = "" Then
        MsgBox "Failed to obtain access token.  Check authentication configuration."
        Exit Sub
    End If


    ' 2. Prepare the Request Body (JSON)
    strRequestBody = BuildRequestBody(prompt)

    ' 3. Create the HTTP Request
    Set objHTTP = CreateObject("MSXML2.XMLHTTP60") ' Or XMLHTTP, depending on your version
    strURL = API_ENDPOINT

    With objHTTP
        .Open "POST", strURL, False ' Synchronous request
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & accessToken
        .Send strRequestBody

        strResponse = .responseText

        If .Status = 200 Then
            ' 4. Parse the JSON Response
            Set objJson = JsonConverter.ParseJson(strResponse)

            ' 5. Extract the Prediction Results
            '  Adjust this based on the actual response structure of your model
            '  For text-bison, it's likely in the 'predictions' array.

            results = objJson("predictions")

            If IsArray(results) Then
              For i = LBound(results) To UBound(results)
                Debug.Print "Result " & i + 1 & ": " & results(i)("content")  'Adjust "content" if necessary
                ' You can output to a cell in Excel here instead of Debug.Print
              Next i
            Else
                Debug.Print "Unexpected response format. Check your model and parsing logic."
            End If
        Else
            MsgBox "Error calling Vertex AI: " & .Status & " - " & .responseText
        End If
    End With

    ' Clean up
    Set objHTTP = Nothing
    Set objJson = Nothing

End Sub


Sub generateScen()

Dim http As Object
Dim Json As Object
Dim apikey As String
Dim prompt As String
Dim response As String
Dim JsonResponse As Object
Dim JsonString As String
Dim Model As Variant
Dim url As String
Dim result As String
Dim rowIndex As Integer
Dim ScenarioArray As Object
Dim Scenario As Object
Dim lastrow As Integer
Dim i As Integer
Dim simulatedResponse As String
Dim ws As Worksheet
Dim wsReq As Worksheet
Dim wsScen As Worksheet

' Define Sheets
Set wsReq = ThisWorkbook.Sheets("Requirement")
Set wsScen = ThisWorkbook.Sheets("Scenarios")

' Clear previous data
wsScen.Cells.Clear
wsScen.Cells(1, 1).Value = "Scenario Name"
wsScen.Cells(1, 2).Value = "Description"
wsScen.Rows(1).Font.Bold = True


' Intialize the rowIndex
rowIndex = 2



Dim cellValue As Variant
cellValue = wsReq.Range("E2").Value


' prepare prompt with requirements
prompt = "List all the Scenarios end to end scenarios in Json format with fields 'ScenarioName',' :" & vbCrLf



' Simulated JSon response
simulatedResponse = "{""Scenarios"": [" & _
                       "{""ScenarioName"": ""Configure Qualifier with OR Condition"", ""Description"": ""User configures a qualifier with an OR condition (e.g., Color = Black OR Blue) and verifies it displays correctly in the Promo Config UI.""}," & _
                       "{""ScenarioName"": ""Configure Qualifier with AND Condition (Default)"", ""Description"": ""User configures a qualifier with the default AND condition (e.g., Color = Black AND Season = Winter) and confirms it works as expected.""}," & _
                       "{""ScenarioName"": ""Configure Target with OR Condition"", ""Description"": ""User sets a target with an OR condition (e.g., Product Class = Shirts OR Tops) and ensures it is reflected in the UI.""}," & _
                       "{""ScenarioName"": ""Configure Target with AND Condition"", ""Description"": ""User sets a target with an AND condition (e.g., Product Class = Shirts AND Style = 15003) and validates the behavior.""}," & _
                       "{""ScenarioName"": ""Mix AND/OR Conditions Across Qualifier and Target"", ""Description"": ""User configures a qualifier with OR (Color = Black OR Blue) and a target with AND (Product Class = Shirts AND Style = 15003), then verifies both conditions are correctly applied.""}," & _
                       "{""ScenarioName"": ""Edit Existing Promo with AND/OR Conditions"", ""Description"": ""User edits an existing promo, modifies the AND/OR conditions (e.g., changes OR to AND), and confirms the updated rules are saved and displayed.""}," & _
                       "{""ScenarioName"": ""Display AND/OR in Inclusion/Exclusion Rules List"", ""Description"": ""User views the Inclusion/Exclusion Rules List screen and verifies that AND/OR conditions between attributes (e.g., Order, Payment) are correctly shown.""}," & _
                       "{""ScenarioName"": ""Confirm Screen with AND/OR Conditions"", ""Description"": ""User navigates to the Confirm screen and ensures AND/OR conditions across Qualifier and Target are accurately displayed.""}" & _
                       "]}"

' Parse the simulated response
Set Json = JsonConverter.ParseJson(simulatedResponse)



' Validate JsonResponse
If Json Is Nothing Or Not Json.Exists("Scenarios") Then
    MsgBox "Error: 'Scenarios' key is not found in JSON response!", vbCritical
    Exit Sub
End If


'Assign Scenario Array
Set ScenarioArray = Json("Scenarios")


' Ensure scenarioArray is an iterable object
If Not IsObject(ScenarioArray) Then
    MsgBox "Error: 'Scenarios' is not an object or array!", vbCritical
    Exit Sub
End If



' Iterate through scenarios and populate the spreadsheet
For Each Scenario In ScenarioArray
    wsScen.Cells(rowIndex, 1).Value = Scenario("ScenarioName")
    wsScen.Cells(rowIndex, 2).Value = Scenario("Description")
    rowIndex = rowIndex + 1
     
Next Scenario

MsgBox "generated scenrios successfully!", vbInformation


End Sub



