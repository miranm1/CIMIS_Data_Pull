Attribute VB_Name = "Module1"
Sub ReadCIMIS()
 ' initialize req'd variables and sheet
    Dim http As Object
    Dim webServiceURL As String
    ActiveSheet.Range("A8:ZZ99999").ClearContents 'clear worksheet
    
    origSelectedCell = ActiveCell.Address ' save the the original address of the cell to return to
    ActiveSheet.Range("A6").Select        'start at cell a6 for collecting weather stations
    wscount = WorksheetFunction.CountA(Range("6:6")) / 2
    Application.ScreenUpdating = False
    wscounter = 1
    While ActiveCell.Value <> ""
    
        'created a status bar to display progress
        Application.StatusBar = Format(wscounter, "000") & " of " & Format(wscount, "000") & " -- " & WorksheetFunction.Floor(100 * (wscounter - 1) / wscount, 1) & "% Complete | Reading data for " & ActiveCell.Offset(-1, 1).Value
        Application.Wait Now + #12:00:01 AM#
        
        'build api url for series api endpoint
        'http://et.water.ca.gov/api/data?appKey=ab1b4d2a-086c-4bb7-aa66-8573e194a589&targets=41&startDate=2010-01-01&endDate=2010-01-05&dataItems=day-precip
        'target = station number
        'dataItems = data type   precip/eto
        
        webServiceURL = "http://et.water.ca.gov/api/data?appKey=ab1b4d2a-086c-4bb7-aa66-8573e194a589&targets=" & _
            ActiveCell.Offset(0, 1).Value & "&startDate=" & Format(Range("B2").Value, "yyyy-mm-dd") & "&endDate=" & _
            Format(Range("B3").Value, "yyyy-mm-dd") & "&dataItems=day-precip"
            
        Set http = CreateObject("msxml2.xmlhttp")
            http.Open "GET", webServiceURL, False
            http.Send
        
        ' Example of JSON format retrieved from API
        ' {
        '   "Data": {
        '       "Providers": [
        '           {
        '               "Name": "cimis",
        '               "Type": "station",
        '               "Owner": "water.ca.gov",
        '               "Records": [
        '                   {
        '                       "Date": "2010-01-01",
        '                       "Julian": "1",
        '                       "Station": "135",
        '                       "Standard": "english",
        '                       "ZipCodes": "92228, 92227, 92226, 92225",
        '                       "Scope": "daily",
        '                       "DayEto": {
        '                           "Value": "0",
        '                           "Qc": "H",
        '                           "Unit": "(in)"
        '                       }
        '                   }
        '               ]
        '           }
        '       ]
        '   }
        '}
        
    On Error GoTo ErrCatcher
                    
        ' Parse JSON response
        JsonText = http.ResponseText
        'creating 2 variant variables because values are nested within json file
        Dim Provider As Variant
        Dim Record As Variant
        
        'parse json file
        Set JSON = JsonConverter.ParseJson(JsonText)
        Dim RecordCount As Long ' variable created to determine how many items are in the json response
        
        RecordCount = 0
        
        ' For loop to count the number of Records in the JSON file
        For Each Provider In JSON("Data")("Providers")
        For Each Record In Provider("Records")
        RecordCount = RecordCount + 1
        Next Record
        Next Provider
        
        'Declare array for and create size of array
        Dim Values As Variant
        ReDim Values(RecordCount, 1)
        Dim Value As Dictionary
        Dim i As Long
        i = 0
        
        ' For loop to store values for printing on to spreadsheet
        For Each Provider In JSON("Data")("Providers")
            For Each Record In Provider("Records")
            Values(i, 0) = Record("Date")
            Values(i, 1) = Record("DayPrecip")("Value")
            i = i + 1
            Next Record
        Next Provider
        
        
        
        ' WRITE TS DATA
        a = UBound(Values)
        If UBound(Values) < 1 Then 'Catch case for no data
            ActiveCell.Offset(1, 0).Value = "No Data..."
        Else
            ActiveSheet.Range(Cells(ActiveCell.Offset(2, 0).Row, ActiveCell.Column), Cells(ActiveCell.Offset(2, 0).Row + RecordCount, ActiveCell.Column + 1)) = Values
        End If
        
    
ErrCatcher:
        Set http = Nothing
        ' MOVE TO NEXT COLUMN OF DATA
        ActiveCell.Offset(0, 2).Select
        wscounter = wscounter + 1
        
    Wend
    Application.ScreenUpdating = True
    ActiveSheet.Range(origSelectedCell).Select
    Application.StatusBar = "Done!"


End Sub
