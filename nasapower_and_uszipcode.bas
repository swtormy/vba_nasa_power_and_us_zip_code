Attribute VB_Name = "subs"
Function response(url, addition)
    Set HTTP = CreateObject("MSXML2.XMLHTTP")
    
    HTTP.Open "GET", url & addition, False
    HTTP.send
    
    response = HTTP.responseText
    
End Function


Sub parse()
    addition = Range("B3")
    
    Set JSON = JsonConverter.ParseJson(response("https://public.opendatasoft.com/api/records/1.0/search/?dataset=us-zip-code-latitude-and-longitude&q=", addition))

    'records/0/fields/city


    latitude = JSON("records")(1)("fields")("latitude")
    longitude = JSON("records")(1)("fields")("longitude")
    city = JSON("records")(1)("fields")("city")
    State = JSON("records")(1)("fields")("state")
    
    Cells(4, 2).Value = latitude
    Cells(5, 2).Value = longitude
    Cells(6, 2).Value = city
    Cells(7, 2).Value = State
    
    addition = Replace("lat=" & latitude & "&lon=" & longitude, ",", ".")
    Set JSON = JsonConverter.ParseJson(response("https://power.larc.nasa.gov/cgi-bin/v1/DataAccess.py?&request=execute&identifier=SinglePoint&parameters=ALLSKY_SFC_SW_DWN&userCommunity=SSE&tempAverage=CLIMATOLOGY&outputList=CSV&", addition))


    For i = 1 To JSON("features")(1)("properties")("parameter")("ALLSKY_SFC_SW_DWN").Count Step 1
        'Debug.Print JSON("features")(1)("properties")("parameter")("ALLSKY_SFC_SW_DWN")(CStr(i))
        Cells(10, 1 + i).Value = Replace(JSON("features")(1)("properties")("parameter")("ALLSKY_SFC_SW_DWN")(CStr(i)), ",", ".")
    Next i
End Sub
