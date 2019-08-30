# GetBusArrivalData
https://www.mytransport.sg/content/dam/datamall/datasets/LTA_DataMall_API_User_Guide.pdf

'Microsoft HTML Object Library
'Microsoft XML, v6.0
'Microsoft Forms 2.0 Object Library

Sub GetBusArrivalData()
    Dim ws As Worksheet: Set ws = Worksheets("API")
    Dim reqURL As String: reqURL = ws.[APIurl]
    Dim strKey As String: strKey = ws.[APIkey]
    Dim lineS As Variant
    Dim req As New MSXML2.ServerXMLHTTP60
    Dim accept As String: accept = "application/json"
    
    req.Open "GET", reqURL, False
    req.setRequestHeader "AccountKey", strKey
    req.setRequestHeader "accept", accept
    req.send
    
    ws.Range("A4").Value = req.responseText
End Sub
