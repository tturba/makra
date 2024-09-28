Sub DownloadFile()
    Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
    Set objADODBStream = CreateObject("ADODB.Stream")
    objXMLHTTP.Open "GET", "[URL]", False
    objXMLHTTP.Send
    objADODBStream.Type = 1
    objADODBStream.Open
    objADODBStream.Write objXMLHTTP.responseBody
    objADODBStream.savetofile "ts.jpg", 2
End Sub
