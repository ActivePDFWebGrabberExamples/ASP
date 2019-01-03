<!-- Copyright (c) 2019 ActivePDF, Inc. -->
<!-- ActivePDF WebGrabber 2016 -->
<!-- Example generated 01/03/19  -->

<%
Dim strPath, results

strPath = Server.MapPath(".") & "\"

' Instantiate Object
Set oWG = Server.CreateObject("APWebGrabber.Object")

' Rendering engine used for the HTML
' 0 = Native, 1 = IE
oWG.EngineToUse = 1

' Convert the HTML background (IE engine only)
oWG.PrintBackground = 1

' PDF output location and filename
oWG.OutputDirectory = strPath
oWG.NewDocumentName = "WG-new.pdf"

' HTML to convert
' Examples:
' http://domain.com/path/file.aspx
' c:\folder\file.html
oWG.URL = "http://examples.activepdf.com/samples/doc"

' Perform the HTML to PDF conversion
Set results = oWG.ConvertToPDF("127.0.0.1", 52525)
If results.WebGrabberStatus <> 0 Then
  ErrorHandler "ConvertToPDF", results, results.WebGrabberStatus
End If

' Release Object
Set oWG = Nothing

' Process Complete
Response.Write "Done!"

' Error Handling
Sub ErrorHandler(method, oResult, errorStatus)
  Response.Write("Error with " & method & ": <br/>" _
    & errorStatus & "<br/>" _
    & oResult.details)
  Response.End
End Sub
%>