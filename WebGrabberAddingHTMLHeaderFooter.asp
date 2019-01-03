<!-- Copyright (c) 2019 ActivePDF, Inc. -->
<!-- ActivePDF WebGrabber 2016 -->
<!-- Example generated 01/03/19  -->

<%
Dim strPath, results

strPath = Server.MapPath(".") & "\"

' Instantiate Object
Set oWG = Server.CreateObject("APWebGrabber.Object")

' Enable extra logging (logging should only be used while troubleshooting)
' C:\ProgramData\activePDF\Logs\
oWG.Debug = true

' Fast web view
oWG.LinearizePDF = true

' Time to wait for conversion to complete (in seconds)
' Set the amount of seconds before a request will time out
oWG.Timeout = 40

' Margins (Top, Bottom, Left, Right) 1.0 = 1"
oWG.SetMargins 0.75, 0.75, 0.75, 0.75

' 0 = Portrait, 1 = Landscape
oWG.Orientation = 0

' Rendering engine used for the HTML
' 0 = Native, 1 = IE
oWG.EngineToUse = 0

' Convert HTML fields to PDF fields
oWG.PreserveButtons = false
oWG.PreserveCheckBoxes = false
oWG.PreserveDropDowns = false
oWG.PreserveRadioButtons = false
oWG.PreserveTextBoxes = false

' Convert links
' 0 = None
' 1 = Internal
' 2 = External
' 3 = Both (default)
oWG.GenerateLinks = 3

' Convert h tags into bookmarks
oWG.GenerateBookmarks = true

' Enable flash conversion
oWG.EmbedFlash = 1

' Add a header and footer to the PDF output
' Must use full paths to additional files
' Using %cp% of %tp% in the HTML equals current and total page numbers
oWG.HeaderHTML = "<html><body>"
oWG.HeaderHTML = "<div style='float: left;'>activePDF.com</div>"
oWG.HeaderHTML = "<div style='float: right;'>01/03/2019 10:17PM</div>"
oWG.HeaderHTML = "</body></html>"
oWG.HeaderHeight = 0.5

oWG.FooterHTML = "<html><body>"
oWG.FooterHTML = "<div style='text-align: center;'>%cp% of %tp%</div>"
oWG.FooterHTML = "</body></html>"
oWG.FooterHeight = 0.5

' PDF output location and filename
oWG.OutputDirectory = strPath
oWG.NewDocumentName = "headfoot.pdf"

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

' Clear variables from HeaderHTML and HeaderURL properties
oWG.ClearHeaderHTML 

' Clear variables from FooterHTML and FooterURL properties
oWG.ClearFooterHTML 

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