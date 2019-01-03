<!-- Copyright (c) 2019 ActivePDF, Inc. -->
<!-- ActivePDF WebGrabber 2016 -->
<!-- Example generated 01/03/19  -->

<%
Dim strPath, results

strPath = Server.MapPath(".") & "\"

' Instantiate Object
Set oWG = Server.CreateObject("APWebGrabber.Object")

' Stamp Images and Text onto the output PDF
oWG.AddStampCollection "TXTinternal"
oWG.StampFont = "Helvetica"
oWG.StampFontSize = 108
oWG.StampFontTransparency = 0.1
oWG.StampRotation = 45.0

oWG.StampFillMode = 2
oWG.SetStampColor 255, 0, 0, 0
oWG.SetStampStrokeColor 100, 0, 0, 0

oWG.AddStampText 116.0, 156.0, "Internal Only"

oWG.AddStampCollection "IMGimage"
oWG.AddStampImage strPath & "logo.png", 508.0, 16.0, 32.0, 32.0, true

' Set whether the stamp collection(s) appears in the background or foreground
oWG.StampBackground = 0

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

' PDF output location and filename
oWG.OutputDirectory = strPath
oWG.NewDocumentName = "stamped.pdf"

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

' If there will be multiple conversions in the same instance
' you can clear the stamp collections or remove one individually
' in order to change what is stamped on the next conversion
oWG.RemoveStampCollection "TXTinternal"
oWG.ClearStampCollections 

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