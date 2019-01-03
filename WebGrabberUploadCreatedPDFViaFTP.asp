<!-- Copyright (c) 2019 ActivePDF, Inc. -->
<!-- ActivePDF WebGrabber 2016 -->
<!-- Example generated 01/03/19  -->

<%
Dim strPath, results

strPath = Server.MapPath(".") & "\"

' Instantiate Object
Set oWG = Server.CreateObject("APWebGrabber.Object")

' Setup the FTP request supplying credentials if needed
oWG.AddFTPRequest "127.0.0.1", "/folder"
oWG.SetFTPCredentials "user", "pass"

' Set which files will upload with the FTP request
' To attach a binary file use AddFTPBinaryAttachment
oWG.FTPAttachOutput = true
oWG.AddFTPAttachment strPath & "file.txt"

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
oWG.NewDocumentName = "ftp.pdf"

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

' Options available to clear or remove FTP requests are only
' needed if the object remains instantiated.
oWG.RemoveFTP "127.0.0.1", "/folder"
oWG.ClearFTPAttachments 
oWG.ClearFTPs 

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