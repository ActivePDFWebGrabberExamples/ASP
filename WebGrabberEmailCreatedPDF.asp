<!-- Copyright (c) 2019 ActivePDF, Inc. -->
<!-- ActivePDF WebGrabber 2016 -->
<!-- Example generated 01/03/19  -->

<%
Dim strPath, results

strPath = Server.MapPath(".") & "\"

' Instantiate Object
Set oWG = Server.CreateObject("APWebGrabber.Object")

' Add an email
oWG.AddEMail 

' Set server information
oWG.SetSMTPInfo "0.0.0.0", 25
oWG.SetSMTPCredentials "john.doe", "activePDF", "asdfasdf"

' Set email addresses
oWG.SetSenderInfo "John Doe", "john.doe@asdidlwenra.com"
oWG.SetReplyToInfo "John Doe", "john.doe@asdidlwenra.com"
oWG.SetRecipientInfo "Jane Doe", "jane.doe@asdidlwenra.com"
oWG.AddToCC "Jim Doe", "jim.doe@asdidlwenra.com"
oWG.AddToBcc "Janice Doe", "janice.doe@asdidlwenra.com"

' Subject and Body
oWG.EMailSubject = "PDF Delivery from activePDF"
oWG.SetEMailBody "<html><body style='background-color: #EEE; padding: 4px;'>Here is your PDF!</body></html>", true

' Attachments - Binary attachments can be added with AddEMailBinaryAttachment
oWG.AddEMailAttachment strPath & "x.pdf"

' Other email options
oWG.EMailReadReceipt = false
oWG.EMailAttachOutput = true

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
oWG.NewDocumentName = "email.pdf"

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

' If running multiple conversions in one instance:
' One email can be removed before the next conversion
oWG.RemoveEMail "john.doe@activepdf.com"
' An attachment can be removed
oWG.ClearEMailAttachments 
' or all emails can be removed
oWG.ClearEMails 

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