<!-- Copyright (c) 2019 ActivePDF, Inc. -->
<!-- ActivePDF WebGrabber 2016 -->
<!-- Example generated 01/03/19  -->

<%
Dim strPath, results

strPath = Server.MapPath(".") & "\"

' Instantiate Object
Set oWG = Server.CreateObject("APWebGrabber.Object")

' Set the quality options for the created PDF (IE engine only)
' For custom settings to take effect set the configuration to custom
oWG.PredefinedSetting = 0

' Specifies if ASCII85 encoding should be applied to binary streams
oWG.ASCIIEncode = true

' Automatically control the page orientation based on text flow
oWG.AutoRotate = true

' Specifies if CMYK colors should be converted to RGB
oWG.ConvertCMYKToRGB = true

' Set the DPI for the created PDF
oWG.Resolution = 300.0

' Color Image Quality Settings
oWG.ColorImageDownsampleThreshold = 1
oWG.ColorImageDownsampleType = 0
oWG.ColorImageFilter = 2
oWG.ColorImageResolution = 72

' Gray Image Quality Settings
oWG.GrayImageDownsampleThreshold = 1
oWG.GrayImageDownsampleType = 0
oWG.GrayImageFilter = 2
oWG.GrayImageResolution = 72

' Monochrome Image Quality Settings
oWG.MonoImageDownsampleThreshold = 1
oWG.MonoImageDownsampleType = 0
oWG.MonoImageFilter = 2
oWG.MonoImageResolution = 72

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
oWG.EngineToUse = 1

' Convert the HTML background (IE engine only)
oWG.PrintBackground = true

' PDF output location and filename
oWG.OutputDirectory = strPath
oWG.NewDocumentName = "quality.pdf"

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