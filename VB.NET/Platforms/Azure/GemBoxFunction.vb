Imports System.IO
Imports Microsoft.AspNetCore.Mvc
Imports Microsoft.Azure.WebJobs
Imports Microsoft.Azure.WebJobs.Extensions.Http
Imports Microsoft.AspNetCore.Http
Imports Microsoft.Extensions.Logging
Imports GemBox.Presentation

Module GemBoxFunction
#Disable Warning BC42356 ' This async method lacks 'Await'.
    <FunctionName("GemBoxFunction")>
    Async Function Run(<HttpTrigger(AuthorizationLevel.Anonymous, "get", Route:=Nothing)> req As HttpRequest, log As ILogger) As Task(Of IActionResult)
#Enable Warning BC42356

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As New PresentationDocument()

        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 5, 4, LengthUnit.Centimeter)

        Dim paragraph = textBox.AddParagraph()

        paragraph.AddRun("Hello World!")

        Dim fileName = "Output.pptx"
        Dim options = SaveOptions.Pptx

        Using stream As New MemoryStream()
            presentation.Save(stream, options)
            Return New FileContentResult(stream.ToArray(), options.ContentType) With { .FileDownloadName = fileName }
        End Using

    End Function
End Module
