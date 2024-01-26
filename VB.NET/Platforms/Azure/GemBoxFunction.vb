Imports System.IO
Imports System.Net
Imports System.Threading.Tasks
Imports Microsoft.Azure.Functions.Worker
Imports Microsoft.Azure.Functions.Worker.Http
Imports GemBox.Presentation

Public Class GemBoxFunction
    <[Function]("GemBoxFunction")>
    Public Async Function Run(<HttpTrigger(AuthorizationLevel.Anonymous, "get")> req As HttpRequestData) As Task(Of HttpResponseData)

        ' If using the Professional version, put your serial key below.
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
            Dim bytes = stream.ToArray()

            Dim response = req.CreateResponse(HttpStatusCode.OK)
            response.Headers.Add("Content-Type", options.ContentType)
            response.Headers.Add("Content-Disposition", "attachment; filename=" & fileName)
            Await response.Body.WriteAsync(bytes, 0, bytes.Length)
            Return response
        End Using

    End Function
End Class
