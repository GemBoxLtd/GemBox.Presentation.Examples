using System.IO;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using GemBox.Presentation;

public class GemBoxFunction
{
    [Function("GemBoxFunction")]
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req)
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 5, 4, LengthUnit.Centimeter);
        var paragraph = textBox.AddParagraph();
        paragraph.AddRun("Hello World!");

        var fileName = "Output.pptx";
        var options = SaveOptions.Pptx;

        using var stream = new MemoryStream();
        presentation.Save(stream, options);
        var bytes = stream.ToArray();

        var response = req.CreateResponse(HttpStatusCode.OK);
        response.Headers.Add("Content-Type", options.ContentType);
        response.Headers.Add("Content-Disposition", "attachment; filename=" + fileName);
        await response.Body.WriteAsync(bytes, 0, bytes.Length);
        return response;
    }
}