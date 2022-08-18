using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using GemBox.Presentation;

public static class GemBoxFunction
{
    [FunctionName("GemBoxFunction")]
#pragma warning disable CS1998 // Async method lacks 'await' operators.
    public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req, ILogger log)
#pragma warning restore CS1998
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 5, 4, LengthUnit.Centimeter);

        var paragraph = textBox.AddParagraph();

        paragraph.AddRun("Hello World!");

        var fileName = "Output.pptx";
        var options = SaveOptions.Pptx;

        using (var stream = new MemoryStream())
        {
            presentation.Save(stream, options);
            return new FileContentResult(stream.ToArray(), options.ContentType) { FileDownloadName = fileName };
        }
    }
}