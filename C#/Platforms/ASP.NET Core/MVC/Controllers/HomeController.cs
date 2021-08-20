using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using PresentationCoreMvc.Models;
using GemBox.Presentation;
using System.Linq;

namespace PresentationCoreMvc.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment environment;

        public HomeController(IWebHostEnvironment environment)
        {
            this.environment = environment;

            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        public IActionResult Index()
        {
            return View(new CardModel());
        }

        public FileStreamResult Download(CardModel model)
        {
            // Load template presentation.
            var path = Path.Combine(this.environment.ContentRootPath, "CardWithPlaceholderElements.pptx");
            var presentation = PresentationDocument.Load(path);

            // Get first slide.
            var slide = presentation.Slides[0];

            // Get placeholder elements.
            var placeholders = slide.Content.Drawings.OfType<Shape>()
                .Where(s => s.Placeholder != null && s.Placeholder.PlaceholderType == PlaceholderType.Text);

            // Set text on placeholders.
            var top = placeholders.First(p => p.Name == "Top Placeholder");
            top.TextContent.LoadText(model.Top);
            var middle = placeholders.First(p => p.Name == "Middle Placeholder");
            middle.TextContent.LoadText(model.Middle);
            var bottom = placeholders.First(p => p.Name == "Bottom Placeholder");
            bottom.TextContent.LoadText(model.Bottom);

            // Save presentation in specified file format.
            var stream = new MemoryStream();
            presentation.Save(stream, model.Options);

            // Download file.
            return File(stream, model.Options.ContentType, $"OutputFromView.{model.Format.ToLower()}");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel() { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}

namespace PresentationCoreMvc.Models
{
    public class CardModel
    {
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Top { get; set; } = "Happy Birthday Jane!";
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Middle { get; set; } = "May this be your best birthday ever.\nMay your joy never end.";
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Bottom { get; set; } = "from John 😊";
        public string Format { get; set; } = "PPTX";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["PPTX"] = new PptxSaveOptions(),
            ["PDF"] = new PdfSaveOptions(),
            ["XPS"] = new XpsSaveOptions(),
            ["BMP"] = new ImageSaveOptions(ImageSaveFormat.Bmp),
            ["PNG"] = new ImageSaveOptions(ImageSaveFormat.Png),
            ["JPG"] = new ImageSaveOptions(ImageSaveFormat.Jpeg),
            ["GIF"] = new ImageSaveOptions(ImageSaveFormat.Gif),
            ["TIF"] = new ImageSaveOptions(ImageSaveFormat.Tiff)
        };
    }
}
