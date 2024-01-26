using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PresentationCorePages.Models;
using GemBox.Presentation;

namespace PresentationCorePages.Pages
{
    public class IndexModel : PageModel
    {
        private readonly IWebHostEnvironment environment;

        [BindProperty]
        public CardModel Card { get; set; }

        public IndexModel(IWebHostEnvironment environment)
        {
            this.environment = environment;
            this.Card = new CardModel();

            // If using the Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        public void OnGet() { }

        public FileContentResult OnPost()
        {
            // Load template presentation.
            var path = Path.Combine(this.environment.ContentRootPath, "CardWithPlaceholderTexts.pptx");
            var presentation = PresentationDocument.Load(path);

            // Get first slide.
            var slide = presentation.Slides[0];

            // Execute find and replace operations.
            slide.TextContent.Replace("{{Top Text}}", this.Card.Top);
            slide.TextContent.Replace("{{Middle Text}}", this.Card.Middle);
            slide.TextContent.Replace("{{Bottom Text}}", this.Card.Bottom);

            // Save presentation in specified file format.
            using var stream = new MemoryStream();
            presentation.Save(stream, this.Card.Options);

            // Download file.
            return File(stream.ToArray(), this.Card.Options.ContentType, $"OutputFromPage.{this.Card.Format.ToLower()}");
        }
    }
}

namespace PresentationCorePages.Models
{
    public class CardModel
    {
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Top { get; set; } = "Happy Birthday Jane!";
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Middle { get; set; } = "May this be your best birthday ever.\nMay your joy never end.";
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Bottom { get; set; } = "from John 😊";
        public string Format { get; set; } = "PDF";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["PPTX"] = new PptxSaveOptions(),
            ["PDF"] = new PdfSaveOptions(),
            ["XPS"] = new XpsSaveOptions(), // XPS is supported only on Windows.
            ["BMP"] = new ImageSaveOptions(ImageSaveFormat.Bmp),
            ["PNG"] = new ImageSaveOptions(ImageSaveFormat.Png),
            ["JPG"] = new ImageSaveOptions(ImageSaveFormat.Jpeg),
            ["GIF"] = new ImageSaveOptions(ImageSaveFormat.Gif),
            ["TIF"] = new ImageSaveOptions(ImageSaveFormat.Tiff),
            ["SVG"] = new ImageSaveOptions(ImageSaveFormat.Svg)
        };
    }
}
