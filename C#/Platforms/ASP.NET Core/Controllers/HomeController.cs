using System;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using GemBox.Presentation;

namespace PresentationCore.Controllers
{
    public class HomeController : Controller
    {
        static HomeController()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Download(string format)
        {
            // Create new presentation.
            var presentation = new PresentationDocument();

            // Add slide.
            var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

            // Add textbox.
            var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 5, 4, LengthUnit.Centimeter);

            // Add paragraph.
            var paragraph = textBox.AddParagraph();

            // Add text.
            paragraph.AddRun("Hello World!");

            using (var stream = new MemoryStream())
            {
                // Save presentation to stream in specified format.
                SaveOptions options = GetSaveOptions(format);
                presentation.Save(stream, options);

                // Download presentation to client's browser.
                return File(stream.ToArray(), options.ContentType, "Hello World." + format.ToLower());
            }
        }

        private static SaveOptions GetSaveOptions(string format)
        {
            switch (format.ToUpper())
            {
                case "PPTX":
                    return SaveOptions.Pptx;
                case "PDF":
                    return SaveOptions.Pdf;

                case "XPS":
                case "PNG":
                case "JPG":
                case "GIF":
                case "TIF":
                case "BMP":
                case "WMP":
                    throw new InvalidOperationException("To enable saving to XPS or image format, add 'Microsoft.WindowsDesktop.App' framework reference.");

                default:
                    throw new NotSupportedException();
            }
        }
    }
}