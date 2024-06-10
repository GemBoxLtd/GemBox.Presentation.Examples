using GemBox.Presentation;
using System.Collections.Generic;

namespace BlazorServerApp.Data
{
    public class CardModel
    {
        public string Top { get; set; } = "Happy Birthday Jane!";
        public string Middle { get; set; } = "May this be your best birthday ever.\nMay your joy never end.";
        public string Bottom { get; set; } = "from John 😊";
        public string Format { get; set; } = "PPTX";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["PPTX"] = new PptxSaveOptions(),
            ["PDF"] = new PdfSaveOptions(),
            ["XPS"] = new XpsSaveOptions(), // XPS is supported only on Windows.
            ["PNG"] = new ImageSaveOptions(ImageSaveFormat.Png),
            ["JPG"] = new ImageSaveOptions(ImageSaveFormat.Jpeg),
            ["BMP"] = new ImageSaveOptions(ImageSaveFormat.Bmp),
            ["GIF"] = new ImageSaveOptions(ImageSaveFormat.Gif),
            ["TIF"] = new ImageSaveOptions(ImageSaveFormat.Tiff)
        };
    }
}