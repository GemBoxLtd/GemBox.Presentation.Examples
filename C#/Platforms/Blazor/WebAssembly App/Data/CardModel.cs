using System.Collections.Generic;
using GemBox.Presentation;

namespace BlazorWebAssemblyApp.Data
{
    public class CardModel
    {
        public string Top { get; set; } = "Happy Birthday Jane!";
        public string Middle { get; set; } = "May this be your best birthday ever.\nMay your joy never end.";
        public string Bottom { get; set; } = "from John 😊";
        public string Format { get; set; } = "PDF";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["PPTX"] = new PptxSaveOptions(),
            ["PDF"] = new PdfSaveOptions()
        };
    }
}