using System.Linq;
using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("FindAndReplace.pptx");

        var slide = presentation.Slides[0];

        slide.TextContent.Replace("companyName", "Acme Corporation");

        var items = new string[][]
        {
            new string[] { "January",  "$14.2M", "$1.6M", "$12.5M", "$2.3M" },
            new string[] { "February", "$15.2M", "$0.6M", "$11.3M", "$4.3M" },
            new string[] { "March",    "$17.2M", "$0M",   "$12.1M", "$52.3M" },
            new string[] { "April",    "$7.2M",  "$1.6M", "$7.2M",  "$0M" },
            new string[] { "May",      "$3.2M",  "$1.6M", "$0M",    "$3.2M" }
        };

        var table = slide.Content.Drawings.OfType<GraphicFrame>().First().Table;

        // Clone the row with custom text tags (e.g. '@month', '@revenue', etc.).
        for (int i = 0; i < items.Length - 1; i++)
            table.Rows.AddClone(table.Rows[1]);

        // Replace custom text tags with data.
        for (int i = 0; i < items.Length; i++)
        {
            var row = table.Rows[i + 1];
            var item = items[i];
            row.TextContent.Replace("@month", item[0]);
            row.TextContent.Replace("@revenue", item[1]);
            row.TextContent.Replace("@cashExpense", item[2]);
            row.TextContent.Replace("@operatingIncome", item[3]);
            row.TextContent.Replace("@operatingExpense", item[4]);
        }

        // Find all text content with "0M" value in the table.
        var zeroValueRanges = table.TextContent.Find("0M");

        presentation.Save("Find And Replace.pptx");
    }
}