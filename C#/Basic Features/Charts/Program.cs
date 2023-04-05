using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using GemBox.Presentation;
using GemBox.Pdf;
using GemBox.Pdf.Content;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your GemBox.Presentation serial key below.
        GemBox.Presentation.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, put your GemBox.Pdf serial key below.
        GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
        Example3();
        Example4();
    }

    static void Example1()
    {
        var presentation = new PresentationDocument();

        // Add new PowerPoint presentation slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Add simple PowerPoint presentation title.
        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle,
            116.8, 20, 105, 10, GemBox.Presentation.LengthUnit.Millimeter);

        textBox.AddParagraph().AddRun("New presentation with chart element.");

        // Create PowerPoint chart and add it to slide.
        var chart = slide.Content.AddChart(GemBox.Presentation.ChartType.Bar,
            49.3, 40, 240, 120, GemBox.Presentation.LengthUnit.Millimeter);

        // Get underlying Excel chart.
        ExcelChart excelChart = (ExcelChart)chart.ExcelChart;
        ExcelWorksheet worksheet = excelChart.Worksheet;

        // Add data for Excel chart.
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["A2"].Value = "John Doe";
        worksheet.Cells["A3"].Value = "Fred Nurk";
        worksheet.Cells["A4"].Value = "Hans Meier";
        worksheet.Cells["A5"].Value = "Ivan Horvat";

        worksheet.Cells["B1"].Value = "Salary";
        worksheet.Cells["B2"].Value = 3600;
        worksheet.Cells["B3"].Value = 2580;
        worksheet.Cells["B4"].Value = 3200;
        worksheet.Cells["B5"].Value = 4100;

        // Select data.
        excelChart.SelectData(worksheet.Cells.GetSubrange("A1:B5"), true);

        presentation.Save("Created Chart.pdf");
    }

    static void Example2()
    {
        // Load input file and save it in selected output format
        var presentation = PresentationDocument.Load("Chart.pptx");

        // Get PowerPoint chart.
        var chart = ((GraphicFrame)presentation.Slides[0].Content.Drawings[0]).Chart;

        // Get underlying Excel chart and cast it as LineChart.
        var lineChart = (LineChart)chart.ExcelChart;

        // Get underlying Excel sheet and add new cell values.
        var sheet = lineChart.Worksheet;
        sheet.Cells["D1"].Value = "Series 3";
        sheet.Cells["D2"].Value = 8.6;
        sheet.Cells["D3"].Value = 5;
        sheet.Cells["D4"].Value = 7;
        sheet.Cells["D5"].Value = 9;

        // Add new line series to the LineChart.
        lineChart.Series.Add(sheet.Cells["D1"].StringValue, "Sheet1!D2:D5");

        presentation.Save("Updated Chart.pptx");
    }

    static void Example3()
    {
        var presentation = new PresentationDocument();
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);
        var chart = slide.Content.AddChart(GemBox.Presentation.ChartType.Column,
            5, 5, 15, 10, GemBox.Presentation.LengthUnit.Centimeter);

        // Get underlying Excel chart.
        var columnChart = (ColumnChart)chart.ExcelChart;

        // Set chart's category labels from array.
        columnChart.SetCategoryLabels(new string[] { "Columns 1", "Columns 2", "Columns 3" });

        // Add chart's series from arrays.
        columnChart.Series.Add("Values 1", new double[] { 3.4, 1.1, 3.7 });
        columnChart.Series.Add("Values 2", new double[] { 4.4, 3.9, 3.5 });
        columnChart.Series.Add("Values 3", new double[] { 2.9, 4.1, 1.9 });

        presentation.Save("Created Chart from Array.pptx");
    }

    static void Example4()
    {
        var presentation = PresentationDocument.Load("Chart.pptx");
        var placeholdersMapping = ReplaceChartsWithPlaceholders(presentation);
        presentation.Save("Chart.pdf");

        using (var pdf = PdfDocument.Load("Chart.pdf"))
        {
            ReplacePlaceholdersWithCharts(pdf, placeholdersMapping);
            pdf.Save();
        }
    }

    static readonly string PlaceholderNameFormat = "GemBox_Chart_Placeholder_{0}";
    static readonly Regex PlaceholderNameRegex = new Regex("GemBox_Chart_Placeholder_\\d+");
    static readonly MemoryStream PlaceholderImage = new MemoryStream(File.ReadAllBytes("placeholder.png"));

    static Dictionary<string, MemoryStream> ReplaceChartsWithPlaceholders(PresentationDocument presentation)
    {
        var placeholdersMapping = new Dictionary<string, MemoryStream>();
        int counter = 0;

        foreach (var slide in presentation.Slides)
        {
            foreach (var frame in slide.Content.Drawings.All()
                .OfType<GraphicFrame>()
                .Where(f => f.Chart != null)
                .ToList())
            {
                var layout = frame.Layout;

                // Replace PowerPoint chart with placeholder image that has specific title.
                var placeholder = slide.Content.AddPicture(PictureContentType.Png, PlaceholderImage,
                    layout.Left, layout.Top, layout.Width, layout.Height);
                string placeholderName = string.Format(PlaceholderNameFormat, ++counter);
                placeholder.AlternativeText.Title = placeholderName;
                frame.Parent.Drawings.Remove(frame);

                // Retrieve Excel chart and export it as PDF.
                var excelChart = (ExcelChart)frame.Chart.ExcelChart;
                excelChart.Position.Width = layout.Width;
                excelChart.Position.Height = layout.Height;
                var chartAsPdfStream = new MemoryStream();
                excelChart.Format().Save(chartAsPdfStream, GemBox.Spreadsheet.SaveOptions.PdfDefault);

                // Map PDF that contains Excel chart to placeholder name.
                placeholdersMapping.Add(placeholderName, chartAsPdfStream);
            }
        }

        return placeholdersMapping;
    }

    static void ReplacePlaceholdersWithCharts(PdfDocument pdfDocument, Dictionary<string, MemoryStream> placeholdersMapping)
    {
        foreach (var page in pdfDocument.Pages)
        {
            // Find placeholders by searching for images with specific title.
            var placeholders = FindPlaceholders(page);

            foreach (var placeholder in placeholders)
            {
                if (!placeholdersMapping.TryGetValue(placeholder.Key, out MemoryStream chartAsPdfStream))
                    continue;

                PdfImageContent image = placeholder.Value.Item1;
                PdfQuad bounds = placeholder.Value.Item2;

                // Replace placeholder image with PDF that contains Excel chart.
                using (var excelDocument = PdfDocument.Load(chartAsPdfStream))
                {
                    var form = excelDocument.Pages[0].ConvertToForm(pdfDocument);
                    var formContentGroup = page.Content.Elements.AddGroup();
                    var formContent = formContentGroup.Elements.AddForm(form);
                    formContent.Transform = PdfMatrix.CreateTranslation(bounds.Left, bounds.Bottom);
                }

                image.Collection.Remove(image);
            }
        }
    }

    static Dictionary<string, Tuple<PdfImageContent, PdfQuad>> FindPlaceholders(PdfPage page)
    {
        var placeholders = new Dictionary<string, Tuple<PdfImageContent, PdfQuad>>();
        var enumerator = page.Content.Elements.All(page.Transform).GetEnumerator();
        while (enumerator.MoveNext())
        {
            var element = enumerator.Current;
            if (element.ElementType != PdfContentElementType.Image)
                continue;

            var imageElement = (PdfImageContent)element;
            var metadata = imageElement.Image.Metadata?.Value;
            if (metadata == null)
                continue;

            var match = PlaceholderNameRegex.Match(metadata);
            if (!match.Success)
                continue;

            var bounds = imageElement.Bounds;
            enumerator.Transform.Transform(ref bounds);
            placeholders.Add(match.Value, Tuple.Create(imageElement, bounds));
        }

        return placeholders;
    }
}