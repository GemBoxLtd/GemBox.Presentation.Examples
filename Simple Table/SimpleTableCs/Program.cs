using System;
using GemBox.Presentation;
using GemBox.Presentation.Tables;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PresentationDocument presentation = new PresentationDocument();

        // Create new presentation slide.
        Slide slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create new table.
        Table table = slide.Content.AddTable(5, 5, 20, 12, LengthUnit.Centimeter);

        // Format table with no-style grid.
        table.Format.Style = presentation.TableStyles.GetOrAdd(
            TableStyleName.NoStyleTableGrid);

        int columnCount = 4;
        int rowCount = 10;

        for (int i = 0; i < columnCount; i++)
        {
            // Create new table column.
            TableColumn column = table.Columns.AddNew(
                Length.From(5, LengthUnit.Centimeter));
        }

        for (int i = 0; i < rowCount; i++)
        {
            // Create new table row.
            TableRow row = table.Rows.AddNew(
                Length.From(1.2, LengthUnit.Centimeter));

            for (int j = 0; j < columnCount; j++)
            {
                // Create new table cell.
                TableCell cell = row.Cells.AddNew();

                // Set table cell text.
                TextRun run = cell.Text.AddParagraph().AddRun(
                    string.Format(null, "Cell {0}-{1}", i + 1, j + 1));
            }
        }

        presentation.Save("Simple Table.pptx");
    }
}
