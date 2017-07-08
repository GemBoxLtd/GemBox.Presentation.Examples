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

        int columnCount = 4;
        int rowCount = 5;

        // Create new table.
        Table table = slide.Content.AddTable(1, 1, 20, 4, LengthUnit.Centimeter);

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
                Length.From(8, LengthUnit.Millimeter));

            for (int j = 0; j < columnCount; j++)
            {
                // Create new table cell.
                row.Cells.AddNew().Text.AddParagraph().AddRun(
                    string.Format(null, "Cell {0}-{1}", i + 1, j + 1));
            }
        }

        // Set table style.
        table.Format.Style = presentation.TableStyles.GetOrAdd(
            TableStyleName.MediumStyle2Accent2);

        // Set table style options.
        table.Format.StyleOptions = TableStyleOptions.FirstRow
            | TableStyleOptions.LastRow
            | TableStyleOptions.BandedRows;

        // Create new table style.
        TableStyle myStyle = presentation.TableStyles.Create("My Table Styles");

        // Get "WholeTable" part style.
        TablePartStyle partStyle = myStyle[TablePartStyleType.WholeTable];

        // Set fill format.
        partStyle.Fill.SetSolid(Color.FromName(ColorName.LightGray));

        // Get table border style.
        TablePartBorderStyle borderStyle = partStyle.Borders;

        // Get "InsideHorizontal" border format.
        LineFormat border = borderStyle[TableCellBorderType.InsideHorizontal];

        // Set border line format.
        border.Fill.SetSolid(Color.FromName(ColorName.DarkGray));
        border.Width = Length.From(2, LengthUnit.Millimeter);

        // Get "InsideVertical" border format.
        border = borderStyle[TableCellBorderType.InsideVertical];

        // Set border line format.
        border.Fill.SetSolid(Color.FromName(ColorName.DarkGray));
        border.Width = Length.From(2, LengthUnit.Millimeter);

        // Get "FirstRow" part style.
        partStyle = myStyle[TablePartStyleType.FirstRow];

        // Set fill format.
        partStyle.Fill.SetSolid(Color.FromName(ColorName.White));

        // Get table border style.
        borderStyle = partStyle.Borders;

        // Get "Top" border format.
        border = borderStyle[TableCellBorderType.Top];

        // Set border line format.
        border.Fill.SetSolid(Color.FromName(ColorName.Black));
        border.Width = Length.From(2, LengthUnit.Millimeter);

        // Get "Bottom" border format.
        border = borderStyle[TableCellBorderType.Bottom];

        // Set border line format.
        border.Fill.SetSolid(Color.FromName(ColorName.Black));
        border.Width = Length.From(2, LengthUnit.Millimeter);

        // Get table text style.
        TablePartTextStyle textStyle = partStyle.Text;

        // Set text format.
        textStyle.Bold = true;
        textStyle.Color = Color.FromName(ColorName.DarkGray);

        // Get "LastRow" part style.
        partStyle = myStyle[TablePartStyleType.LastRow];

        // Set fill format.
        partStyle.Fill.SetSolid(Color.FromName(ColorName.White));

        // Get table border style.
        borderStyle = partStyle.Borders;

        // Set "Top" border line format.
        borderStyle[TableCellBorderType.Top].Fill.SetSolid(
            Color.FromName(ColorName.Black));

        borderStyle[TableCellBorderType.Top].Width =
            Length.From(2, LengthUnit.Millimeter);

        // Set "Bottom" border line format.
        borderStyle[TableCellBorderType.Bottom].Fill.SetSolid(
            Color.FromName(ColorName.Black));

        borderStyle[TableCellBorderType.Bottom].Width =
            Length.From(2, LengthUnit.Millimeter);

        // Set text format.
        partStyle.Text.Bold = true;
        partStyle.Text.Color = Color.FromName(ColorName.DarkGray);

        // Create new table.
        table = slide.Content.AddTable(1, 8, 20, 4, LengthUnit.Centimeter);

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
                Length.From(8, LengthUnit.Millimeter));

            for (int j = 0; j < columnCount; j++)
            {
                // Create new table cell.
                row.Cells.AddNew().Text.AddParagraph().AddRun(
                    string.Format(null, "Cell {0}-{1}", i + 1, j + 1));
            }
        }

        // Set table style.
        table.Format.Style = myStyle;

        // Set table style options.
        table.Format.StyleOptions = TableStyleOptions.FirstRow
            | TableStyleOptions.LastRow;

        presentation.Save("Table Styles.pptx");
    }
}
