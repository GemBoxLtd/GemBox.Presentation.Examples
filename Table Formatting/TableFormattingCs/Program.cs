using GemBox.Presentation;
using GemBox.Presentation.Tables;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create new presentation slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create new table.
        var table = slide.Content.AddTable(5, 5, 20, 5, LengthUnit.Centimeter);

        // Format table with no-style grid.
        table.Format.Style = presentation.TableStyles.GetOrAdd(
            TableStyleName.NoStyleTableGrid);

        table.Format.Fill.SetSolid(Color.FromName(ColorName.Orange));

        table.Columns.AddNew(Length.From(7, LengthUnit.Centimeter));
        table.Columns.AddNew(Length.From(10, LengthUnit.Centimeter));
        table.Columns.AddNew(Length.From(5, LengthUnit.Centimeter));

        var row = table.Rows.AddNew(Length.From(5, LengthUnit.Centimeter));

        var cell = row.Cells.AddNew();

        cell.Format.Fill.SetSolid(Color.FromName(ColorName.Red));

        cell.Text.Format.VerticalAlignment = VerticalAlignment.Top;

        cell.Text.AddParagraph().AddRun("Cell 1-1");

        cell = row.Cells.AddNew();

        var border = cell.Format.DiagonalDownBorderLine;

        border.Fill.SetSolid(Color.FromName(ColorName.White));
        border.Width = Length.From(5, LengthUnit.Millimeter);

        border = cell.Format.DiagonalUpBorderLine;

        border.Fill.SetSolid(Color.FromName(ColorName.White));
        border.Width = Length.From(5, LengthUnit.Millimeter);

        cell.Text.Format.VerticalAlignment = VerticalAlignment.Middle;

        cell.Text.AddParagraph().AddRun("Cell 1-2");

        cell = row.Cells.AddNew();

        cell.Format.Fill.SetSolid(Color.FromName(ColorName.DarkBlue));

        cell.Text.Format.VerticalAlignment = VerticalAlignment.Bottom;

        cell.Text.AddParagraph().AddRun("Cell 1-3");

        presentation.Save("Table Formatting.pptx");
    }
}