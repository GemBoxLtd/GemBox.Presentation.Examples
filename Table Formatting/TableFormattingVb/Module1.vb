Imports System
Imports GemBox.Presentation
Imports GemBox.Presentation.Tables

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        ' Create new presentation slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create new table.
        Dim table As Table = slide.Content.AddTable(5, 5, 20, 5, LengthUnit.Centimeter)

        ' Format table with no-style grid.
        table.Format.Style = presentation.TableStyles.GetOrAdd(
            TableStyleName.NoStyleTableGrid)

        table.Format.Fill.SetSolid(Color.FromName(ColorName.Orange))

        table.Columns.AddNew(Length.From(7, LengthUnit.Centimeter))
        table.Columns.AddNew(Length.From(10, LengthUnit.Centimeter))
        table.Columns.AddNew(Length.From(5, LengthUnit.Centimeter))

        Dim row As TableRow = table.Rows.AddNew(Length.From(5, LengthUnit.Centimeter))

        Dim cell As TableCell = row.Cells.AddNew()

        cell.Format.Fill.SetSolid(Color.FromName(ColorName.Red))

        cell.Text.Format.VerticalAlignment = VerticalAlignment.Top

        cell.Text.AddParagraph().AddRun("Cell 1-1")

        cell = row.Cells.AddNew()

        Dim border As LineFormat = cell.Format.DiagonalDownBorderLine

        border.Fill.SetSolid(Color.FromName(ColorName.White))
        border.Width = Length.From(5, LengthUnit.Millimeter)

        border = cell.Format.DiagonalUpBorderLine

        border.Fill.SetSolid(Color.FromName(ColorName.White))
        border.Width = Length.From(5, LengthUnit.Millimeter)

        cell.Text.Format.VerticalAlignment = VerticalAlignment.Middle

        cell.Text.AddParagraph().AddRun("Cell 1-2")

        cell = row.Cells.AddNew()

        cell.Format.Fill.SetSolid(Color.FromName(ColorName.DarkBlue))

        cell.Text.Format.VerticalAlignment = VerticalAlignment.Bottom

        cell.Text.AddParagraph().AddRun("Cell 1-3")

        presentation.Save("Table Formatting.pptx")

    End Sub

End Module