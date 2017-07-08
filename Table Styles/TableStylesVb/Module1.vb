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

        Dim columnCount As Integer = 4
        Dim rowCount As Integer = 5

        ' Create new table.
        Dim table As Table = slide.Content.AddTable(1, 1, 20, 4, LengthUnit.Centimeter)

        For i As Integer = 0 To columnCount - 1
            ' Create new table column.
            Dim column As TableColumn = table.Columns.AddNew(
                Length.From(5, LengthUnit.Centimeter))
        Next

        For i As Integer = 0 To rowCount - 1
            ' Create new table row.
            Dim row As TableRow = table.Rows.AddNew(
                Length.From(8, LengthUnit.Millimeter))

            For j As Integer = 0 To columnCount - 1
                ' Create new table cell.
                row.Cells.AddNew().Text.AddParagraph().AddRun(
                    String.Format(Nothing, "Cell {0}-{1}", i + 1, j + 1))
            Next
        Next

        ' Set table style.
        table.Format.Style = presentation.TableStyles.GetOrAdd(
            TableStyleName.MediumStyle2Accent2)

        ' Set table style options.
        table.Format.StyleOptions = TableStyleOptions.FirstRow Or
            TableStyleOptions.LastRow Or
            TableStyleOptions.BandedRows

        ' Create New table style.
        Dim myStyle As TableStyle = presentation.TableStyles.Create("My Table Styles")

        ' Get "WholeTable" part style.
        Dim partStyle As TablePartStyle = myStyle(TablePartStyleType.WholeTable)

        ' Set fill format.
        partStyle.Fill.SetSolid(Color.FromName(ColorName.LightGray))

        ' Get table border style.
        Dim borderStyle As TablePartBorderStyle = partStyle.Borders

        ' Get "InsideHorizontal" border format.
        Dim border As LineFormat = borderStyle(TableCellBorderType.InsideHorizontal)

        ' Set border line format.
        border.Fill.SetSolid(Color.FromName(ColorName.DarkGray))
        border.Width = Length.From(2, LengthUnit.Millimeter)

        ' Get "InsideVertical" border format.
        border = borderStyle(TableCellBorderType.InsideVertical)

        ' Set border line format.
        border.Fill.SetSolid(Color.FromName(ColorName.DarkGray))
        border.Width = Length.From(2, LengthUnit.Millimeter)

        ' Get "FirstRow" part style.
        partStyle = myStyle(TablePartStyleType.FirstRow)

        ' Set fill format.
        partStyle.Fill.SetSolid(Color.FromName(ColorName.White))

        ' Get table border style.
        borderStyle = partStyle.Borders

        ' Get "Top" border format.
        border = borderStyle(TableCellBorderType.Top)

        ' Set border line format.
        border.Fill.SetSolid(Color.FromName(ColorName.Black))
        border.Width = Length.From(2, LengthUnit.Millimeter)

        ' Get "Bottom" border format.
        border = borderStyle(TableCellBorderType.Bottom)

        ' Set border line format.
        border.Fill.SetSolid(Color.FromName(ColorName.Black))
        border.Width = Length.From(2, LengthUnit.Millimeter)

        ' Get table text style.
        Dim textStyle As TablePartTextStyle = partStyle.Text

        ' Set text format.
        textStyle.Bold = True
        textStyle.Color = Color.FromName(ColorName.DarkGray)

        ' Get "LastRow" part style.
        partStyle = myStyle(TablePartStyleType.LastRow)

        ' Set fill format.
        partStyle.Fill.SetSolid(Color.FromName(ColorName.White))

        ' Get table border style.
        borderStyle = partStyle.Borders

        ' Set "Top" border line format.
        borderStyle(TableCellBorderType.Top).Fill.SetSolid(
            Color.FromName(ColorName.Black))

        borderStyle(TableCellBorderType.Top).Width =
            Length.From(2, LengthUnit.Millimeter)

        ' Set "Bottom" border line format.
        borderStyle(TableCellBorderType.Bottom).Fill.SetSolid(
            Color.FromName(ColorName.Black))

        borderStyle(TableCellBorderType.Bottom).Width =
            Length.From(2, LengthUnit.Millimeter)

        ' Set text format.
        partStyle.Text.Bold = True
        partStyle.Text.Color = Color.FromName(ColorName.DarkGray)

        ' Create new table.
        table = slide.Content.AddTable(1, 8, 20, 4, LengthUnit.Centimeter)

        For i As Integer = 0 To columnCount - 1
            ' Create new table column.
            Dim column As TableColumn = table.Columns.AddNew(
                Length.From(5, LengthUnit.Centimeter))
        Next

        For i As Integer = 0 To rowCount - 1
            ' Create new table row.
            Dim row As TableRow = table.Rows.AddNew(
                Length.From(8, LengthUnit.Millimeter))

            For j As Integer = 0 To columnCount - 1
                ' Create new table cell.
                row.Cells.AddNew().Text.AddParagraph().AddRun(
                    String.Format(Nothing, "Cell {0}-{1}", i + 1, j + 1))
            Next
        Next

        ' Set table style.
        table.Format.Style = myStyle

        ' Set table style options.
        table.Format.StyleOptions = TableStyleOptions.FirstRow Or
            TableStyleOptions.LastRow

        presentation.Save("Table Styles.pptx")

    End Sub

End Module