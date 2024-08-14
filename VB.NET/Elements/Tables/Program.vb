Imports GemBox.Presentation
Imports GemBox.Presentation.Tables

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Create new presentation slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create new table.
        Dim table = slide.Content.AddTable(5, 5, 20, 12, LengthUnit.Centimeter)

        ' Format table with no-style grid.
        table.Format.Style = presentation.TableStyles.GetOrAdd(
            TableStyleName.NoStyleTableGrid)

        Dim columnCount As Integer = 4
        Dim rowCount As Integer = 10

        For i As Integer = 0 To columnCount - 1
            ' Create new table column.
            table.Columns.AddNew(Length.From(5, LengthUnit.Centimeter))
        Next

        For i As Integer = 0 To rowCount - 1

            ' Create new table row.
            Dim row = table.Rows.AddNew(
                Length.From(1.2, LengthUnit.Centimeter))

            For j As Integer = 0 To columnCount - 1

                ' Create new table cell.
                Dim cell = row.Cells.AddNew()

                ' Set table cell text.
                cell.Text.AddParagraph().AddRun(
                    String.Format(Nothing, "Cell {0}-{1}", i + 1, j + 1))
            Next
        Next

        presentation.Save("Simple Table.pptx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Create new presentation slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim columnCount As Integer = 4
        Dim rowCount As Integer = 5

        ' Create new table.
        Dim table = slide.Content.AddTable(1, 1, 20, 4, LengthUnit.Centimeter)

        For i As Integer = 0 To columnCount - 1
            ' Create new table column.
            table.Columns.AddNew(Length.From(5, LengthUnit.Centimeter))
        Next

        For i As Integer = 0 To rowCount - 1
            ' Create new table row.
            Dim row = table.Rows.AddNew(
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
        Dim myStyle = presentation.TableStyles.Create("My Table Styles")

        ' Get "WholeTable" part style.
        Dim partStyle = myStyle(TablePartStyleType.WholeTable)

        ' Set fill format.
        partStyle.Fill.SetSolid(Color.FromName(ColorName.LightGray))

        ' Get table border style.
        Dim borderStyle = partStyle.Borders

        ' Get "InsideHorizontal" border format.
        Dim border = borderStyle(TableCellBorderType.InsideHorizontal)

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
        Dim textStyle = partStyle.Text

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
            table.Columns.AddNew(Length.From(5, LengthUnit.Centimeter))
        Next

        For i As Integer = 0 To rowCount - 1
            ' Create new table row.
            Dim row = table.Rows.AddNew(
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
