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
        Dim table As Table = slide.Content.AddTable(5, 5, 20, 12, LengthUnit.Centimeter)

        ' Format table with no-style grid.
        table.Format.Style = presentation.TableStyles.GetOrAdd(
            TableStyleName.NoStyleTableGrid)

        Dim columnCount As Integer = 4
        Dim rowCount As Integer = 10

        For i As Integer = 0 To columnCount - 1
            ' Create new table column.
            Dim column As TableColumn = table.Columns.AddNew(
                Length.From(5, LengthUnit.Centimeter))
        Next

        For i As Integer = 0 To rowCount - 1
            ' Create new table row.
            Dim row As TableRow = table.Rows.AddNew(
                Length.From(1.2, LengthUnit.Centimeter))

            For j As Integer = 0 To columnCount - 1
                ' Create new table cell.
                Dim cell As TableCell = row.Cells.AddNew()

                ' Set table cell text.
                Dim run As TextRun = cell.Text.AddParagraph().AddRun(
                    String.Format(Nothing, "Cell {0}-{1}", i + 1, j + 1))
            Next
        Next

        presentation.Save("Simple Table.pptx")

    End Sub

End Module