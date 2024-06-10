Imports GemBox.Pdf
Imports GemBox.Pdf.Content
Imports GemBox.Presentation
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions

Module Program

    Sub Main()

        Example1()
        Example2()
        Example3()
        Example4()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your GemBox.Presentation serial key below.
        GemBox.Presentation.ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As New PresentationDocument()

        ' Add new PowerPoint presentation slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Add simple PowerPoint presentation title.
        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle,
            116.8, 20, 105, 10, GemBox.Presentation.LengthUnit.Millimeter)

        textBox.AddParagraph().AddRun("New presentation with chart element.")

        ' Create PowerPoint chart and add it to slide.
        Dim chart = slide.Content.AddChart(GemBox.Presentation.ChartType.Bar,
            49.3, 40, 240, 120, GemBox.Presentation.LengthUnit.Millimeter)

        ' Get underlying Excel chart.
        Dim excelChart As ExcelChart = DirectCast(chart.ExcelChart, ExcelChart)
        Dim worksheet As ExcelWorksheet = excelChart.Worksheet

        ' Add data for Excel chart.
        worksheet.Cells("A1").Value = "Name"
        worksheet.Cells("A2").Value = "John Doe"
        worksheet.Cells("A3").Value = "Fred Nurk"
        worksheet.Cells("A4").Value = "Hans Meier"
        worksheet.Cells("A5").Value = "Ivan Horvat"

        worksheet.Cells("B1").Value = "Salary"
        worksheet.Cells("B2").Value = 3600
        worksheet.Cells("B3").Value = 2580
        worksheet.Cells("B4").Value = 3200
        worksheet.Cells("B5").Value = 4100

        ' Select data.
        excelChart.SelectData(worksheet.Cells.GetSubrange("A1:B5"), True)

        presentation.Save("Created Chart.pdf")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your GemBox.Presentation serial key below.
        GemBox.Presentation.ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load input file and save it in selected output format
        Dim presentation = PresentationDocument.Load("Chart.pptx")

        ' Get PowerPoint chart.
        Dim chart = DirectCast(presentation.Slides(0).Content.Drawings(0), GraphicFrame).Chart

        ' Get underlying Excel chart and cast it as LineChart.
        Dim lineChart = DirectCast(chart.ExcelChart, LineChart)

        ' Get underlying Excel sheet and add new cell values.
        Dim sheet = lineChart.Worksheet
        sheet.Cells("D1").Value = "Series 3"
        sheet.Cells("D2").Value = 8.6
        sheet.Cells("D3").Value = 5
        sheet.Cells("D4").Value = 7
        sheet.Cells("D5").Value = 9

        ' Add new line series to the LineChart.
        lineChart.Series.Add(sheet.Cells("D1").StringValue, "Sheet1!D2:D5")

        presentation.Save("Updated Chart.pptx")
    End Sub

    Sub Example3()
        ' If using the Professional version, put your GemBox.Presentation serial key below.
        GemBox.Presentation.ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As New PresentationDocument()
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)
        Dim chart = slide.Content.AddChart(GemBox.Presentation.ChartType.Column,
            5, 5, 15, 10, GemBox.Presentation.LengthUnit.Centimeter)

        ' Get underlying Excel chart.
        Dim columnChart = DirectCast(chart.ExcelChart, ColumnChart)

        ' Set chart's category labels from array.
        columnChart.SetCategoryLabels(New String() {"Columns 1", "Columns 2", "Columns 3"})

        ' Add chart's series from arrays.
        columnChart.Series.Add("Values 1", New Double() {3.4, 1.1, 3.7})
        columnChart.Series.Add("Values 2", New Double() {4.4, 3.9, 3.5})
        columnChart.Series.Add("Values 3", New Double() {2.9, 4.1, 1.9})

        presentation.Save("Created Chart from Array.pptx")
    End Sub

    Sub Example4()
        ' If using the Professional versions, put your serial keys below.
        GemBox.Presentation.ComponentInfo.SetLicense("FREE-LIMITED-KEY")
        GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY")
        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Chart.pptx")
        Dim placeholdersMapping = ReplaceChartsWithPlaceholders(presentation)
        presentation.Save("Chart.pdf")

        Using pdf = PdfDocument.Load("Chart.pdf")
            ReplacePlaceholdersWithCharts(pdf, placeholdersMapping)
            pdf.Save()
        End Using
    End Sub

    ReadOnly PlaceholderNameFormat As String = "GemBox_Chart_Placeholder_{0}"
    ReadOnly PlaceholderNameRegex As Regex = New Regex("GemBox_Chart_Placeholder_\d+")
    ReadOnly PlaceholderImage As MemoryStream = New MemoryStream(File.ReadAllBytes("placeholder.png"))

    Function ReplaceChartsWithPlaceholders(presentation As PresentationDocument) As Dictionary(Of String, MemoryStream)
        Dim placeholdersMapping = New Dictionary(Of String, MemoryStream)()
        Dim counter As Integer = 0

        For Each slide In presentation.Slides
            For Each frame In slide.Content.Drawings.All() _
                .OfType(Of GraphicFrame)() _
                .Where(Function(f) f.Chart IsNot Nothing) _
                .ToList()
                Dim layout = frame.Layout

                ' Replace PowerPoint chart with placeholder image that has specific title.
                Dim placeholder = slide.Content.AddPicture(PictureContentType.Png, PlaceholderImage,
                    layout.Left, layout.Top, layout.Width, layout.Height)
                counter += 1
                Dim placeholderName As String = String.Format(PlaceholderNameFormat, counter)
                placeholder.AlternativeText.Title = placeholderName
                frame.Parent.Drawings.Remove(frame)

                ' Retrieve Excel chart and export it as PDF.
                Dim excelChart = CType(frame.Chart.ExcelChart, ExcelChart)
                excelChart.Position.Width = layout.Width
                excelChart.Position.Height = layout.Height
                Dim chartAsPdfStream = New MemoryStream()
                excelChart.Format().Save(chartAsPdfStream, GemBox.Spreadsheet.SaveOptions.PdfDefault)

                ' Map PDF that contains Excel chart to placeholder name.
                placeholdersMapping.Add(placeholderName, chartAsPdfStream)
            Next
        Next

        Return placeholdersMapping
    End Function

    Sub ReplacePlaceholdersWithCharts(pdfDocument As PdfDocument, placeholdersMapping As Dictionary(Of String, MemoryStream))
        Dim chartAsPdfStream As MemoryStream = Nothing

        For Each page In pdfDocument.Pages
            ' Find placeholders by searching for images with specific title.
            Dim placeholders = FindPlaceholders(page)

            For Each placeholder In placeholders
                If Not placeholdersMapping.TryGetValue(placeholder.Key, chartAsPdfStream) Then Continue For

                Dim image As PdfImageContent = placeholder.Value.Item1
                Dim bounds As PdfQuad = placeholder.Value.Item2

                ' Replace placeholder image with PDF that contains Excel chart.
                Using excelDocument = pdfDocument.Load(chartAsPdfStream)
                    Dim form = excelDocument.Pages(0).ConvertToForm(pdfDocument)
                    Dim formContentGroup = page.Content.Elements.AddGroup()
                    Dim formContent = formContentGroup.Elements.AddForm(form)
                    formContent.Transform = PdfMatrix.CreateTranslation(bounds.Left, bounds.Bottom)
                End Using

                image.Collection.Remove(image)
            Next
        Next
    End Sub

    Function FindPlaceholders(page As PdfPage) As Dictionary(Of String, Tuple(Of PdfImageContent, PdfQuad))
        Dim placeholders = New Dictionary(Of String, Tuple(Of PdfImageContent, PdfQuad))()
        Dim enumerator = page.Content.Elements.All(page.Transform).GetEnumerator()

        While enumerator.MoveNext()

            Dim element = enumerator.Current
            If element.ElementType <> PdfContentElementType.Image Then Continue While

            Dim imageElement = CType(element, PdfImageContent)
            Dim metadata = imageElement.Image.Metadata?.Value
            If metadata Is Nothing Then Continue While

            Dim match = PlaceholderNameRegex.Match(metadata)
            If Not match.Success Then Continue While

            Dim bounds = imageElement.Bounds
            enumerator.Transform.Transform(bounds)
            placeholders.Add(match.Value, Tuple.Create(imageElement, bounds))

        End While

        Return placeholders
    End Function
End Module
