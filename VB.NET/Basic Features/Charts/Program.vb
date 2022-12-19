Imports System.Linq
Imports GemBox.Presentation
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Program

    Sub Main()

        ' If using the Professional version, put your GemBox.Presentation serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()

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

        ' Load input file and save it in selected output format
        Dim presentation = PresentationDocument.Load("Chart.pptx")

        ' Get PowerPoint chart.
        Dim chart = DirectCast(presentation.Slides(0).Content.Drawings(0), GraphicFrame).Chart

        ' Get underlying Excel chart and cast it as LineChart.
        Dim lineChart = DirectCast(chart.ExcelChart, LineChart)

        ' Add new line series which has doubled values from the first series.
        lineChart.Series.Add("Series 3", lineChart.Series.First() _
            .Values.Cast(Of Double)().Select(Function(val) val * 2))

        presentation.Save("Updated Chart.pptx")

    End Sub

End Module