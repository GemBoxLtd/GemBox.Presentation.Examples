Imports GemBox.Presentation
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Program

    Sub Main()
        Example1()
        Example2()
        Example3()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your GemBox.Presentation serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

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
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

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
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

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

End Module
