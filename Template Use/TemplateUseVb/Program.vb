Imports System.Linq
Imports GemBox.Presentation
Imports GemBox.Presentation.Tables

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Template.pptx")

        ' Retrieve first slide.
        Dim slide = presentation.Slides(0)

        ' Retrieve "Title" placeholder And set shape text.
        Dim shape = slide.Content.Drawings.OfType(Of Shape).Where(Function(item) item.Placeholder IsNot Nothing AndAlso item.Placeholder.PlaceholderType = PlaceholderType.CenteredTitle).First()
        shape.Text.Paragraphs(0).AddRun("ACME Corp - 4th Quarter Financial Results")

        ' Retrieve second slide.
        slide = presentation.Slides(1)

        ' Retrieve "Title" placeholder And set shape text.
        shape = slide.Content.Drawings.OfType(Of Shape).Where(Function(item) item.Placeholder IsNot Nothing AndAlso item.Placeholder.PlaceholderType = PlaceholderType.Title).First()
        shape.Text.Paragraphs(0).AddRun("4th Quarter Summary")

        ' Retrieve "Content" placeholder.
        shape = slide.Content.Drawings.OfType(Of Shape).Where(Function(item) item.Placeholder IsNot Nothing AndAlso item.Placeholder.PlaceholderType = PlaceholderType.Content).First()

        ' Set list text.
        shape.Text.Paragraphs(0).Elements.Clear()
        shape.Text.Paragraphs(0).AddRun("3 new products/services in Research and Development.")

        shape.Text.Paragraphs(1).Elements.Clear()
        shape.Text.Paragraphs(1).AddRun("Rollout planned for new division.")

        shape.Text.Paragraphs(2).Elements.Clear()
        shape.Text.Paragraphs(2).AddRun("Campaigns targeting new markets.")

        ' Retrieve third slide.
        slide = presentation.Slides(2)

        ' Retrieve "Title" placeholder And set shape text.
        shape = slide.Content.Drawings.OfType(Of Shape).Where(Function(item) item.Placeholder IsNot Nothing AndAlso item.Placeholder.PlaceholderType = PlaceholderType.Title).First()
        shape.Text.Paragraphs(0).AddRun("4th Quarter Financial Highlights")

        ' Retrieve a table.
        Dim table = slide.Content.Drawings.OfType(Of GraphicFrame).Where(Function(item) item.Table IsNot Nothing).Select(Function(item) item.Table).First()

        ' Fill table data.
        table.Rows(1).Cells(1).Text.Paragraphs(0).Elements.OfType(Of TextRun).First().Text = "$14.2M"
        table.Rows(1).Cells(2).Text.Paragraphs(0).Elements.OfType(Of TextRun).First().Text = "(0.5%)"

        table.Rows(2).Cells(1).Text.Paragraphs(0).Elements.OfType(Of TextRun).First().Text = "$1.6M"
        table.Rows(2).Cells(2).Text.Paragraphs(0).Elements.OfType(Of TextRun).First().Text = "0.7%"

        table.Rows(3).Cells(1).Text.Paragraphs(0).Elements.OfType(Of TextRun).First().Text = "$12.5M"
        table.Rows(3).Cells(2).Text.Paragraphs(0).Elements.OfType(Of TextRun).First().Text = "0.3%"

        table.Rows(4).Cells(1).Text.Paragraphs(0).Elements.OfType(Of TextRun).First().Text = "$2.3M"
        table.Rows(4).Cells(2).Text.Paragraphs(0).Elements.OfType(Of TextRun).First().Text = "(0.2%)"

        presentation.Save("Template Use.pptx")
    End Sub
End Module