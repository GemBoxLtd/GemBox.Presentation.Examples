Imports System.Linq
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Create New master slide.
        Dim master = presentation.MasterSlides.AddNew()

        ' Create New empty layout slide for existing master slide.
        Dim layout = master.LayoutSlides.AddNew(SlideLayoutType.Custom)

        ' Create title And subtitle placeholders on layout slide.
        layout.Content.AddPlaceholder(PlaceholderType.Title)
        layout.Content.AddPlaceholder(PlaceholderType.Subtitle)

        ' Create New slide; will inherit all placeholders (title And subtitle) from template layout slide.
        Dim slide = presentation.Slides.AddNew(layout)

        ' Retrieve "Title" placeholder.
        Dim shape = slide.Content.Drawings.OfType(Of Shape).Where(Function(item) item.Placeholder IsNot Nothing AndAlso item.Placeholder.PlaceholderType = PlaceholderType.Title).First()

        ' Set shape fill and outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))
        shape.Format.Fill.SetSolid(Color.FromName(ColorName.LightGray))

        ' Set shape text.
        shape.Text.AddParagraph().AddRun("Placeholders, GemBox.Presentation")

        ' Retrieve "Subtitle" placeholder.
        shape = slide.Content.Drawings.OfType(Of Shape).Where(Function(item) item.Placeholder IsNot Nothing AndAlso item.Placeholder.PlaceholderType = PlaceholderType.Subtitle).First()

        ' Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Set shape text.
        shape.Text.AddParagraph().AddRun("Shows how to use placeholders with GemBox.Presentation component.")

        presentation.Save("Placeholders.pptx")
    End Sub
End Module