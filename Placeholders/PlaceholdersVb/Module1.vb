Imports System
Imports System.Linq
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        ' Create New master slide.
        Dim master As MasterSlide = presentation.MasterSlides.AddNew()

        ' Create New empty layout slide for existing master slide.
        Dim layout As LayoutSlide = master.LayoutSlides.AddNew(SlideLayoutType.Custom)

        ' Create title And subtitle placeholders on layout slide.
        layout.Content.AddPlaceholder(PlaceholderType.Title)
        layout.Content.AddPlaceholder(PlaceholderType.Subtitle)

        ' Create New slide; will inherit all placeholders (title And subtitle) from template layout slide.
        Dim slide As Slide = presentation.Slides.AddNew(layout)

        ' Retrieve "Title" placeholder.
        Dim shape As Shape = slide.Content.Drawings.OfType(Of Shape).Where(Function(item) item.Placeholder?.PlaceholderType = PlaceholderType.Title).First()

        ' Set shape fill And outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))
        shape.Format.Fill.SetSolid(Color.FromName(ColorName.LightGray))

        ' Set shape text.
        shape.Text.AddParagraph().AddRun("Placeholders, GemBox.Presentation")

        ' Retrieve "Subtitle" placeholder.
        shape = slide.Content.Drawings.OfType(Of Shape).Where(Function(item) item.Placeholder?.PlaceholderType = PlaceholderType.Subtitle).First()

        ' Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Set shape text.
        shape.Text.AddParagraph().AddRun("Shows how to use placeholders with GemBox.Presentation component.")

        presentation.Save("Placeholders.pptx")

    End Sub

End Module