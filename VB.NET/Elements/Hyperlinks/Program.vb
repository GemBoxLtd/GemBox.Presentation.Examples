Imports GemBox.Presentation
Imports GemBox.Presentation.Media
Imports System.IO

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Create first presentation slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create first shape.
        Dim shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 2, 2, 8, 3, LengthUnit.Centimeter)

        ' Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Create a paragraph.
        Dim paragraph = shape.Text.AddParagraph()

        ' Add and format paragraph plain text.
        Dim run = paragraph.AddRun("Powered by ")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Add And format paragraph hyperlink text.
        run = paragraph.AddRun("GemBox")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed))
        run.Format.Action.Click.Set(ActionType.HyperlinkToWebPage, "http://www.gemboxsoftware.com/")

        ' Create second presentation slide.
        slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create first shape.
        shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 2, 2, 4, 3, LengthUnit.Centimeter)

        ' Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Set text run.
        run = shape.Text.AddParagraph().AddRun("Play Sound")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed))

        ' Set "play sound" action on created shape.
        Dim action = shape.Action.Click
        Using stream As Stream = File.OpenRead("Applause.wav")
            action.PlaySound(stream, "applause.wav")
        End Using

        action.Set(ActionType.None)
        action.Highlight = True

        ' Create second shape.
        shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 7, 2, 4, 3, LengthUnit.Centimeter)

        ' Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Set text run.
        run = shape.Text.AddParagraph().AddRun("Hyperlink To Previous Slide")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed))

        ' Set "hyperlink to previous slide" action.
        shape.Action.Click.Set(ActionType.HyperlinkToPreviousSlide)

        ' Create third shape.
        shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 12, 2, 4, 3, LengthUnit.Centimeter)

        ' Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Set text run.
        run = shape.Text.AddParagraph().AddRun("Hyperlink To Next Slide")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed))

        ' Set "hyperlink to next slide" action.
        shape.Action.Click.Set(ActionType.HyperlinkToNextSlide)

        ' Create forth shape.
        shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 17, 2, 4, 3, LengthUnit.Centimeter)

        ' Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Set text run.
        run = shape.Text.AddParagraph().AddRun("Hyperlink To Web Page")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed))

        ' Set "hyperlink to web page" action.
        shape.Action.Click.Set(ActionType.HyperlinkToWebPage, "http://www.gemboxsoftware.com/")

        ' Create third presentation slide.
        slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        presentation.Save("Hyperlinks.pptx")
    End Sub
End Module
