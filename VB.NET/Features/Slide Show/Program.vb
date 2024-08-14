Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Create first slide used in custom slideshows.
        Dim slide1 = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create a text box.
        Dim textBox = slide1.Content.AddTextBox(ShapeGeometryType.RoundedRectangle, 2, 2, 12, 4, LengthUnit.Centimeter)

        ' Set shape fill and outline format.
        textBox.Shape.Format.Fill.SetSolid(Color.FromName(ColorName.BlueViolet))
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Violet))

        ' Create a paragraph with single run element.
        Dim run As TextRun = textBox.AddParagraph().AddRun("Shows how to create and customize slide shows using GemBox.Presentation API.")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.White))
        run.Format.Bold = True

        ' Create other two slides used in custom slideshows.
        Dim slide2 = presentation.Slides.AddNew(SlideLayoutType.Custom)
        Dim slide3 = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Get presentation slide show settings.
        Dim settings = presentation.SlideShow

        ' Create first custom slideshow.
        Dim slideShow1 = settings.CustomShows.AddNew("CustomShow1")
        slideShow1.Slides.Add(slide1)
        slideShow1.Slides.Add(slide2)
        slideShow1.Slides.Add(slide3)

        ' Create first custom slideshow.
        Dim slideShow2 = settings.CustomShows.AddNew("CustomShow2")
        slideShow2.Slides.Add(slide3)
        slideShow2.Slides.Add(slide2)
        slideShow2.Slides.Add(slide1)

        ' Show the slides from the first custom show.
        settings.ShowCustomShowSlides("CustomShow1")

        ' Slides should be manually advanced when presenting.
        settings.AdvanceMode = SlideShowAdvanceMode.Manually

        ' Slide show should loop at the end.
        settings.LoopContinuously = True

        ' Slides should be browsed at a kiosk (full screen).
        settings.ShowType = SlideShowType.Kiosk

        presentation.Save("Slide Show.pptx")
    End Sub
End Module
