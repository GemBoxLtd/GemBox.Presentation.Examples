Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Get slide size.
        Dim size = presentation.SlideSize

        ' Set slide size.
        size.SizedFor = SlideSizeType.OnscreenShow16X10
        size.Orientation = Orientation.Landscape
        size.NumberSlidesFrom = 1

        ' Create New master slide.
        Dim master = presentation.MasterSlides.AddNew()

        ' Create New layout slide for existing master slide.
        Dim layout = master.LayoutSlides.AddNew(SlideLayoutType.TitleAndObject)

        ' Create New slide from existing template layout slide.
        Dim slide = presentation.Slides.AddNew(layout)

        ' If master slide collection is empty, this method will add a new master slide.
        ' If layout slide collection of the last master slide doesn't contain a layout slide with the specified type, 
        ' then a new layout slide with the specified type will be added.
        slide = presentation.Slides.AddNew(SlideLayoutType.TwoTextAndTwoObjects)

        presentation.Save("Slides.pptx")
    End Sub
End Module