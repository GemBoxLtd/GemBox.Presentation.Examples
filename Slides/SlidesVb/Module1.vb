Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        ' Get slide size.
        Dim size As SlideSize = presentation.SlideSize

        ' Set slide size.
        size.SizedFor = SlideSizeType.OnscreenShow16X10
        size.Orientation = Orientation.Landscape
        size.NumberSlidesFrom = 1

        ' Create New master slide.
        Dim master As MasterSlide = presentation.MasterSlides.AddNew()

        ' Create New layout slide for existing master slide.
        Dim layout As LayoutSlide = master.LayoutSlides.AddNew(SlideLayoutType.TitleAndObject)

        ' Create New slide from existing template layout slide.
        Dim slide As Slide = presentation.Slides.AddNew(layout)

        ' If master slide collection is empty, this method will add a new master slide.
        ' If layout slide collection of the last master slide doesn't contain a layout slide with the specified type, 
        ' then a new layout slide with the specified type will be added.
        slide = presentation.Slides.AddNew(SlideLayoutType.TwoTextAndTwoObjects)

        presentation.Save("Slides.pptx")

    End Sub

End Module