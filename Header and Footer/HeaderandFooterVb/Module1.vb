Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        ' Create New master slide.
        Dim master As MasterSlide = presentation.MasterSlides.AddNew()
        master.Content.AddPlaceholder(PlaceholderType.Date)
        master.Content.AddPlaceholder(PlaceholderType.SlideNumber)

        ' Set "DateTime" And "SlideNumber" placeholders visible on slides.
        master.HeaderFooter.IsDateTimeEnabled = True
        master.HeaderFooter.IsSlideNumberEnabled = True

        ' Create New slides; will inherit "DateTime" And "SlideNumber" placeholders from master slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.VerticalTitleAndText)
        slide = presentation.Slides.AddNew(SlideLayoutType.TwoObjects)
        slide = presentation.Slides.AddNew(SlideLayoutType.TwoObjectsAndText)

        presentation.Save("Header and Footer.pptx")

    End Sub

End Module