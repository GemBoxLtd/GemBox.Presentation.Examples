Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        ' Create New presentation slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Adds a New comment with a New author in the top-left corner of the slide.
        Dim comment As Comment = slide.Comments.Add("GBP", "GemBox.Presentation", "Shows how to use comments with GemBox.Presentation component.")

        ' Change comment position.
        comment.Left = Length.From(2, LengthUnit.Centimeter)
        comment.Top = Length.From(1, LengthUnit.Centimeter)

        ' Adds a New comment with the same author as the previously added comment.
        slide.Comments.Add("Another comment from GemBox.Presentation.")

        presentation.Save("Comments.pptx")

    End Sub

End Module