Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 20, 2, LengthUnit.Centimeter) _
            .AddParagraph() _
            .AddRun("This presentation has been opened in read-only mode, no changes can be made to a slide.")

        ' ModifyProtection class is supported only for PPTX file format.
        Dim protection As ModifyProtection = presentation.ModifyProtection
        protection.SetPassword("1234")

        presentation.Save("PPTX Modify Protection.pptx")

    End Sub

End Module