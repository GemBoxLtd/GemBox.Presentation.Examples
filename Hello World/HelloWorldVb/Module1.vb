Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim textBox As TextBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 5, 4, LengthUnit.Centimeter)

        Dim paragraph As TextParagraph = textBox.AddParagraph()

        paragraph.AddRun("Hello World!")

        presentation.Save("Hello World.pptx")

    End Sub

End Module