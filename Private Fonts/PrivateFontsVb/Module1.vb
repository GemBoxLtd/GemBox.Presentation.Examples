Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        Dim pathToResources As String = "Resources"

        ' Sets the base directory path where component looks for fonts.
        FontSettings.FontsBaseDirectory = pathToResources

        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim textBox As TextBox = slide.Content.AddTextBox(
            ShapeGeometryType.Rectangle, 2, 2, 8, 8, LengthUnit.Centimeter)

        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        Dim run As TextRun = textBox.AddParagraph().AddRun(
            "Shows how to use private fonts with GemBox.Presentation component.")

        run.Format.Font = "Almonte Snow"
        run.Format.Size = Length.From(16, LengthUnit.Point)
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black))

        presentation.Save("Private Fonts.pdf")

    End Sub

End Module