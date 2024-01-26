Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Set the directory path where the component will look for additional font files.
        ' The "." targets the current directory, so besides the installed fonts,
        ' the component will be able to use the fonts within the specified directory.
        FontSettings.FontsBaseDirectory = "."

        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim textBox = slide.Content.AddTextBox(
            ShapeGeometryType.Rectangle, 2, 2, 8, 8, LengthUnit.Centimeter)

        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        Dim run = textBox.AddParagraph().AddRun(
            "Shows how to use private fonts with GemBox.Presentation component.")

        run.Format.Font = "Almonte Snow"
        run.Format.Size = Length.From(16, LengthUnit.Point)
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black))

        presentation.Save("Private Fonts.pdf")
    End Sub
End Module