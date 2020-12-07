Imports System.IO
Imports GemBox.Presentation

Module Program

    Sub Main()

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create new presentation.
        Dim presentation As New PresentationDocument()

        ' Create new slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Blank)

        ' Add sample text.
        Dim shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 120, 120, 250, 150)
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Black))
        shape.Format.Fill.SetSolid(Color.FromHsl(0, 0, 240))

        Dim run = shape.Text.AddParagraph().AddRun("Lorem Ipsum")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Red))
        run.Format.Bold = True

        run = shape.Text.AddParagraph().AddRun("Lorem Ipsum")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Green))
        run.Format.Italic = True

        run = shape.Text.AddParagraph().AddRun("Lorem Ipsum")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Blue))
        run.Format.UnderlineStyle = UnderlineStyle.Single

        ' Add sample image.
        Using stream = File.OpenRead("Dices.png")
            slide.Content.AddPicture(PictureContentType.Png, stream, 480, 200, 240, 180)
        End Using

        ' Save presentation in PPTX and PDF format.
        presentation.Save("output.pptx")
        presentation.Save("output.pdf")
    End Sub
End Module