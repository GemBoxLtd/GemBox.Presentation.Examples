Imports System.IO
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Create New presentation slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create first picture from resource data.
        Dim picture As Picture = Nothing
        Using stream As Stream = File.OpenRead("Dices.png")
            picture = slide.Content.AddPicture(PictureContentType.Png, stream, 2, 2, 6, 5, LengthUnit.Centimeter)
        End Using

        ' Create "rounded rectangle" shape.
        Dim shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 10, 2, 8, 5, LengthUnit.Centimeter)

        ' Fill shape with picture content.
        Dim fillFormat = shape.Format.Fill.SetPicture(picture.Fill.Data.Content)

        ' Set the offset of the edges of the stretched picture fill.
        fillFormat.StretchLeft = 0.1
        fillFormat.StretchRight = 0.4
        fillFormat.StretchTop = 0.1
        fillFormat.StretchBottom = 0.4

        ' Get shape outline format.
        Dim lineFormat = shape.Format.Outline

        ' Set shape red outline.
        lineFormat.Fill.SetSolid(Color.FromName(ColorName.Red))
        lineFormat.Width = Length.From(0.2, LengthUnit.Centimeter)

        presentation.Save("Pictures.pptx")
    End Sub
End Module