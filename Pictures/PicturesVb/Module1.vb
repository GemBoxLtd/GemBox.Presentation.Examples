Imports System
Imports System.IO
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        Dim pathToResources As String = "Resources"

        ' Create New presentation slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create first picture from resource data.
        Dim picture As Picture = Nothing
        Using stream As Stream = File.OpenRead(Path.Combine(pathToResources, "Dices.png"))
            picture = slide.Content.AddPicture(PictureContentType.Png, stream, 2, 2, 6, 5, LengthUnit.Centimeter)
        End Using

        ' Create "rounded rectangle" shape.
        Dim shape As Shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 10, 2, 8, 5, LengthUnit.Centimeter)

        ' Fill shape with picture content.
        Dim fillFormat As PictureFillFormat = shape.Format.Fill.SetPicture(picture.Fill.Data.Content)

        ' Set the offset of the edges of the stretched picture fill.
        fillFormat.StretchLeft = 0.1
        fillFormat.StretchRight = 0.4
        fillFormat.StretchTop = 0.1
        fillFormat.StretchBottom = 0.4

        ' Get shape outline format.
        Dim lineFormat As LineFormat = shape.Format.Outline

        ' Set shape red outline.
        lineFormat.Fill.SetSolid(Color.FromName(ColorName.Red))
        lineFormat.Width = Length.From(0.2, LengthUnit.Centimeter)

        presentation.Save("Pictures.pptx")

    End Sub

End Module