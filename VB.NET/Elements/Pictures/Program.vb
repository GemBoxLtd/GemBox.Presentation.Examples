Imports GemBox.Presentation
Imports System.IO
Imports System.Linq

Module Program

    Sub Main()
        Example1()
        Example2()
        Example3()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
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

        ' Create second picture from SVG resource.
        Using stream As Stream = File.OpenRead("Graphics1.svg")
            picture = slide.Content.AddPicture(PictureContentType.Svg, stream, 2, 8, 6, 3, LengthUnit.Centimeter)
        End Using

        presentation.Save("Pictures.pptx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("InputPictures.pptx")
        Dim slide = presentation.Slides(0)

        ' Get all pictures from first slide.
        Dim pictures = slide.Content.Drawings.All().OfType(Of Picture)()

        ' Get first picture data.
        Dim picture As Picture = pictures.First()
        Dim pictureContent As PictureContent = picture.Fill.Data

        ' Export picture data to image file.
        Using fileStream = File.Create($"Output.{pictureContent.ContentType}")
            Using pictureStream = pictureContent.Content.Open()
                pictureStream.CopyTo(fileStream)
            End Using
        End Using
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("InputPictures.pptx")
        Dim slide = presentation.Slides(0)

        ' Get all pictures from first slide.
        Dim pictures = slide.Content.Drawings.All().OfType(Of Picture)()

        ' Replace pictures data with image file.
        For Each picture In pictures
            Using fileStream = File.OpenRead("Jellyfish.jpg")
                picture.Fill.SetData(fileStream, PictureContentType.Jpeg)
            End Using
        Next

        presentation.Save("UpdatedPictures.pptx")
    End Sub
End Module
