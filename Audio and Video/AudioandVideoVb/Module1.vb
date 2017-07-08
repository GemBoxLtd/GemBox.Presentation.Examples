Imports System
Imports System.IO
Imports GemBox.Presentation
Imports GemBox.Presentation.Media

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        Dim pathToResources As String = "Resources"

        ' Create New presentation slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create and add audio content.
        Dim audio As AudioContent = Nothing
        Using stream As Stream = File.OpenRead(Path.Combine(pathToResources, "Applause.wav"))
            audio = slide.Content.AddAudio(AudioContentType.Wav, stream, 2, 2, LengthUnit.Centimeter)
        End Using

        ' Set the ending fade durations for the media.
        audio.Fade.End = TimeOffset.From(300, TimeOffsetUnit.Millisecond)

        ' Get the picture associated with this media.
        Dim picture As Picture = audio.Picture

        ' Set drawing properties.
        picture.Layout.Width = Length.From(7, LengthUnit.Centimeter)
        picture.Layout.Height = Length.From(7, LengthUnit.Centimeter)
        picture.Name = "Applause.wav"

        ' Create and add video content.
        Dim video As VideoContent = Nothing
        Using stream As Stream = File.OpenRead(Path.Combine(pathToResources, "Wildlife.wmv"))
            video = slide.Content.AddVideo("video/x-ms-wmv", stream, 10, 2, 10, 5.6, LengthUnit.Centimeter)
        End Using

        ' Set drawing properties.
        video.Picture.Name = "Wildlife.wmv"

        ' Set the amount of time to be trimmed from the start And end of the media.
        video.Trim.Start = TimeOffset.From(600, TimeOffsetUnit.Millisecond)
        video.Trim.End = TimeOffset.From(800, TimeOffsetUnit.Millisecond)

        ' Set the starting And ending fade durations for the media.
        video.Fade.Start = TimeOffset.From(100, TimeOffsetUnit.Millisecond)
        video.Fade.End = TimeOffset.From(200, TimeOffsetUnit.Millisecond)

        ' Add video bookmarks.
        video.Bookmarks.Add(TimeOffset.From(1500, TimeOffsetUnit.Millisecond))
        video.Bookmarks.Add(TimeOffset.From(3000, TimeOffsetUnit.Millisecond))

        presentation.Save("Audio and Video.pptx")

    End Sub

End Module