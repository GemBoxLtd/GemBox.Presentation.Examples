Imports System.Linq
Imports System.IO
Imports System.IO.Compression
Imports GemBox.Presentation
Imports GemBox.Presentation.Media

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim presentation = New PresentationDocument

        ' Create New presentation slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create and add audio content.
        Dim audio As AudioContent = Nothing
        Using stream As Stream = File.OpenRead("Applause.wav")
            audio = slide.Content.AddAudio(AudioContentType.Wav, stream, 2, 2, LengthUnit.Centimeter)
        End Using

        ' Set the ending fade durations for the media.
        audio.Fade.End = TimeOffset.From(300, TimeOffsetUnit.Millisecond)

        ' Get the picture associated with this media.
        Dim picture = audio.Picture

        ' Set drawing properties.
        picture.Action.Click.Set(ActionType.PlayMedia)
        picture.Layout.Width = Length.From(7, LengthUnit.Centimeter)
        picture.Layout.Height = Length.From(7, LengthUnit.Centimeter)
        picture.Name = "Applause.wav"

        ' Create and add video content.
        Dim video As VideoContent = Nothing
        Using stream As Stream = File.OpenRead("Wildlife.wmv")
            video = slide.Content.AddVideo("video/x-ms-wmv", stream, 10, 2, 10, 5.6, LengthUnit.Centimeter)
        End Using

        ' Set drawing properties.
        video.Picture.Action.Click.Set(ActionType.PlayMedia)
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

    Sub Example2()
        Dim presentation = PresentationDocument.Load("Input Audio and Video.pptx")
        Dim slide = presentation.Slides(0)

        ' Get audios from first slide.
        Dim audios = slide.Content.Drawings.All() _
            .OfType(Of Picture)() _
            .Where(Function(p) p.Media?.MediaType = MediaType.Audio) _
            .Select(Function(p) TryCast(p.Media, AudioContent))

        ' Get videos from first slide.
        Dim videos = slide.Content.Drawings.All() _
            .OfType(Of Picture)() _
            .Where(Function(p) p.Media?.MediaType = MediaType.Video) _
            .Select(Function(p) TryCast(p.Media, VideoContent))

        ' Create a ZIP file for storing audio and video files.
        Using archiveStream = File.OpenWrite("Output.zip")
            Using archive = New ZipArchive(archiveStream, ZipArchiveMode.Create)

                Dim counter As Integer = 0
                For Each audio In audios
                    counter += 1
                    Dim extension As String = audio.Content.ContentType.Replace("audio/", String.Empty)
                    Dim entry = archive.CreateEntry($"Audio {counter}.{extension}")

                    ' Export audio from PowerPoint file to the ZIP entry.
                    Using entryStream = entry.Open()
                        Using audioStream = audio.Content.Open()
                            audioStream.CopyTo(entryStream)
                        End Using
                    End Using
                Next

                counter = 0
                For Each video In videos
                    counter += 1
                    Dim extension As String = video.Content.ContentType.Replace("video/", String.Empty)
                    Dim entry = archive.CreateEntry($"Video {counter}.{extension}")

                    ' Export video from PowerPoint file to the ZIP entry.
                    Using entryStream = entry.Open()
                        Using videoStream = video.Content.Open()
                            videoStream.CopyTo(entryStream)
                        End Using
                    End Using
                Next
            End Using
        End Using
    End Sub

End Module