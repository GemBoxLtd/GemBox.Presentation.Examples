Imports System.IO
Imports System.IO.Compression
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
        ' Load a PowerPoint file into the PresentationDocument object.
        Dim presentation = PresentationDocument.Load("Input.pptx")

        ' Create image save options.
        Dim imageOptions As New ImageSaveOptions(ImageSaveFormat.Png) With
        {
            .SlideNumber = 0, ' Select the first slide.
            .Width = 1240 ' Set the image width and keep the aspect ratio.
        }

        ' Save the PresentationDocument object to a PNG file.
        presentation.Save("Output.png", imageOptions)
    End Sub

    Sub Example2()
        ' Load a PowerPoint file.
        Dim presentation = PresentationDocument.Load("Input.pptx")

        ' Max integer value indicates that all presentation slides should be saved.
        Dim imageOptions As New ImageSaveOptions(ImageSaveFormat.Tiff) With
        {
            .SlideCount = Integer.MaxValue
        }

        ' Save the TIFF file with multiple frames, each frame represents a single PowerPoint slide.
        presentation.Save("Output.tiff", imageOptions)
    End Sub

    Sub Example3()
        ' Load a PowerPoint file.
        Dim presentation = PresentationDocument.Load("Input.pptx")

        Dim imageOptions As New ImageSaveOptions()

        ' Get PowerPoint pages, one for each slide.
        Dim pages = presentation.GetPaginator().Pages

        ' Create a ZIP file for storing PNG files.
        Using archiveStream = File.OpenWrite("Output.zip")
            Using archive As New ZipArchive(archiveStream, ZipArchiveMode.Create)
                ' Iterate through the PowerPoint pages.
                For pageIndex As Integer = 0 To pages.Count - 1

                    Dim page As PresentationDocumentPage = pages(pageIndex)

                    ' Create a ZIP entry for each slide.
                    Dim entry = archive.CreateEntry($"Slide {pageIndex + 1}.png")

                    ' Save each slide as a PNG image to the ZIP entry.
                    Using imageStream As New MemoryStream()
                        Using entryStream = entry.Open()
                            page.Save(imageStream, imageOptions)
                            imageStream.CopyTo(entryStream)
                        End Using
                    End Using
                Next
            End Using
        End Using
    End Sub

End Module