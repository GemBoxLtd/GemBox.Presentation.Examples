Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1
        Dim presentation = PresentationDocument.Load("Reading.pptx")

        ' In order to achieve the conversion of a loaded PowerPoint file to PDF,
        ' we just need to save a PresentationDocument object to desired 
        ' output file format.
        presentation.Save("Convert.pdf")
    End Sub

    Sub Example2
        Dim conformanceLevel As PdfConformanceLevel = PdfConformanceLevel.PdfA1a

        ' Load PowerPoint file.
        Dim presentation = PresentationDocument.Load("Reading.pptx")

        ' Create PDF save options.
        Dim options As New PdfSaveOptions() With
        {
            .ConformanceLevel = conformanceLevel
        }

        ' Save to PDF file.
        presentation.Save("Output.pdf", options)
    End Sub
End Module