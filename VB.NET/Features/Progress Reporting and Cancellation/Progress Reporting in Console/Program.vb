Imports GemBox.Presentation
Imports System

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Console.WriteLine("Creating presentation")

        ' Create large presentation.
        Dim presentation As New PresentationDocument()
        For i As Integer = 0 To 10000
            Dim slide = presentation.Slides.AddNew()
            Dim textBox = slide.Content.AddTextBox(100, 100, 100, 100)
            textBox.AddParagraph().AddRun(i.ToString())
        Next

        ' Create save options.
        Dim saveOptions = New PptxSaveOptions()
        AddHandler saveOptions.ProgressChanged,
            Sub(eventSender, args)
                Console.WriteLine($"Progress changed - {args.ProgressPercentage}%")
            End Sub

        ' Save presentation.
        presentation.Save("presentation.pptx", saveOptions)

    End Sub

End Module
