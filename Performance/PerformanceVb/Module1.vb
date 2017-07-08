Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If sample exceeds Free version limitations then continue as trial version: 
        ' https://www.gemboxsoftware.com/presentation/help/html/Evaluation_and_Licensing.htm
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Console.WriteLine("Performance sample:")
        Console.WriteLine()

        Dim stopwatch As New Stopwatch()
        stopwatch.Start()

        Dim presentation As PresentationDocument = PresentationDocument.Load("Template.pptx", LoadOptions.Pptx)

        Console.WriteLine("Load file (seconds): " & stopwatch.Elapsed.TotalSeconds)

        stopwatch.Reset()
        stopwatch.Start()

        Dim numberOfShapes As Integer = 0
        Dim numberOfParagraphs As Integer = 0

        For Each slide As Slide In presentation.Slides
            For Each shape As Shape In slide.Content.Drawings.OfType(Of Shape)

                For Each paragraph As TextParagraph In shape.Text.Paragraphs
                    numberOfParagraphs += 1
                Next

                numberOfShapes += 1
            Next
        Next

        Console.WriteLine("Iterate through " + numberOfShapes + " shapes and " + numberOfParagraphs + " paragraphs (seconds): " + stopwatch.Elapsed.TotalSeconds)

        stopwatch.Reset()
        stopwatch.Start()

        presentation.Save("Report.pptx")

        Console.WriteLine("Save file (seconds): " & stopwatch.Elapsed.TotalSeconds)

    End Sub

End Module
