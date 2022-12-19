Imports System
Imports System.Diagnostics
Imports System.Linq
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If example exceeds Free version limitations then continue as trial version: 
        ' https://www.gemboxsoftware.com/Presentation/help/html/Evaluation_and_Licensing.htm
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Console.WriteLine("Performance example:")
        Console.WriteLine()

        Dim stopwatch = New Stopwatch()
        stopwatch.Start()

        Dim presentation = PresentationDocument.Load("Template.pptx", LoadOptions.Pptx)

        Console.WriteLine("Load file (seconds): " & stopwatch.Elapsed.TotalSeconds)

        stopwatch.Reset()
        stopwatch.Start()

        Dim numberOfShapes As Integer = 0
        Dim numberOfParagraphs As Integer = 0

        For Each slide In presentation.Slides
            For Each shape In slide.Content.Drawings.OfType(Of Shape)

                For Each paragraph In shape.Text.Paragraphs
                    numberOfParagraphs += 1
                Next

                numberOfShapes += 1
            Next
        Next

        Console.WriteLine("Iterate through " & numberOfShapes & " shapes and " & numberOfParagraphs & " paragraphs (seconds): " & stopwatch.Elapsed.TotalSeconds)

        stopwatch.Reset()
        stopwatch.Start()

        presentation.Save("Report.pptx")

        Console.WriteLine("Save file (seconds): " & stopwatch.Elapsed.TotalSeconds)
    End Sub
End Module