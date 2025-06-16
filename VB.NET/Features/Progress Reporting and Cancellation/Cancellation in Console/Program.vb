Imports GemBox.Presentation
Imports System
Imports System.Diagnostics

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create document.
        Dim presentation As New PresentationDocument()
        For i As Integer = 0 To 10000
            Dim slide = presentation.Slides.AddNew()
            Dim textBox = slide.Content.AddTextBox(100, 100, 100, 100)
            textBox.AddParagraph().AddRun(i.ToString())
        Next

        Dim stopwatch = New Stopwatch()
        stopwatch.Start()

        ' Create save options.
        Dim saveOptions = New PptxSaveOptions()
        AddHandler saveOptions.ProgressChanged,
            Sub(eventSender, args)
                ' Cancel operation after five seconds.
                If stopwatch.Elapsed.Seconds >= 5 Then
                    args.CancelOperation()
                End If
            End Sub

        Try
            presentation.Save("Cancellation.pptx", saveOptions)
            Console.WriteLine("Operation fully finished")
        Catch ex As OperationCanceledException
            Console.WriteLine("Operation was cancelled")
        End Try

    End Sub

End Module
