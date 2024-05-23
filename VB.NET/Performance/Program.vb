Imports System.Collections.Generic
Imports System.IO
Imports BenchmarkDotNet.Attributes
Imports BenchmarkDotNet.Engines
Imports BenchmarkDotNet.Jobs
Imports BenchmarkDotNet.Running
Imports GemBox.Presentation

<SimpleJob(RuntimeMoniker.Net80)>
<SimpleJob(RuntimeMoniker.Net48)>
Public Class Program

    Private presentation As PresentationDocument
    Private ReadOnly consumer As Consumer = New Consumer()

    Public Shared Sub Main()
        BenchmarkRunner.Run(Of Program)()
    End Sub

    <GlobalSetup>
    Public Sub SetLicense()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using Free version and example exceeds its limitations, use Trial or Time Limited version:
        ' https://www.gemboxsoftware.com/presentation/examples/free-trial-professional/901

        Me.presentation = PresentationDocument.Load("RandomSlides.pptx")
    End Sub

    <Benchmark>
    Public Function Reading() As PresentationDocument
        Return PresentationDocument.Load("RandomSlides.pptx")
    End Function

    <Benchmark>
    Public Sub Writing()
        Using stream = New MemoryStream()
            Me.presentation.Save(stream, New PptxSaveOptions())
        End Using
    End Sub

    <Benchmark>
    Public Sub Iterating()
        Me.LoopThroughAllDrawings().Consume(Me.consumer)
    End Sub

    Public Iterator Function LoopThroughAllDrawings() As IEnumerable(Of Object)
        For Each slide In Me.presentation.Slides
            For Each drawing In slide.Content.Drawings.All()
                Yield drawing
            Next
        Next
    End Function

End Class