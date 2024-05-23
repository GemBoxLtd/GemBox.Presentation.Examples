Imports GemBox.Presentation
Imports System

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("CloneDestination.pptx")

        Dim sourcePresentation = PresentationDocument.Load("CloneSource.pptx")

        ' Use context so that references between 
        ' shapes and slides are maintained between all cloning operations.
        Dim context = CloneContext.Create(sourcePresentation, presentation)

        ' Clone all drawings from the first slide of another presentation 
        ' into the first slide of the current presentation.
        For Each drawing In sourcePresentation.Slides(0).Content.Drawings
            presentation.Slides(0).Content.Drawings.AddClone(drawing, context)
        Next

        ' Establish explicit mapping between slides so that 
        ' hyperlink on the second slide is correctly cloned.
        context.Set(sourcePresentation.Slides(0), presentation.Slides(0))

        ' Clone the second slide from another presentation.
        presentation.Slides.AddClone(sourcePresentation.Slides(1), context)

        presentation.Save("Cloning.pptx")
    End Sub
End Module
