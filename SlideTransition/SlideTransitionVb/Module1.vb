Imports System
Imports System.IO
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        Dim pathToResources As String = "Resources"

        ' Create new slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create a text box.
        Dim textBox As TextBox = slide.Content.AddTextBox(ShapeGeometryType.RoundedRectangle, 2, 2, 12, 4, LengthUnit.Centimeter)

        ' Set shape fill and outline format.
        textBox.Shape.Format.Fill.SetSolid(Color.FromName(ColorName.OrangeRed))
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Red))

        ' Create a paragraph with single run element.
        Dim run As TextRun = textBox.AddParagraph().AddRun("Shows how to create and customize slide transitions using GemBox.Presentation API.")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.White))
        run.Format.Bold = True

        ' Get slide transition information.
        Dim transition As SlideShowTransition = slide.Transition

        ' Set the transition type.
        transition.TransitionType = TransitionType.Fade

        ' Set the transition effect.
        transition.Effect = TransitionEffect.Smoothly

        ' Slide should advance automatically after 1 second.
        transition.AdvanceOnTime = True
        transition.AdvanceAfterTime = 1000

        ' Slide should advance on mouse click.
        transition.AdvanceOnClick = True

        ' Set the transition speed.
        transition.Speed = TransitionSpeed.Slow

        ' Specify the sound to be played when the transition is activated.
        Using stream As Stream = File.OpenRead(Path.Combine(pathToResources, "Applause.wav"))
            transition.PlaySound(stream)
        End Using

        presentation.Save("SlideTransition.pptx")

    End Sub

End Module