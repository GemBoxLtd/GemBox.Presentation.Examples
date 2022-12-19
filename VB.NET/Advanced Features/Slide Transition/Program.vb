Imports System.IO
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Create new slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create a text box.
        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.RoundedRectangle, 2, 2, 12, 4, LengthUnit.Centimeter)

        ' Set shape fill and outline format.
        textBox.Shape.Format.Fill.SetSolid(Color.FromName(ColorName.OrangeRed))
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Red))

        ' Create a paragraph with single run element.
        Dim run = textBox.AddParagraph().AddRun("Shows how to create and customize slide transitions using GemBox.Presentation API.")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.White))
        run.Format.Bold = True

        ' Get slide transition information.
        Dim transition = slide.Transition

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
        Using stream = File.OpenRead("Applause.wav")
            transition.PlaySound(stream)
        End Using

        presentation.Save("Slide Transitions.pptx")
    End Sub
End Module