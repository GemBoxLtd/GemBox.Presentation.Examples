' Create ComHelper object and set license.
' NOTE: If you're using a Professional version you'll need to put your serial key below.
Set comHelper = CreateObject("GemBox.Presentation.ComHelper")
comHelper.ComSetLicense("FREE-LIMITED-KEY")

' Load input presentation.
Set presentation = comHelper.Load(Server.MapPath(".") & "\ComTemplate.pptx")

' Get first slide in the presentation.
Set slide = comHelper.GetCollectionItem(presentation.Slides, 0)

' Remove first drawing from the first slide.
comHelper.RemoveCollectionItemAt(slide.Content.Drawings, 0)

' Get master slide.
Set masterSlide = comHelper.GetCollectionItem(presentation.MasterSlides, 0)

' Get layout slide.
Set layoutSlide = comHelper.GetCollectionItem(masterSlide.LayoutSlides, 0)

' Add new slide to the presentation.
slide = comHelper.AddNewSlide(presentation.Slides, layoutSlide)

' Add new shape to the new slide.
Set shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 2, 2, 8, 4)

' Set shape fill to solid blue color.
shape.Format.Fill.SetSolid(comHelper.CreateColor(0, 0, 255))

' Add new paragraph with text.
Set run = shape.Text.AddParagraph().AddRun("This example shows how to create a new PowerPoint slide with GemBox.Presentation in COM.")

' Set text fill to solid white color.
run.Format.Fill.SetSolid(comHelper.CreateColor(255, 255, 255))

' Get output path and save presentation as PDF file.
path = Server.MapPath(".") & "\ComExample.pdf"
presentation.Save(path)
Response.Write("Presentation saved as '" & path & "'")