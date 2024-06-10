import os
import win32com.client as COM

# Create ComHelper object.
comHelper = COM.Dispatch("GemBox.Presentation.ComHelper")

# If using the Professional version, put your serial key below.
comHelper.ComSetLicense("FREE-LIMITED-KEY")

# Read input presentation.
presentation = comHelper.Load(os.getcwd() + "\\Input.pptx")

# Get first slide.
slide = comHelper.GetCollectionItem(presentation.Slides, 0)

# Remove first drawing.
comHelper.RemoveCollectionItemAt(slide.Content.Drawings, 0)

# Get master slide.
masterSlide = comHelper.GetCollectionItem(presentation.MasterSlides, 0)

# Get layout slide.
layoutSlide = comHelper.GetCollectionItem(masterSlide.LayoutSlides, 0)

# Create new slide.
slide = comHelper.AddNewSlide(presentation.Slides, layoutSlide)

# Create new shape.
shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 5, 5, 12, 6)

# Set shape fill to light blue color.
shape.Format.Fill.SetSolid(comHelper.CreateColor(91, 155, 213))

# Create new paragraph with text.
run = shape.Text.AddParagraph().AddRun("This is a new text box on a new slide.")

# Set text fill to white color.
run.Format.Fill.SetSolid(comHelper.CreateColor(255, 255, 255))

# Write output presentation.
presentation.Save(s.getcwd()  + "\\Output.pptx")