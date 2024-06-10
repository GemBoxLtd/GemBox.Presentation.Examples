<%
  ' Create ComHelper object.
  Set comHelper = CreateObject("GemBox.Presentation.ComHelper")
  
  ' If using the Professional version, put your serial key below.
  comHelper.ComSetLicense("FREE-LIMITED-KEY")
  
  ' Read input presentation.
  Set presentation = comHelper.Load(Server.MapPath(".") & "\Input.pptx")
  
  ' Get first slide.
  Set slide = comHelper.GetCollectionItem(presentation.Slides, 0)
  
  ' Remove first drawing.
  comHelper.RemoveCollectionItemAt(slide.Content.Drawings, 0)
  
  ' Get master slide.
  Set masterSlide = comHelper.GetCollectionItem(presentation.MasterSlides, 0)
  
  ' Get layout slide.
  Set layoutSlide = comHelper.GetCollectionItem(masterSlide.LayoutSlides, 0)
  
  ' Create new slide.
  slide = comHelper.AddNewSlide(presentation.Slides, layoutSlide)
  
  ' Create new shape.
  Set shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 5, 5, 12, 6)
  
  ' Set shape fill to light blue color.
  shape.Format.Fill.SetSolid(comHelper.CreateColor(91, 155, 213))
  
  ' Create new paragraph with text.
  Set run = shape.Text.AddParagraph().AddRun("This is a new text box on a new slide.")
  
  ' Set text fill to white color.
  run.Format.Fill.SetSolid(comHelper.CreateColor(255, 255, 255))
  
  ' Write output presentation.
  presentation.Save(Server.MapPath(".") & "\Output.pptx")
%>