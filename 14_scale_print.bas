Option Explicit
 
Sub scalePrint()
 
  With Sheet1
    .PageSetup.PaperSize = xlPaperA4
    .PageSetup.Orientation = xlPortrait
    .PageSetup.Zoom = False
    .PageSetup.FitToPagesWide = 1
    .PageSetup.FitToPagesTall = 1
  End With
 
End Sub
