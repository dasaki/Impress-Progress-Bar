
Sub InsertProgressBar

Dim oDoc As Object
oDoc = ThisComponent

' Check if this is an Impress document
If thisComponent.supportsService("com.sun.star.drawing.GenericDrawingDocument") Then

	Dim oDrawPage As Object
	Dim Point As New com.sun.star.awt.Point
	Dim Size As New com.sun.star.awt.Size
	Dim Factor As Double
	Dim Percent As Integer

	For n = 0 To oDoc.DrawPages.Count-1
		oDrawPage = oDoc.DrawPages(n)
		Factor = oDrawPage.Width / oDoc.DrawPages.Count
		Size.Width = (n + 1) * Factor
		Size.Height = oDrawPage.Height * 0.05
		Point.x = 0
		Point.y = oDrawPage.Height - Size.Height
		
		m = oDrawPage.getCount-1
		Do While (m >= 0)
			mShape = oDrawPage.getByIndex(m)
			If (InStr(mShape.Name, "Progress Bar Macro") <> 0) Then
				oDrawPage.remove(mShape)
			End If
			m = m-1
		Loop
	
		
		RectangleShape = oDoc.createInstance("com.sun.star.drawing.RectangleShape")
		RectangleShape.CharWeight = com.sun.star.awt.FontWeight.BOLD
		RectangleShape.CharFontName = "Arial"
		RectangleShape.Size = Size
		RectangleShape.Position = Point
	
		RectangleShape.Name = "Progress Bar Macro " + (n+1)
		Percent = 100 * (n + 1) / oDoc.DrawPages.Count
		oDrawPage.add(RectangleShape)
		' RectangleShape.String = "" + oDrawPage.Number + " / " + Doc.DrawPages.Count + "  (%" + Percent + ")"
		RectangleShape.String = "%" + Percent
	Next n
End If


End Sub
