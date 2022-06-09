# PowerPointMacros

|**Code**|
|--------|
Add Section Name to each Slide

```VB
Sub AddSectionNameToEachSlide()
    Dim oSl As Slide
        ' Make sure there ARE sections
        If ActivePresentation.SectionProperties.Count > 0 Then
        For Each oSl In ActivePresentation.Slides
            Dim output As String
            output = GetSection(oSl)
            Set myDocument = oSl
            myDocument.Shapes.AddTextbox(Orientation:=msoTextOrientationUpward, _
            Left:=10, Top:=350, Width:=200, Height:=500).TextFrame _
            .TextRange.Text = "Section: " + output
        Next
    End If
End Sub

Function GetSection(oSl As Slide) As String
' Returns the name of the section that this slide belongs to.

    With oSl
        Debug.Print .sectionIndex
        GetSection = ActivePresentation.SectionProperties.Name(.sectionIndex)
    End With
End Function
```

Source: https://stackoverflow.com/questions/63545062/adding-section-chapter-label-to-slide-header
