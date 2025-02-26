Sub ResizeAllImages()
    Dim pic As InlineShape
    For Each pic In ActiveDocument.InlineShapes
        With pic
            ' Optionally lock the aspect ratio to true so only width is adjusted.
            .LockAspectRatio = msoTrue
            ' Set the width (change 200 to your desired width in points)
            .Width = 200
        End With
    Next pic
End Sub
