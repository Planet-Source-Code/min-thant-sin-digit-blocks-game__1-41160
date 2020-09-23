Attribute VB_Name = "basOverlap"
Option Explicit

'Check if the block is almost overlap on a square.
Function Overlap(picSquare As PictureBox, picBlock As PictureBox) As Boolean
      Overlap = False
      With picBlock
            If (.Left + .Width >= picSquare.Left + picSquare.Width / 2) And (.Left <= picSquare.Left + picSquare.Width / 2) Then
                  If .Top + .Height >= picSquare.Top + picSquare.Height / 2 Then
                        If .Top <= picSquare.Top + picSquare.Height / 2 Then
                              Overlap = True
                        End If
                  End If
            End If
      End With
End Function

' Check if the two objects intersect, using the IntersectRect API.
Function Collided(obj1 As Object, obj2 As Object) As Boolean
      Dim obj1Rect As RECT, obj2Rect As RECT, DestRect As RECT
      
      ' Copy info into Rect structures
      'obj1
      With obj1Rect
            .Left = obj1.Left
            .Top = obj1.Top
            .Right = .Left + obj1.Width - 1
            .Bottom = .Top + obj1.Height - 1
      End With
      
      'obj2
      With obj2Rect
            .Left = obj2.Left
            .Top = obj2.Top
            .Right = .Left + obj2.Width - 1
            .Bottom = .Top + obj2.Height - 1
      End With
      
      ' IntersectRect will only return 0 (false) if the
      ' two rectangles do NOT intersect.
      Collided = IntersectRect(DestRect, obj1Rect, obj2Rect)
End Function
