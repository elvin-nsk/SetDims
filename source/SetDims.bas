Attribute VB_Name = "SetDims"
Option Explicit

Sub start()

  Const STEPDIV# = 20

  Dim tRange As ShapeRange
  Dim tBgShape As Shape
  Dim tPoint1 As SnapPoint, tPoint2 As SnapPoint
  Dim i&, tStep#
    
  If ActiveSelectionRange.Count < 2 Then Exit Sub
  'ActiveDocument.Unit = cdrMillimeter
  
  Set tRange = ActiveSelectionRange
  Set tBgShape = tRange.Shapes(1)
  tStep = tBgShape.SizeWidth / STEPDIV
  
  For i = 2 To tRange.Count
    With tRange(i)
    
      'слева
      Set tPoint1 = .SnapPoints.AddUserSnapPoint(.LeftX, .BottomY)
      Set tPoint2 = tBgShape.SnapPoints.AddUserSnapPoint(tBgShape.LeftX, .BottomY)
      ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, tPoint1, tPoint2, TextY:=tBgShape.BottomY - ((i - 1) * tStep)
      
      'справа
      Set tPoint1 = .SnapPoints.AddUserSnapPoint(.RightX, .TopY)
      Set tPoint2 = tBgShape.SnapPoints.AddUserSnapPoint(tBgShape.RightX, .TopY)
      ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, tPoint1, tPoint2, TextY:=tBgShape.TopY + ((i - 1) * tStep)
      
      'сверху
      Set tPoint1 = .SnapPoints.AddUserSnapPoint(.LeftX, .TopY)
      Set tPoint2 = tBgShape.SnapPoints.AddUserSnapPoint(.LeftX, tBgShape.TopY)
      ActiveLayer.CreateLinearDimension cdrDimensionVertical, tPoint1, tPoint2, TextX:=tBgShape.LeftX - ((i - 1) * tStep)
      
      'снизу
      Set tPoint1 = .SnapPoints.AddUserSnapPoint(.RightX, .BottomY)
      Set tPoint2 = tBgShape.SnapPoints.AddUserSnapPoint(.RightX, tBgShape.BottomY)
      ActiveLayer.CreateLinearDimension cdrDimensionVertical, tPoint1, tPoint2, TextX:=tBgShape.RightX + ((i - 1) * tStep)
    
    End With
  Next

End Sub
