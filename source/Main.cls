VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  LeftLine As Double
  ShapeRange As ShapeRange
  TopLine As Double
End Type

Private This As typeThis

'===============================================================================

'фабрика
Public Function Create() As Main
  If Not Me Is Main Then Exit Function
  With New Main
    Set Create = .Self
  End With
End Function

Friend Property Get Self() As Main
  Set Self = Me
End Property

'===============================================================================

Public Sub Start()

  With This
    Set .ShapeRange = ActiveSelectionRange
    .TopLine = .ShapeRange.TopY + lib_elvin.AverageDim(.ShapeRange) / 15
    .LeftLine = .ShapeRange.LeftX - lib_elvin.AverageDim(.ShapeRange) / 15
    CreateDimsTop .ShapeRange, .TopLine
    CreateDimsLeft .ShapeRange, .LeftLine
  End With

End Sub

'===============================================================================

Private Function CreateDimsTop(ShapeRange As ShapeRange, PositionY As Double) As ShapeRange
  Dim Shape As Shape, NextShape As Shape
  For Each Shape In ShapeRange
    Set NextShape = FindMyPeersInRow.Create(Shape, ShapeRange).GetNeighborNext
    If Not NextShape Is Nothing Then _
      CreateDim.Create.DrawSingleBetween Shape, NextShape, cdrUp, PositionY
  Next Shape
End Function

Private Function CreateDimsLeft(ShapeRange As ShapeRange, PositionX As Double) As ShapeRange
  Dim Shape As Shape, NextShape As Shape
  For Each Shape In ShapeRange
    Set NextShape = FindMyPeersInCol.Create(Shape, ShapeRange).GetNeighborNext
    If Not NextShape Is Nothing Then _
      CreateDim.Create.DrawSingleBetween Shape, NextShape, cdrLeft, PositionX
  Next Shape
End Function

Private Function SortRow(ShapeRange As ShapeRange) As ShapeRange
  Set SortRow = CreateShapeRange
  SortRow.AddRange ShapeRange
  SortRow.Sort "@shape1.Left < @shape2.Left"
End Function

Private Function SortColumn(ShapeRange As ShapeRange) As ShapeRange
  Set SortColumn = CreateShapeRange
  SortColumn.AddRange ShapeRange
  SortColumn.Sort "@shape1.Bottom < @shape2.Bottom"
End Function

Private Sub Old()
  Const STEPDIV# = 20

  Dim Range As ShapeRange
  Dim BgShape As Shape
  Dim Point1 As SnapPoint, Point2 As SnapPoint
  Dim i&, Step#
    
  If ActiveSelectionRange.Count < 2 Then Exit Sub
  'ActiveDocument.Unit = cdrMillimeter
  
  Set Range = ActiveSelectionRange
  Set BgShape = Range.Shapes(1)
  Step = BgShape.SizeWidth / STEPDIV
  
  For i = 2 To Range.Count
    With Range(i)
    
      'слева
      Set Point1 = .SnapPoints.AddUserSnapPoint(.LeftX, .BottomY)
      Set Point2 = BgShape.SnapPoints.AddUserSnapPoint(BgShape.LeftX, .BottomY)
      ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, Point1, Point2, TextY:=BgShape.BottomY - ((i - 1) * Step)
      
      'справа
      Set Point1 = .SnapPoints.AddUserSnapPoint(.RightX, .TopY)
      Set Point2 = BgShape.SnapPoints.AddUserSnapPoint(BgShape.RightX, .TopY)
      ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, Point1, Point2, TextY:=BgShape.TopY + ((i - 1) * Step)
      
      'сверху
      Set Point1 = .SnapPoints.AddUserSnapPoint(.LeftX, .TopY)
      Set Point2 = BgShape.SnapPoints.AddUserSnapPoint(.LeftX, BgShape.TopY)
      ActiveLayer.CreateLinearDimension cdrDimensionVertical, Point1, Point2, TextX:=BgShape.LeftX - ((i - 1) * Step)
      
      'снизу
      Set Point1 = .SnapPoints.AddUserSnapPoint(.RightX, .BottomY)
      Set Point2 = BgShape.SnapPoints.AddUserSnapPoint(.RightX, BgShape.BottomY)
      ActiveLayer.CreateLinearDimension cdrDimensionVertical, Point1, Point2, TextX:=BgShape.RightX + ((i - 1) * Step)
    
    End With
  Next
End Sub
