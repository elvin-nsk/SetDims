VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FindMyPeersInCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Implements IFindMyPeers

'===============================================================================

Private Type typeThis
  AllShapes As ShapeRange
  ShapeToCompare As Shape
End Type

Private This As typeThis

'===============================================================================

'фабрика
Public Function Create(ShapeToCompare As Shape, AllShapes As ShapeRange) As IFindMyPeers
  With New FindMyPeersInCol
    .AllShapes = AllShapes
    .ShapeToCompare = ShapeToCompare
    Set Create = .Self
  End With
End Function

Friend Property Let AllShapes(Value As ShapeRange)
  Set This.AllShapes = Value
End Property

Friend Property Let ShapeToCompare(Value As Shape)
  Set This.ShapeToCompare = Value
End Property

Friend Property Get Self() As IFindMyPeers
  Set Self = Me
End Property

'===============================================================================

Private Function IFindMyPeers_GetNeighborNext() As Shape
  Dim SourceRange As ShapeRange
  Set SourceRange = IFindMyPeers_GetPeers
  If SourceRange.Count = 0 Then Exit Function
  Dim DestRange As ShapeRange
  Set DestRange = CreateShapeRange
  Dim Shape As Shape
  For Each Shape In SourceRange
    If Shape.BottomY > This.ShapeToCompare.TopY Then _
      DestRange.Add Shape
  Next Shape
  If DestRange.Count = 0 Then Exit Function
  DestRange.Sort "@shape1.Left < @shape2.Left"
  Set IFindMyPeers_GetNeighborNext = DestRange.FirstShape
End Function

Private Function IFindMyPeers_GetNeighborPrev() As Shape
  Dim SourceRange As ShapeRange
  Set SourceRange = IFindMyPeers_GetPeers
  If SourceRange.Count = 0 Then Exit Function
  Dim DestRange As ShapeRange
  Set DestRange = CreateShapeRange
  Dim Shape As Shape
  For Each Shape In SourceRange
    If Shape.TopY < This.ShapeToCompare.BottomY Then _
      DestRange.Add Shape
  Next Shape
  If DestRange.Count = 0 Then Exit Function
  DestRange.Sort "@shape1.Left < @shape2.Left"
  Set IFindMyPeers_GetNeighborPrev = DestRange.LastShape
End Function

Private Function IFindMyPeers_GetPeers() As ShapeRange
  Dim Shape As Shape
  Set IFindMyPeers_GetPeers = CreateShapeRange
  For Each Shape In This.AllShapes
    If (Shape.RightX <= This.ShapeToCompare.RightX And _
       Shape.RightX >= This.ShapeToCompare.LeftX) Or _
       (Shape.LeftX <= This.ShapeToCompare.RightX And _
       Shape.LeftX >= This.ShapeToCompare.LeftX) Or _
       (Shape.RightX >= This.ShapeToCompare.RightX And _
       Shape.LeftX <= This.ShapeToCompare.LeftX) Or _
       (Shape.RightX <= This.ShapeToCompare.RightX And _
       Shape.LeftX >= This.ShapeToCompare.LeftX) Then _
      IFindMyPeers_GetPeers.Add Shape
  Next Shape
End Function
