VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  SomeValue As Boolean
End Type

Private This As typeThis

'===============================================================================

'фабрика
Public Function Create() As CreateDim
  If Not Me Is CreateDim Then Exit Function
  With New CreateDim
    Set Create = .Self
  End With
End Function

Friend Property Get Self() As CreateDim
  Set Self = Me
End Property

'===============================================================================

Public Sub AddToStack()
  'Shape1 As Shape, Shape2 As Shape, StepIndex As Long, StepLength As Double
End Sub

Public Function DrawSingleBetween(Shape1 As Shape, _
                                  Shape2 As Shape, _
                                  Direction As cdrDirection, _
                                  TextPositionAbsolute As Double, _
                                  Optional Layer As Layer)
  Dim Point1 As SnapPoint, Point2 As SnapPoint
  Select Case Direction
    Case cdrLeft
      Set Point1 = Shape1.SnapPointsOfType(cdrSnapPointBBox).Item(1)
      Set Point2 = Shape2.SnapPointsOfType(cdrSnapPointBBox).Item(7)
    Case cdrRight
      Set Point1 = Shape1.SnapPointsOfType(cdrSnapPointBBox).Item(3)
      Set Point2 = Shape2.SnapPointsOfType(cdrSnapPointBBox).Item(9)
    Case cdrUp
      Set Point1 = Shape1.SnapPointsOfType(cdrSnapPointBBox).Item(3)
      Set Point2 = Shape2.SnapPointsOfType(cdrSnapPointBBox).Item(1)
    Case cdrDown
      Set Point1 = Shape1.SnapPointsOfType(cdrSnapPointBBox).Item(9)
      Set Point2 = Shape2.SnapPointsOfType(cdrSnapPointBBox).Item(7)
  End Select
  Me.DrawSingleTwoPoint Point1, Point2, Direction, TextPositionAbsolute, Layer
End Function

Public Function DrawSingleOver(Shape As Shape, _
                               Direction As cdrDirection, _
                               TextPositionAbsolute As Double, _
                               Optional Layer As Layer)
  Dim Point1 As SnapPoint, Point2 As SnapPoint
  With Shape.SnapPointsOfType(cdrSnapPointBBox)
    Select Case Direction
      Case cdrLeft
        Set Point1 = .Item(1)
        Set Point2 = .Item(7)
      Case cdrRight
        Set Point1 = .Item(3)
        Set Point2 = .Item(9)
      Case cdrUp
        Set Point1 = .Item(1)
        Set Point2 = .Item(3)
      Case cdrDown
        Set Point1 = .Item(7)
        Set Point2 = .Item(9)
    End Select
  End With
  Me.DrawSingleTwoPoint Point1, Point2, Direction, TextPositionAbsolute, Layer
End Function

Public Function DrawSingleTwoPoint(Point1 As SnapPoint, _
                                   Point2 As SnapPoint, _
                                   Direction As cdrDirection, _
                                   TextPositionAbsolute As Double, _
                                   Optional Layer As Layer)
  If Direction = cdrUp Or Direction = cdrDown Then
    Set DrawSingleTwoPoint = _
      VerifyLayer(Layer).CreateLinearDimension(cdrDimensionHorizontal, Point1, Point2, TextY:=TextPositionAbsolute)
  Else
    Set DrawSingleTwoPoint = _
      VerifyLayer(Layer).CreateLinearDimension(cdrDimensionVertical, Point1, Point2, TextX:=TextPositionAbsolute)
  End If
End Function

Public Function DrawStack()

End Function

'===============================================================================

Private Function VerifyLayer(Layer As Layer) As Layer
  If Layer Is Nothing Then
    Set VerifyLayer = ActiveLayer
  Else
    Set VerifyLayer = Layer
    If Not VerifyLayer Is ActiveLayer Then VerifyLayer.Activate
  End If
End Function
