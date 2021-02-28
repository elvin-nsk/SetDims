Attribute VB_Name = "SetDims"
Option Explicit

Private Const RELEASE As Boolean = True

Sub Start()
  
  If RELEASE Then On Error GoTo ErrHandler
  lib_elvin.BoostStart "Расстановка измерений", RELEASE
  Main.Create.Start

CleanExit:
  lib_elvin.BoostFinish
  Exit Sub

ErrHandler:
  MsgBox "Ошибка: " & Err.Description, vbCritical
  Resume CleanExit

End Sub

Private Sub test()
  With ActiveLayer.Shapes
    FindMyPeersInCol.Create(ActiveShape, .All).GetPeers.CreateSelection
  End With
End Sub

Private Sub test2()
  Dim NewRange As ShapeRange
  Set NewRange = CreateShapeRange
  NewRange.AddRange ActiveLayer.Shapes.All
  With NewRange
    .Sort "@shape1.Left > @shape2.Left"
    .Item(1).CreateSelection
  End With
End Sub

Private Sub test3()
  With ActiveLayer.Shapes
    Dim Shape As Shape
    Set Shape = FindMyPeersInCol.Create(ActiveShape, .All).GetNeighborPrev
    If Shape Is Nothing Then Exit Sub
    Shape.CreateSelection
  End With
End Sub

Private Sub test4()
  CreateDim.Create.DrawSingleOver ActiveShape, cdrDown, 0
End Sub

