Sub Auto_Open()
  ' [F1] Kill F1 key
  Application.OnKey "{F1}", ""

  ' [Insert] Insert new row
  Application.OnKey "{INSERT}", "MyRowInsert"

  ' [Ctrl-G / Ctrl-Shift-G] Group / Un-Group
  Application.OnKey "^g", "MyGroupObj"
  Application.OnKey "^+g", "MyUnGroupObj"

  ' [Ctrl-; / Ctrl-- / Ctrl-0] Zoom up / Zoom back / Zoom 100%
  Application.OnKey "^;", "MyZoomUp"
  Application.OnKey "^-", "MyZoomBack"
  Application.OnKey "^0", "MyZoomDefault"

  ' Close initial workbook
  ThisWorkbook.Close
End Sub


Private Sub MyRowInsert()
  Application.SendKeys "%ir"
End Sub

Private Sub MyGroupObj()
  On Error Resume Next
  If TypeOf Selection Is Range Then Exit Sub

  If Selection.ShapeRange.Count > 1 Then
    Selection.ShapeRange.Group.Select
  Else
    Call MyUnGroupObj
  End If
End Sub
Private Sub MyUnGroupObj()
  On Error Resume Next
  If TypeOf Selection Is Range Then Exit Sub

  If Selection.ShapeRange.Count > 0 Then
    Selection.ShapeRange.Ungroup.Select
  End If
End Sub

Private Sub MyZoomUp()
  Const max = 400
  Const i = 5

  If ActiveWindow.zoom > max - i Then
    ActiveWindow.zoom = max
  Else
    ActiveWindow.zoom = ActiveWindow.zoom + i
  End If
End Sub
Private Sub MyZoomBack()
  Const min = 10
  Const i = 5

  If ActiveWindow.zoom < min + i Then
    ActiveWindow.zoom = min
  Else
    ActiveWindow.zoom = ActiveWindow.zoom - i
  End If
End Sub
Private Sub MyZoomDefault()
  ActiveWindow.zoom = 100
End Sub

