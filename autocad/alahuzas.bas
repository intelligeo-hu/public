Public Sub Alahuz()
  Dim ss As AcadSelectionSet
  Dim obj As AcadObject
  Dim txt As AcadText
  Dim Str As String
  On Error Resume Next
  ThisDrawing.SelectionSets("ALAHUZAS").Delete
  Set ss = ThisDrawing.SelectionSets.Add("ALAHUZAS")
  ThisDrawing.Utility.Prompt "Válassz aláhúzandó feliratokat" & vbCrLf
  ss.SelectOnScreen
  For Each obj In ss
    If TypeOf obj Is AcadText Then
      Set txt = obj
      Str = txt.TextString
      If UCase(Left(Str, 3)) <> "%%U" Then      
        txt.TextString = "%%u" + txt.TextString
        txt.Update        
      End If    
    End If     
  Next
End Sub

Public Sub Alahuz_ki()
  Dim ss As AcadSelectionSet
  Dim obj As AcadObject
  Dim txt As AcadText
  Dim Str As String
  On Error Resume Next
  ThisDrawing.SelectionSets("ALAHUZAS").Delete
  Set ss = ThisDrawing.SelectionSets.Add("ALAHUZAS")
  ThisDrawing.Utility.Prompt "Aláhúzás ki: " & vbCrLf
  ss.SelectOnScreen
  For Each obj In ss
    If TypeOf obj Is AcadText Then
      Set txt = obj
      Str = txt.TextString
      If UCase(Left(Str, 3)) = "%%U" Then
        txt.TextString = Right(Str, Len(Str) - 3)
        txt.Update
      End If
    End If
  Next
End Sub
