Sub Assert(boolCondition)
    If Not boolCondition Then
        Err.Raise 5, "ArrayList", "Assertion Failed."
    End If
End Sub

Set oArr = CreateObject("Collections").ArrayList

oArr.Add "A"
oArr.Add "B"
oArr.Add "C"
oArr.Add "D"
oArr.Add "E"
Assert oArr.Count = 5

'