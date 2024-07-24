Sub Assert(boolCondition)
    If Not boolCondition Then
        Err.Raise 5, "ArrayList", "Assertion Failed."
    End If
End Sub

Set oArr = CreateObject("Collections.ArrayList")

WSH.Echo TypeName(oArr) 'ArrayList
msgbox oArr.Capacity


'Capacity
'Count
'Item[Int32]
'Add(Object)
'Clear()
'Clone()
'Contains(Object):w
'GetRange(Int32, Int32)
'IndexOf(Object) 'com hasn't
'Insert(Int32, Object)
'LastIndexOf(Object) 'com hasn't
'Remove(Object)
'RemoveAt(Int32)
'RemoveRange(Int32, Int32)
'Reverse()
'Sort()
'ToArray()
'TrimToSize()

' Test Capacity >= Count
For i = 0 To 10000
    oArr.Add i
    Assert oArr.Capacity >= oArr.Count
Next

' Test Clone & Count & Clear
Set oArr2 = oArr.Clone
Assert oArr.Count = oArr2.Count
oArr.Clear
Assert oArr.Count = 0

'Test TrimToSize & Capacity
For i = 0 To 10000
    oArr.Add i
Next
Assert oArr.Capacity >= 10000
oArr.TrimToSize
Assert oArr.Capacity = 10000

'Test Default & Item & Add
For i = 0 To 10000
    oArr.Add i
    Assert oArr(i) = i
    oArr(i) = i + 1
    Assert oArr(i) = i + 1
    oArr.Item(i) = i + 2
    Assert oArr(i) = i + 2
Next