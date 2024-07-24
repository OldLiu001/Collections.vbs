Class StackNode
	Public [Next]
	Public Value

	Sub Class_Initialize()
		Set [Next] = Nothing
	End Sub
End Class

Class Stack
	'Push Count Clear Clone ToArray Contains Peek Pop
	Private objDummyHead, lngCount
	Private Sub Class_Initialize()
		Set objDummyHead = New StackNode
		lngCount = 0
	End Sub

	Public Sub Push(Value)
		Dim objNode
		Set objNode = New StackNode
		If IsObject(Value) Then
			Set objNode.Value = Value
		Else
			objNode.Value = Value
		End If
		Set objNode.[Next] = objDummyHead.[Next]
		Set objDummyHead.[Next] = objNode
		lngCount = lngCount + 1
	End Sub

	Public Property Get Count
		Count = lngCount
	End Property

	Public Sub Clear()
		Dim objNode
		Do While Not objDummyHead.[Next] Is Nothing
			Set objNode = objDummyHead.[Next]
			Set objDummyHead.[Next] = objNode.[Next]
			Set objNode = Nothing
		Loop
		lngCount = 0
	End Sub

	Public Function Clone()
		Dim objStack, objStack2
		Set objStack = New Stack
		Dim objNode
		Set objNode = objDummyHead.[Next]
		Do While Not objNode Is Nothing
			objStack.Push objNode.Value
			Set objNode = objNode.[Next]
		Loop
		Set objStack2 = New Stack
		Do While Not objStack.Count = 0
			objStack2.Push objStack.Pop
		Loop
		Set Clone = objStack2
	End Function

	Public Function ToArray()
		Dim arrResult()
		ReDim arrResult(lngCount - 1)
		Dim objNode
		Set objNode = objDummyHead.[Next]
		For i = 0 To lngCount - 1
			If IsObject(objNode.Value) Then
				Set arrResult(i) = objNode.Value
			Else
				arrResult(i) = objNode.Value
			End If
			Set objNode = objNode.[Next]
		Next
		ToArray = arrResult
	End Function

	Public Function Contains(Value)
		Dim objNode
		Set objNode = objDummyHead.[Next]
		Do While Not objNode Is Nothing
			If IsObject(objNode.Value) Then
				If objNode.Value Is Value Then
					Contains = True
					Exit Function
				End If
			Else
				
				If objNode.Value = Value Then
					Contains = True
					Exit Function
				End If
			End If
			Set objNode = objNode.[Next]
		Loop
		Contains = False
	End Function

	Public Function Peek()
		If lngCount = 0 Then
			Err.Raise 5, , "Stack is empty"
		End If
		
		Dim objNode
		Set objNode = objDummyHead.[Next]
		If IsObject(objNode.Value) Then
			Set Peek = objNode.Value
		Else
			Peek = objNode.Value
		End If
	End Function

	Public Function Pop()
		If lngCount = 0 Then
			Err.Raise 5, , "Stack is empty"
		End If
		
		Dim objNode
		Set objNode = objDummyHead.[Next]
		Set objDummyHead.[Next] = objNode.[Next]
		lngCount = lngCount - 1
		If IsObject(objNode.Value) Then
			Set Pop = objNode.Value
		Else
			Pop = objNode.Value
		End If
		Set objNode = Nothing
	End Function
End Class
' Import "..\Classes\Class_Stack.vbs"
Sub Import(strFileName)
    With CreateObject("Scripting.FileSystemObject")
        ExecuteGlobal .OpenTextFile( _
            .GetParentFolderName( _
            .GetFile(WScript.ScriptFullName)) & _
            "\" & strFileName).ReadAll
    End With
End Sub
Set oStack = New Stack

Function AssertEqual(Expected, Actual)
	If IsObject(Expected) Then
		If Not IsObject(Actual) Then
			WScript.StdErr.WriteLine "Fail, expected: " & TypeName(Expected) & ", actual: " & TypeName(Actual)
			Exit Function
		End If

		If Expected Is Actual Then
			WScript.Echo "Pass"
		Else
			WScript.StdErr.WriteLine "Fail, expected: " & TypeName(Expected) & ", actual: " & TypeName(Actual)
		End If
	Else
		If IsObject(Actual) Then
			WScript.StdErr.WriteLine "Fail, expected: " & TypeName(Expected) & ", actual: " & TypeName(Actual)
			Exit Function
		End If
		If Expected = Actual Then
			WScript.Echo "Pass"
		Else
			WScript.StdErr.WriteLine "Fail, expected: " & expected & ", actual: " & actual
		End If
	End If
End Function

AssertEqual Array(1, 2, 3), Array(1, 2, 3)
AssertEqual Array(1, 2, 3), Array(1, 2, 4)

' Test Push & Pop
oStack.Push 1
AssertEqual 1, oStack.Pop
oStack.Push 1
oStack.Push 2
AssertEqual 2, oStack.Pop
AssertEqual 1, oStack.Pop
oStack.Push CreateObject("Scripting.Dictionary")
oStack.Push CreateObject("Scripting.FileSystemObject")
AssertEqual "FileSystemObject", TypeName(oStack.Pop)
AssertEqual "Dictionary", TypeName(oStack.Pop)
WScript.Echo "Test Push & Pop passed"

' Test Count & Peek
For i = 1 To 10
	oStack.Push i
	AssertEqual i, oStack.Count
	AssertEqual i, oStack.Peek
Next
WScript.Echo "Test Count & Peek passed"

' Test Clear
For i = 1 To 10000
	oStack.Push i
	oStack.Push CreateObject("Scripting.Dictionary")
	oStack.Push oStack
	oStack.Push New RegExp
	oStack.Push "Hello"
	oStack.Push 3.14
	oStack.Push #1/1/2000#
Next
oStack.Clear
AssertEqual 0, oStack.Count
WScript.Echo "Test Clear passed"

'Test Clone & ToArray & Contains
oStack.Push 1
oStack.Push 2
oStack.Push CreateObject("Scripting.Dictionary")
oStack.Push oStack
oStack.Push New RegExp
oStack.Push "Hello"
oStack.Push 3.14
oStack.Push #1/1/2000#
Set oStack2 = oStack.Clone
AssertEqual oStack.Count, oStack2.Count
arr1 = oStack.ToArray
arr2 = oStack2.ToArray
AssertEqual UBound(arr1), UBound(arr2)
For i = 0 To UBound(arr1)
	WScript.Echo TypeName(arr1(i)), TypeName(arr2(i))
	AssertEqual arr1(i), arr2(i)
Next
N = oStack.Count
For i = 1 To N
	AssertEqual oStack.Contains(arr2(Fix(Rnd * N))), True
	AssertEqual oStack.Pop, oStack2.Pop
Next