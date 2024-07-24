Option Explicit

Class StackNode
	Public [Next]
	Public Value

	Sub Class_Initialize()
		Set [Next] = Nothing
	End Sub
End Class

Class Stack
	Private objDummyHead, lngCount
	Private Sub Class_Initialize()
		Set objDummyHead = New StackNode
		lngCount = 0
	End Sub
	
	'Push Count Clear Clone ToArray Contains Peek Pop

	Public Sub Push(varVal)
		Dim objNode
		Set objNode = New StackNode
		If IsObject(varVal) Then
			Set objNode.Value = varVal
		Else
			objNode.Value = varVal
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
		Dim objRevStack
		Set objRevStack = New Stack

		Dim objNode
		Set objNode = objDummyHead.[Next]
		Do While Not objNode Is Nothing
			objRevStack.Push objNode.Value
			Set objNode = objNode.[Next]
		Loop
		Set Clone = New Stack
		Do While Not objRevStack.Count = 0
			Clone.Push objRevStack.Pop
		Loop
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
		Contains = False
		Do While Not objNode Is Nothing
			Contains = Contains Or TypeName(objNode.Value) = TypeName(Value)
			If Contains Then
				Exit Function
			End If
			
			Set objNode = objNode.[Next]
		Loop
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