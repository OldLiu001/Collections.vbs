Option Explicit

Class StackNode
	Public [Next]
	Public Value

	Sub Class_Initialize()
		Set [Next] = Nothing
	End Sub
End Class

Class Stack
	'Push Count Clear Clone ToArray Contains Peek Pop Equals ToString
	Private pDummyHead, pCount
	Private Sub Class_Initialize()
		Set DummyHead = New StackNode
		Count = 0
	End Sub

	Public Sub Push(Value)
		Dim NewNode
		Set NewNode = New StackNode
		NewNode.Value = Value
		Set NewNode.[Next] = DummyHead.[Next]
		Set DummyHead.[Next] = NewNode
		Count = Count + 1
	End Sub

	Property Get Count
		Count = Count
	End Property

	Public Function Pop()
		Dim Node, Value
		If Count = 0 Then
			Err.Raise 5, "Stack.Pop", "Stack is empty"
		End If
		Set Node = DummyHead.[Next]
		If IsObject(Node.Value) Then
			Set Value = Node.Value
		Else
			Value = Node.Value
		End If
		Set DummyHead.[Next] = Node.[Next]
		Set Node = Nothing
		Count = Count - 1
		If IsObject(Value) Then
			Set Pop = Value
		Else
			Pop = Value
		End If
	End Function
End Class