Option Explicit

Class LinkedListNode
	Public [Next]
	Public Previous
	Public Value
	Public List
	
	Private Sub Class_Initialize()
		[Next] = Null
		Previous = Null
		Value = Null
		List = Null
	End Sub
End Class

Class LinkedList
	Private []
	Private lngCount
	Private objFirst
	Private objLast
	Private lngId
	
	Private Sub Class_Initialize()
		Randomize 20011228
		lngId = Fix(Rnd() * CLng(2 ^ 31 - 1))
		lngCount = 0
		
		objFirst = Null
		objLast = Null
		
		Set [] = CreateObject("Brackets")
	End Sub
	
	Private Sub Class_Terminate()
		Call Clear()
	End Sub
	
	Public Property Get Count()
		Count = lngCount
	End Property
	
	Public Property Get First()
		If IsObject(objFirst) Then
			Set First = objFirst
		Else
			First = objFirst
		End If
	End Property
	
	Public Property Get Last()
		If IsObject(objLast) Then
			Set Last = objLast
		Else
			Last = objLast
		End If
	End Property
	
	Public Sub AddAfter(objNode, varNew)
		
		If TypeName(varNew) <> "LinkedListNode" Then
			Dim objNew
			Set objNew = New LinkedListNode
			If IsObject(varNew) Then
				Set objNew.Value = varNew
			Else
				objNew.Value = varNew
			End If
			AddAfter objNode, objNew
			Exit Sub
		End If
		
		If lngCount = 0 And IsNull(objNode) Then
			Set objFirst = varNew
			Set objLast = varNew
			varNew.List = lngId
			lngCount = 1
			Exit Sub
		End If
		
		[].Assert Not IsNull(objNode), _
			"LinkedList", "objNode Can't be Empty."
		[].Assert objNode.List = lngId, _
			"LinkedList", "objNode not belong to this LinkedList."
		[].Assert IsNull(varNew.List), _
			"LinkedList", "objNewNode already belong to a LinkedList."
		
		varNew.List = lngId
			
		If IsObject(objNode.Next) Then
			Set varNew.Next = objNode.Next
		Else
			varNew.Next = objNode.Next
		End If
		Set varNew.Previous = objNode
		Set objNode.Next = varNew
		If Not IsNull(varNew.Next) Then
			If IsObject(varNew) Then
				Set varNew.Next.Previous = varNew
			Else
				varNew.Next.Previous = varNew
			End If
		End If
		[].Inc lngCount
		If objNode Is Last() Then
			Set objLast = varNew
		End If
	End Sub
	
	Public Sub AddBefore(objNode, varNew)
		If TypeName(varNew) <> "LinkedListNode" Then
			Dim objNew
			Set objNew = New LinkedListNode
			If IsObject(varNew) Then
				Set objNew.Value = varNew
			Else
				objNew.Value = varNew
			End If
			AddBefore objNode, objNew
			Exit Sub
		End If
		
		If lngCount = 0 And IsNull(objNode) Then
			Set objFirst = varNew
			Set objLast = varNew
			varNew.List = lngId
			lngCount = 1
			Exit Sub
		End If
		
		[].Assert Not IsNull(objNode), _
			"LinkedList", "objNode Can't be Empty."
		[].Assert objNode.List = lngId, _
			"LinkedList", "objNode not belong to this LinkedList."
		[].Assert IsNull(varNew.List), _
			"LinkedList", "objNewNode already belong to a LinkedList."
	
		varNew.List = lngId
			
		If IsObject(objNode.Previous) Then
			Set varNew.Previous = objNode.Previous
		Else
			varNew.Previous = objNode.Previous
		End If
		Set objNode.Previous = varNew
		Set varNew.Next = objNode
		If Not IsNull(varNew.Previous) Then
			If IsObject(varNew) Then
				Set varNew.Previous.Next = varNew
			Else
				varNew.Previous.Next = varNew
			End If
		End If
		[].Inc lngCount
		
		If objNode Is First() Then
			Set objFirst = varNew
		End If
	End Sub
	
	Public Sub AddFirst(varNew)
		AddBefore First(), varNew
	End Sub
	
	Public Sub AddLast(varNew)
		AddAfter Last(), varNew
	End Sub
	
	Public Sub Clear()
		While Not IsNull(First())
			Call RemoveFirst()
		Wend
	End Sub
	
	Public Function Contains(varValue)
		Contains = Not IsNull(Find(varValue))
	End Function
	
	Public Function Find(varValue)
		Dim objTemp
		Find = Null
		If IsObject(First()) Then
			Set objTemp = First()
		Else
			objTemp = First()
		End If
		While Not IsNull(objTemp)
			If objTemp.Value = varValue Then
				Set Find = objTemp
				Exit Function
			End If
			If IsObject(objTemp.Next) Then
				Set objTemp = objTemp.Next
			Else
				objTemp = objTemp.Next
			End If
		Wend
	End Function
	
	Public Function FindLast(varValue)
		Dim objTemp
		FindLast = Null
		If IsObject(Last()) Then
			Set objTemp = Last()
		Else
			objTemp = Last()
		End If
		While Not IsNull(objTemp)
			If objTemp.Value = varValue Then
				Set FindLast = objTemp
				Exit Function
			End If
			If IsObject(objTemp.Previous) Then
				Set objTemp = objTemp.Previous
			Else
				objTemp = objTemp.Previous
			End If
		Wend
	End Function
	
	Public Function Remove(varOld)
		If TypeName(varOld) = "LinkedListNode" Or IsNull(varOld) Then
			[].Assert lngCount >= 1, "LinkedList", _
				"LinkedList is empty."
			[].Assert varOld.List = lngId, "LinkedList", _
				"The Node you want to remove not belong to this LinkedList."
			
			If IsNull(varOld.Next) Then
				If IsObject(varOld.Previous) Then
					Set objLast = varOld.Previous
				Else
					objLast = varOld.Previous
				End If
			Else
				If IsObject(varOld.Previous) Then
					Set varOld.Next.Previous = varOld.Previous
				Else
					varOld.Next.Previous = varOld.Previous
				End If
			End If
			If IsNull(varOld.Previous) Then
				If IsObject(varOld.Next) Then
					Set objFirst = varOld.Next
				Else
					objFirst = varOld.Next
				End If
			Else
				If IsObject(varOld.Next) Then
					Set varOld.Previous.Next = varOld.Next
				Else
					varOld.Previous.Next = varOld.Next
				End If
			End If
			
			[].Dec lngCount
		Else
			If Contains(varOld) Then	
				Call Remove(Find(varOld))
				Remove = True
			Else
				Remove = False
			End If
		End If
	End Function
	
	Public Sub RemoveFirst()
		Call Remove(First())
	End Sub
	
	Public Sub RemoveLast()
		Call Remove(Last())
	End Sub
End Class