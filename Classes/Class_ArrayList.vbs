Option Explicit

Function NewArrayList()
	Set NewArrayList = New ArrayList
End Function

Class ArrayList
	Private lngCount
	Private arrData()

	Private Sub Assert(boolCondition, strSource, strMessage)
		If Not boolCondition Then
			Err.Raise 5, strSource, strMessage
		End If
	End Sub

	Private Sub Assign(ByRef varDest, ByVal varSrc)
		If IsObject(varSrc) Then
			Set varDest = varSrc
		Else
			varDest = varSrc
		End If
	End Sub

	Private Sub Inc(ByRef lngValue)
		lngValue = lngValue + 1
	End Sub

	Private Sub Dec(ByRef lngValue)
		lngValue = lngValue - 1
	End Sub

	Private Sub Swap(ByRef varA, ByRef varB)
		Dim varTemp
		Assign varTemp, varA
		Assign varA, varB
		Assign varB, varTemp
	End Sub
	
	Private Sub Class_Initialize()
		ReDim arrData(0)
		lngCount = 0
	End Sub
	
	Public Property Get Capacity()
		Capacity = UBound(arrData) + 1
	End Property
	
	Public Property Get Count()
		Count = lngCount
	End Property
	
	Public Property Get Item(lngIndex)
		Assert lngIndex >= 0 And lngIndex < lngCount, _
			"ArrayList", "Invaild Index. Expect 0 <= lngIndex < lngCount."
		Assign Item, arrData(lngIndex)
	End Property
	
	Public Property Let Item(lngIndex, varElement)
		Assert lngIndex >= 0 And lngIndex < lngCount, _
			"ArrayList", "Invaild Index. Expect 0 <= lngIndex < lngCount."
		Assign arrData(lngIndex), varElement
	End Property
	
	Public Property Set Item(lngIndex, objElement)
		Assert lngIndex >= 0 And lngIndex < lngCount, _
			"ArrayList", "Invaild Index. Expect 0 <= lngIndex < lngCount."
		Assign arrData(lngIndex), objElement
	End Property

	'Default
	Public Property Get Default(lngIndex)
		Default = Item(lngIndex)
	End Property

	Public Property Let Default(lngIndex, varElement)
		Item(lngIndex) = varElement
	End Property

	Public Property Set Default(lngIndex, objElement)
		Set Item(lngIndex) = objElement
	End Property

	
	Public Function Add(varValue)
		Insert lngCount, varValue
		Add = lngCount - 1
	End Function
	
	Public Sub Clear()
		Call Class_Initialize()
	End Sub
	
	Public Function Clone()
		Set Clone = GetRange(0, lngCount)
	End Function
	
	Public Function Contains(varElement)
		Contains = (IndexOf(varElement) <> -1)
	End Function
	
	Public Function GetRange(lngIndex, lngLength)
		Assert lngLength >= 0, _
			"ArrayList", "Expect lngLength >= 0."
		Assert lngIndex >= 0 And lngIndex <= lngCount, _
			"ArrayList", "Invaild Index. Expect 0 <= lngIndex <= lngCount."
		Assert lngIndex + lngLength <= lngCount, _
			"ArrayList", "Invaild Range. Expect lngIndex + lngLength <= lngCount."
		
		Set GetRange = New ArrayList
		
		Dim i
		For i = lngIndex To lngIndex + lngLength - 1
			GetRange.Add arrData(i)
		Next
	End Function
	
	Public Function IndexOf(varElement)
		' If there has many varElement, return smallest index.
		
		IndexOf = -1
		Dim i
		For i = 0 To lngCount - 1
			If varElement = arrData(i) Then
				IndexOf = i
				Exit For
			End If
		Next
	End Function
	
	Public Sub Insert(lngIndex, varElement)
		Assert lngIndex >= 0 And lngIndex <= lngCount, _
			"ArrayList", "Invaild Index. Expect 0 <= lngIndex <= lngCount."
		
		If UBound(arrData) = lngCount - 1 Then
			' arrData full.
			ResizeTo(2 * lngCount)
		End If
		
		Dim i
		For i = lngCount - 1 To lngIndex Step -1
			Assign arrData(i + 1), arrData(i)
		Next
		
		Assign arrData(lngIndex), varElement
		Inc lngCount
	End Sub
	
	Public Function LastIndexOf(varElement)
		' If there has many varElement, return biggest index.
		
		LastIndexOf = -1
		Dim i
		For i = 0 To lngCount - 1
			If varElement = arrData(i) Then
				LastIndexOf = i
			End If
		Next
	End Function
	
	Public Sub Remove(varElement)
		'Remove first matched element.
		
		Dim lngIndex
		lngIndex = IndexOf(varElement)
		If lngIndex <> -1 Then
			RemoveAt lngIndex
		End If
	End Sub
	
	Public Sub RemoveAt(lngIndex)
		Assert lngIndex >= 0 And lngIndex < lngCount, _
			"ArrayList", "Invaild Index. Expect 0 <= Index < lngCount."
		
		Dim i
		For i = lngIndex + 1 To lngCount - 1
			Assign arrData(i - 1), arrData(i)
		Next
		
		Dec lngCount
		
		If Capacity() \ 3 = lngCount - 1 Then
			' arrData is sparse, free some memory.
			ResizeTo(Capacity() \ 2)
		End If
	End Sub
	
	Public Sub RemoveRange(lngIndex, lngLength)
		Assert lngLength >= 0, _
			"ArrayList", "Expect lngLength >= 0."
		Assert lngIndex >= 0 And lngIndex <= lngCount, _
			"ArrayList", "Invaild Index. Expect 0 <= lngIndex <= lngCount."
		Assert lngIndex + lngLength <= lngCount, _
			"ArrayList", "Invaild Range. Expect lngIndex + lngLength <= lngCount."
		
		Dim i
		i = 0
		While i + lngIndex + lngLength < lngCount
			Assign arrData(i + lngIndex), arrData(i + lngIndex + lngLength)
			Inc i
		Wend
		
		lngCount = lngCount - lngLength
	End Sub
	
	Public Sub Reverse()
		Dim i
		
		Dim varTemp
		For i = 1 To lngCount \ 2
			Swap arrData(i - 1), arrData(lngCount - i)
		Next
	End Sub
	
	Public Sub Sort()
		'TODO
	End Sub
	
	Public Function ToArray()
		Dim arrRet
		arrRet = arrData
		ReDim Preserve arrRet(lngCount - 1)
		ToArray = arrRet
	End Function
	
	Public Sub TrimToSize()
		ResizeTo lngCount - 1
	End Sub
	
	Private Sub ResizeTo(lngUpperBound)
		ReDim Preserve arrData(lngUpperBound)
	End Sub
End Class