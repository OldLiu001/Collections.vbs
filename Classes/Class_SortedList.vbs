Option Explicit

Class SortedListNode
	Public Key, Value
	Public [Left], [Right]
	Public Height, Size
	Private Sub Class_Initialize
		Set [Left] = Nothing
		Set [Right] = Nothing
		Height = 1
		Size = 1
		Key = Null
		Value = Null
	End Sub
End Class

Class SortedList ' AVL Tree
	Private []
	
	'Private objRoot
	Public objRoot
	
	Private Sub Class_Initialize
		Set [] = CreateObject("Brackets")
		Set objRoot = Nothing
	End Sub
	
	Public Property Get Count()
		Count = GetSize(objRoot)
	End Property
	
	Public Property Get Item(varKey)
		[].Set Item, RecursionGet(objRoot, varKey)
	End Property
	
	Private varTemp
	Private Function SetTemp(varValue)
		[].Set varTemp, varValue
	End Function
	
	Private Function RecursionGet(objNode, varKey)
		If TypeName(objNode) = "Nothing" Then
			[].Set RecursionGet, Empty
			Exit Function
		End If
		
		If varKey < objNode.Key Then
			[].Set RecursionGet, RecursionGet(objNode.Left, varKey)
		ElseIf varKey > objNode.Key Then
			[].Set RecursionGet, RecursionGet(objNode.Right, varKey)
		Else
			[].Set RecursionGet, objNode.Value
		End If
	End Function
	
	Public Property Set Item(varKey, objValue)
		Set objRoot = RecursionAdd(objRoot, varKey, objValue, True)
	End Property
	
	Public Property Let Item(varKey, varValue)
		Set objRoot = RecursionAdd(objRoot, varKey, varValue, True)
	End Property
	
	Public Sub Add(varKey, varValue)
		Set objRoot = RecursionAdd(objRoot, varKey, varValue, False)
	End Sub
	
	Private Function RecursionAdd(objNode, varKey, varValue, boolAllowCover)
		If TypeName(objNode) = "Nothing" Then
			Set RecursionAdd = NewNode(varKey, varValue)
			Exit Function
		End If
		
		If varKey < objNode.Key Then
			Set objNode.Left = _
				RecursionAdd(objNode.Left, varKey, varValue, boolAllowCover)
			Update objNode
		ElseIf varKey > objNode.Key Then
			Set objNode.Right = _
				RecursionAdd(objNode.Right, varKey, varValue, boolAllowCover)
			Update objNode
		Else
			[].Assert boolAllowCover, "SortedList", "Key exists."
			SetValue objNode, varValue
		End If
		
		Select Case GetHeight(objNode.Left) - GetHeight(objNode.Right)
			Case 2
				If GetHeight(objNode.Left.Left) > GetHeight(objNode.Left.Right) Then
					' Left Left
					Set RecursionAdd = RightRotate(objNode)
				Else
					' Left Right
					Set objNode.Left = LeftRotate(objNode.Left)
					Set RecursionAdd = RightRotate(objNode)
				End If
			Case -2
				If GetHeight(objNode.Right.Right) > GetHeight(objNode.Right.Left) Then
					' Right Right
					Set RecursionAdd = LeftRotate(objNode)
				Else
					' Right Left
					Set objNode.right = RightRotate(objNode.Right)
					Set RecursionAdd = LeftRotate(objNode)
				End If
			Case Else
				Set RecursionAdd = objNode
		End Select
	End Function
	
	Private Function GetSize(objNode)
		If TypeName(objNode) = "Nothing" Then
			GetSize = 0
		Else
			GetSize = objNode.Size
		End If
	End Function
	
	Private Sub Update(objNode)
		If TypeName(objNode) <> "Nothing" Then
			UpdateSize objNode
			UpdateHeight objNode
		End If
	End Sub
	
	Private Sub UpdateSize(objNode)
		objNode.Size = GetSize(objNode.Left) + GetSize(objNode.Right) + 1
	End Sub
	
	Private Sub UpdateHeight(objNode)
		objNode.Height = [].Max(GetHeight(objNode.Left), GetHeight(objNode.Right)) + 1
	End Sub
	
	Private Function LeftRotate(objNode)
		Set LeftRotate = objNode.Right
		Dim objTemp
		Set objTemp = objNode.Right.Left
		Set objNode.Right.Left = objNode
		Set objNode.Right = objTemp
		
		Update LeftRotate.Left
		Update LeftRotate
	End Function
	
	Private Function RightRotate(objNode)
		Set RightRotate = objNode.Left
		Dim objTemp
		Set objTemp = objNode.Left.Right
		Set objNode.Left.Right = objNode
		Set objNode.Left = objTemp
		
		Update RightRotate.Right
		Update RightRotate
	End Function
	
	Private Function GetHeight(objNode)
		If TypeName(objNode) = "Nothing" Then
			GetHeight = 0
		Else
			wsh.echo TypeName(objNode)
			GetHeight = objNode.Height
		End If
	End Function
	
	Private Function NewNode(varKey, varValue)
		Set NewNode = New SortedListNode
		SetKey NewNode, varKey
		SetValue NewNode, varValue
	End Function
	
	Private Sub SetKey(objNode, varKey)
		If IsObject(varKey) Then
			Set objNode.Key = varKey
		Else
			objNode.Key = varKey
		End If
	End Sub
	
	Private Sub SetValue(objNode, varValue)
		If IsObject(varValue) Then
			Set objNode.Value = varValue
		Else
			objNode.Value = varValue
		End If
	End Sub
	
	Public Sub Clear()
	End Sub
	
	Public Function Clone()
	End Function
	
	Public Function Contains(varKey)
		Contains = RecursionContains(objRoot, varKey)
	End Function
	
	Private Function RecursionContains(objNode, varKey)
		If TypeName(objNode) = "Nothing" Then
			[].Set RecursionContains, False
			Exit Function
		End If
		
		If varKey < objNode.Key Then
			[].Set RecursionContains, RecursionContains(objNode.Left, varKey)
		ElseIf varKey > objNode.Key Then
			[].Set RecursionContains, RecursionContains(objNode.Right, varKey)
		Else
			[].Set RecursionContains, True
		End If
	End Function
	
	Public Function ContainsKey(varKey)
		ContainsKey = Contains(varKey)
	End Function
	
	Public Function ContainsValue(varValue)
		ContainsValue = RecursionContainsValue(objRoot, varValue)
	End Function
	
	Private Function RecursionContainsValue(objNode, varValue)
		If TypeName(objNode) = "Nothing" Then
			RecursionContainsValue = False
			Exit Function
		End If
		
		If objNode.Value = varValue Then
			RecursionContainsValue = True
		Else
			RecursionContainsValue = _
				RecursionContainsValue(objNode.Left, varValue) Or _
				RecursionContainsValue(objNode.Right, varValue)
		End If
	End Function
	
	Public Function GetByIndex(lngIndex)
		[].Assert Count() > lngIndex And lngIndex >= 0, _
			"SortedList", "Invaild index."
		
		GetByIndex = RecursionGetByIndex(objRoot, lngIndex)
	End Function
	
	Private Function RecursionGetByIndex(objNode, lngIndex)
		'[].Assert TypeName(objNode) <> "Nothing", _
		'	"SortedList", "Invaild index."
		
		If lngIndex - GetSize(objNode.Left) < 0 Then
			[].Set RecursionGetByIndex, _
				RecursionGetByIndex(objNode.Left, lngIndex)
		ElseIf lngIndex - GetSize(objNode.Left) > 0 Then
			[].Set RecursionGetByIndex, _
				RecursionGetByIndex(objNode.Right, lngIndex - GetSize(objNode.Left))
		Else
			[].Set RecursionGetByIndex, objNode.Value
		End If
	End Function
	
	Public Function GetKey(lngIndex)
		[].Assert Count() > lngIndex And lngIndex >= 0, _
			"SortedList", "Invaild index."
		
		GetByIndex = RecursionGetKey(objRoot, lngIndex)
	End Function
	
	Private Function RecursionGetKey(objNode, lngIndex)
		'[].Assert TypeName(objNode) <> "Nothing", _
		'	"SortedList", "Invaild index."
		
		If lngIndex - GetSize(objNode.Left) < 0 Then
			[].Set RecursionGetKey, _
				RecursionGetKey(objNode.Left, lngIndex)
		ElseIf lngIndex - GetSize(objNode.Left) > 0 Then
			[].Set RecursionGetKey, _
				RecursionGetKey(objNode.Right, lngIndex - GetSize(objNode.Left))
		Else
			[].Set RecursionGetKey, objNode.Key
		End If
	End Function
	
	Public Function IndexOfKey(varKey)
		IndexOfKey = RecursionIndexOfKey(objRoot, varKey)
	End Function
	
	Private Function RecursionIndexOfKey(objNode, varKey)
		If TypeName(objNode) = "Nothing" Then
			RecursionIndexOfKey = -1
			Exit Function
		End If
		
		If varKey < objNode.Key Then
			RecursionIndexOfKey = _
				RecursionIndexOfKey(objNode.Left, varKey)
		ElseIf varKey > objNode.Key Then
			SetTemp RecursionIndexOfKey(objNode.Left, varKey)
			If varTemp = -1 Then
				RecursionIndexOfKey = -1
			Else
				RecursionIndexOfKey = varTemp + GetSize(objNode.Left)
			End If
		Else
	End Function
	
	Public Function IndexOfValue(varValue)
		'find first
		'-1 not found
	End Function
	
	Public Sub Remove(varKey)
	End Sub
	
	Public Sub RemoveAt(lngIndex)
	End Sub
	
	Public Sub SetByIndex(lngIndex, varValue)
	End Sub
	
	Public Function s()
		'  a
		' / \
		'b   c
		
	End Function
End Class

Dim a
Set a=CreateObject("System.collections.sortedlist")
a.item(3) = 2
a.item(3) = 3
'MsgBox a.contains(4)
'MsgBox TypeName(a.item(4))
'MsgBox a.contains(4)
a.add 1,1
'MsgBox TypeName(a.item(1))
'MsgBox a.contains(3)
'a.add 1,2
'MsgBox a.item(3)
'MsgBox TypeName(a.GetKeyList().Value)

Dim b
Set b=New SortedList
b.Add 1,1
wsh.echo b.objRoot.value
b.Add 0,0
b.Item(0)=2
b.Add -2,-2
wsh.echo "added"
wsh.echo b.objRoot.value, b.objRoot.Left.Value, b.objRoot.Right.Value
wsh.echo b.objroot.key, b.objroot.Left.key, b.objRoot.Right.key
wsh.echo b.objroot.height, b.objroot.Left.height, b.objRoot.Right.height
wsh.echo b.objroot.size, b.objroot.Left.size, b.objRoot.Right.size
wsh.echo b.Item(3)
wsh.echo b.Item(1)
wsh.echo b.Item(0)
wsh.echo b.ContainsValue(-2)
wsh.echo b.ContainsValue(2)
wsh.echo b.ContainsValue(3)