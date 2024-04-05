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
		[].Set Item, Get_(objRoot, varKey)
	End Property
	
	Private varTemp
	Private Function SetTemp(varValue)
		[].Set varTemp, varValue
	End Function
	
	Private Function Get_(objNode, varKey)
		If TypeName(objNode) = "Nothing" Then
			[].Set Get_, Empty
			Exit Function
		End If
		
		If varKey < objNode.Key Then
			[].Set Get_, Get_(objNode.Left, varKey)
		ElseIf varKey > objNode.Key Then
			[].Set Get_, Get_(objNode.Right, varKey)
		Else
			[].Set Get_, objNode.Value
		End If
	End Function
	
	Public Property Set Item(varKey, objValue)
		Set objRoot = Add_(objRoot, varKey, objValue, True)
	End Property
	
	Public Property Let Item(varKey, varValue)
		Set objRoot = Add_(objRoot, varKey, varValue, True)
	End Property
	
	Public Sub Add(varKey, varValue)
		Set objRoot = Add_(objRoot, varKey, varValue, False)
	End Sub
	
	Private Function Add_(objNode, varKey, varValue, boolAllowCover)
		If TypeName(objNode) = "Nothing" Then
			Set Add_ = NewNode(varKey, varValue)
			Exit Function
		End If
		
		If varKey < objNode.Key Then
			Set objNode.Left = _
				Add_(objNode.Left, varKey, varValue, boolAllowCover)
			Update objNode
		ElseIf varKey > objNode.Key Then
			Set objNode.Right = _
				Add_(objNode.Right, varKey, varValue, boolAllowCover)
			Update objNode
		Else
			[].Assert boolAllowCover, "SortedList", "Key exists."
			SetValue objNode, varValue
		End If
		
		Select Case GetHeight(objNode.Left) - GetHeight(objNode.Right)
			Case 2
				If GetHeight(objNode.Left.Left) > GetHeight(objNode.Left.Right) Then
					' Left Left
					Set Add_ = RightRotate(objNode)
				Else
					' Left Right
					Set objNode.Left = LeftRotate(objNode.Left)
					Set Add_ = RightRotate(objNode)
				End If
			Case -2
				If GetHeight(objNode.Right.Right) > GetHeight(objNode.Right.Left) Then
					' Right Right
					Set Add_ = LeftRotate(objNode)
				Else
					' Right Left
					Set objNode.right = RightRotate(objNode.Right)
					Set Add_ = LeftRotate(objNode)
				End If
			Case Else
				Set Add_ = objNode
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
		Contains = Contains_(objRoot, varKey)
	End Function
	
	Private Function Contains_(objNode, varKey)
		If TypeName(objNode) = "Nothing" Then
			[].Set Contains_, False
			Exit Function
		End If
		
		If varKey < objNode.Key Then
			[].Set Contains_, Contains_(objNode.Left, varKey)
		ElseIf varKey > objNode.Key Then
			[].Set Contains_, Contains_(objNode.Right, varKey)
		Else
			[].Set Contains_, True
		End If
	End Function
	
	Public Function ContainsKey(varKey)
		ContainsKey = Contains(varKey)
	End Function
	
	Public Function ContainsValue(varValue)
		ContainsValue = ContainsValue_(objRoot, varValue)
	End Function
	
	Private Function ContainsValue_(objNode, varValue)
		If TypeName(objNode) = "Nothing" Then
			ContainsValue_ = False
			Exit Function
		End If
		
		If objNode.Value = varValue Then
			ContainsValue_ = True
		Else
			ContainsValue_ = _
				ContainsValue_(objNode.Left, varValue) Or _
				ContainsValue_(objNode.Right, varValue)
		End If
	End Function
	
	Public Function GetByIndex(lngIndex)
		[].Assert Count() > lngIndex And lngIndex >= 0, _
			"SortedList", "Invaild index."
		
		GetByIndex = GetByIndex_(objRoot, lngIndex)
	End Function
	
	Private Function GetByIndex_(objNode, lngIndex)
		If lngIndex - GetSize(objNode.Left) < 0 Then
			[].Set GetByIndex_, _
				GetByIndex_(objNode.Left, lngIndex)
		ElseIf lngIndex - GetSize(objNode.Left) > 0 Then
			[].Set GetByIndex_, _
				GetByIndex_(objNode.Right, lngIndex - GetSize(objNode.Left))
		Else
			[].Set GetByIndex_, objNode.Value
		End If
	End Function
	
	Public Function GetKey(lngIndex)
		[].Assert Count() > lngIndex And lngIndex >= 0, _
			"SortedList", "Invaild index."
		
		GetByIndex = GetKey_(objRoot, lngIndex)
	End Function
	
	Private Function GetKey_(objNode, lngIndex)
		'[].Assert TypeName(objNode) <> "Nothing", _
		'	"SortedList", "Invaild index."
		
		If lngIndex - GetSize(objNode.Left) < 0 Then
			[].Set GetKey_, _
				GetKey_(objNode.Left, lngIndex)
		ElseIf lngIndex - GetSize(objNode.Left) > 0 Then
			[].Set GetKey_, _
				GetKey_(objNode.Right, lngIndex - GetSize(objNode.Left))
		Else
			[].Set GetKey_, objNode.Key
		End If
	End Function
	
	Public Function IndexOfKey(varKey)
		IndexOfKey = IndexOfKey_(objRoot, varKey)
	End Function
	
	Private Function IndexOfKey_(objNode, varKey)
		If TypeName(objNode) = "Nothing" Then
			IndexOfKey_ = -1
			Exit Function
		End If
		
		If varKey < objNode.Key Then
			IndexOfKey_ = _
				IndexOfKey_(objNode.Left, varKey)
		ElseIf varKey > objNode.Key Then
			SetTemp IndexOfKey_(objNode.Left, varKey)
			If varTemp = -1 Then
				IndexOfKey_ = -1
			Else
				IndexOfKey_ = varTemp + GetSize(objNode.Left)
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