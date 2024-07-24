Set oStack = CreateObject("Collections").NewStack
msgbox TypeName(oStack)

Function AssertEqual(Expected, Actual)
	If Expected = Actual Then
		WScript.Echo "Pass"
	Else
		WScript.Echo "Fail, expected: " & expected & ", actual: " & actual
	End If
End Function

oStack.Push 1
AssertEqual 1, oStack.Pop

oStack.Push 1
oStack.Push 2
AssertEqual 2, oStack.Pop
AssertEqual 1, oStack.Pop

oStack.Push New RegExp
AssertEqual "[object]", TypeName(oStack.Pop)