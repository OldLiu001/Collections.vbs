Import "..\Classes\Class_Stack.vbs"
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