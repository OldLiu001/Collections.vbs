Option Explicit

Class Matrix
	Private lngRow, lngColumn
	Private adblValues()

	Private boolReadOnly

	Private lngErrorNumber

	Private objMatrixGenerator, objVectorGenerator, objLinearSystem

	Private Sub Assert(ByVal boolCondition, ByRef strMessage)
		If Not boolCondition Then
			Err.Raise vbObjectError, "Matrix", strMessage
		End If
	End Sub
	
	Private Sub Class_Initialize()
		boolReadOnly = False

		Set objMatrixGenerator = New MatrixGenerator
		Set objVectorGenerator = New VectorGenerator
		Set objLinearSystem = New LinearSystem
	End Sub

	Public Property Let Values(ByRef avarRaw)
		Assert Not boolReadOnly, "Matrix is read-only."
		
		Assert TypeName(avarRaw) = "Vector" Or IsArray(avarRaw), _
			"Input is not a vector or array."

		Rem Turn any Input into Array2D(Number).
		If TypeName(avarRaw) = "Vector" Then
			Rem Input is Vector(Number).

			Rem Assume that the Vector is a row vector.
			ReDim adblValues(0, avarRaw.Length - 1)
			Dim lngIndex
			For lngIndex = 0 To UBound(avarRaw.Values())
				adblValues(0, lngIndex) = avarRaw.Value(lngIndex)
			Next

		ElseIf IsArray(avarRaw) Then
			Rem Input is Array(Number).

			On Error Resume Next
				Call UBound(avarRaw, 1)
				lngErrorNumber = Err.Number
			On Error GoTo 0
			Assert lngErrorNumber = 0, "Input array is empty."
			Assert UBound(avarRaw, 1) > -1, "Input array is empty."

			Dim lngRowIndex
			Dim lngColumnIndex
			Dim varElement
			varElement = GetFirstElement(avarRaw)
			Assert TypeName(varElement) = "Variant()" Or _
				IsNumeric(varElement), _
				"Input array has unexpected structure(s)."
			
			If TypeName(varElement) = "Variant()" Then
				Rem Input is Array(Array(Number)).

				Rem Turning Array(Array(Number)) into Array2d(Number).
				ReDim adblValues(UBound(avarRaw), UBound(varElement))
				For lngRowIndex = 0 To UBound(avarRaw)
					For lngColumnIndex = 0 To UBound(varElement)
						Assert IsArray(avarRaw(lngRowIndex)), _
							"Input array has unexpected structure(s)."
						Assert UBound(varElement) = UBound(avarRaw(lngRowIndex)), _
							"Input array is not rectangular."
						Assert IsNumeric(avarRaw(lngRowIndex)(lngColumnIndex)), _
							"Array contains non-numeric value(s)."

						adblValues(lngRowIndex, lngColumnIndex) = _
							CDbl(avarRaw(lngRowIndex)(lngColumnIndex))
					Next
				Next
			ElseIf IsNumeric(varElement) Then
				Rem Input is Array2D(Number).

				Rem Just copy & check.
				ReDim adblValues(UBound(avarRaw, 1), UBound(avarRaw, 2))
				For lngRowIndex = 0 To UBound(avarRaw)
					For lngColumnIndex = 0 To UBound(avarRaw, 2)
						Assert IsNumeric(avarRaw(lngRowIndex, lngColumnIndex)), _
							"Array contains non-numeric value(s)."
						
						adblValues(lngRowIndex, lngColumnIndex) = _
							CDbl(avarRaw(lngRowIndex, lngColumnIndex))
					Next
				Next
			End If
			
		End If

		lngRow = UBound(adblValues, 1) + 1
		lngColumn = UBound(adblValues, 2) + 1
		boolReadOnly = True
	End Property

	Private Function GetFirstElement(ByRef avarArray)
		Dim varElement
		For Each varElement In avarArray
			GetFirstElement = varElement
			Exit For
		Next
	End Function

	Public Property Get Stringify()
		Dim lngRowIndex
		Dim lngColumnIndex
		Stringify = "[" & vbNewLine
		For lngRowIndex = 0 To UBound(Values, 1)
			Stringify = Stringify & "	[ "
			For lngColumnIndex = 0 To UBound(Values, 2)
				Stringify = Stringify & Value(lngRowIndex, lngColumnIndex) & " "
			Next
			Stringify = Stringify & "]" & vbNewLine
		Next
		Stringify = Stringify & "]"
	End Property

	Public Property Get RowCount()
		RowCount = lngRow
	End Property

	Public Property Get ColumnCount()
		ColumnCount = lngColumn
	End Property
	
	Public Property Get Length()
		Length = RowCount * ColumnCount
	End Property

	Private Function IsInteger(ByRef varValue)
		IsInteger = IsNumeric(varValue) And Fix(varValue) = varValue
	End Function
	
	Public Property Get Row(ByVal lngRowIndex)
		Assert IsInteger(lngRowIndex), "Index must be an integer."
		Assert lngRowIndex < RowCount And lngRowIndex >= 0, _
			"Index out of range."
		
		Dim adblRow
		ReDim adblRow(ColumnCount - 1)
		Dim lngColumnIndex
		For lngColumnIndex = 0 To UBound(Values, 2)
			adblRow(lngColumnIndex) = Value(lngRowIndex, lngColumnIndex)
		Next
		Row = adblRow
	End Property

	Public Property Get RowVector(ByVal lngRowIndex)
		Set RowVector = objVectorGenerator.Init(Row(lngRowIndex))
	End Property

	Public Property Get RowMatrix(ByVal lngRowIndex)
		Set RowMatrix = objMatrixGenerator.Init(Array(Row(lngRowIndex)))
	End Property

	Public Property Get Column(ByVal lngColumnIndex)
		Column = Transpose().Row(lngColumnIndex)
	End Property

	Public Property Get ColumnVector(ByVal lngColumnIndex)
		Set ColumnVector = objVectorGenerator.Init(Column(lngColumnIndex))
	End Property

	Public Property Get ColumnMatrix(ByVal lngColumnIndex)
		Set ColumnMatrix = objMatrixGenerator.Init(Array(Column(lngColumnIndex))).Transpose()
	End Property

	Public Property Get Value(ByVal lngRowIndex, ByVal lngColumnIndex)
		Assert IsInteger(lngRowIndex) And IsInteger(lngColumnIndex), _
			"Index must be an integer."
		Assert lngRowIndex < RowCount And lngRowIndex >= 0 And _
			lngColumnIndex < ColumnCount And lngColumnIndex >= 0, _
			"Index out of range."
		
		Value = adblValues(lngRowIndex, lngColumnIndex)
	End Property

	Public Property Get Values()
		Values = adblValues
	End Property

	Public Function Transpose()
		Dim adblTransposed()
		Dim lngRowIndex
		Dim lngColumnIndex
		ReDim adblTransposed(ColumnCount - 1, RowCount - 1)
		For lngRowIndex = 0 To UBound(Values, 1)
			For lngColumnIndex = 0 To UBound(Values, 2)
				adblTransposed(lngColumnIndex, lngRowIndex) = Value(lngRowIndex, lngColumnIndex)
			Next
		Next
		Set Transpose = objMatrixGenerator.Init(adblTransposed)
	End Function

	Public Function Add(ByVal objAnotherMatrix)
		Assert TypeName(objAnotherMatrix) = "Matrix", "Type mismatch."
		Assert RowCount = objAnotherMatrix.RowCount And ColumnCount = objAnotherMatrix.ColumnCount, _
			"Dimension mismatch."
		
		Dim adblAdded()
		Dim lngRowIndex
		Dim lngColumnIndex
		ReDim adblAdded(RowCount - 1, ColumnCount - 1)
		For lngRowIndex = 0 To UBound(Values, 1)
			For lngColumnIndex = 0 To UBound(Values, 2)
				adblAdded(lngRowIndex, lngColumnIndex) = _
					Value(lngRowIndex, lngColumnIndex) + _
					objAnotherMatrix.Value(lngRowIndex, lngColumnIndex)
			Next
		Next
		Set Add = objMatrixGenerator.Init(adblAdded)
	End Function

	Public Function Negate()
		Dim adblNegated()
		Dim lngRowIndex
		Dim lngColumnIndex
		ReDim adblNegated(RowCount - 1, ColumnCount - 1)
		For lngRowIndex = 0 To UBound(Values, 1)
			For lngColumnIndex = 0 To UBound(Values, 2)
				adblNegated(lngRowIndex, lngColumnIndex) = -Value(lngRowIndex, lngColumnIndex)
			Next
		Next
		Set Negate = objMatrixGenerator.Init(adblNegated)
	End Function

	Public Function Subtract(ByVal objAnotherMatrix)
		Set Subtract = Add(objAnotherMatrix.Negate)
	End Function

	Public Function Multiply(ByVal objAnother)
		Assert IsNumeric(objAnother) Or _
			TypeName(objAnother) = "Vector" Or _
			TypeName(objAnother) = "Matrix", "Type mismatch."
		
		Dim adblMultiplied()
		Dim lngRowIndex, lngColumnIndex
		If IsNumeric(objAnother) Then
			Rem Matrix * Number

			ReDim adblMultiplied(RowCount - 1, ColumnCount - 1)
			For lngRowIndex = 0 To UBound(Values, 1)
				For lngColumnIndex = 0 To UBound(Values, 2)
					adblMultiplied(lngRowIndex, lngColumnIndex) = _
						Value(lngRowIndex, lngColumnIndex) * objAnother
				Next
			Next
			Set Multiply = objMatrixGenerator.Init(adblMultiplied)
		ElseIf TypeName(objAnother) = "Vector" Then
			Rem Matrix * Vector

			Rem Assume Vector is a column matrix.
			Dim adblMultipliedRow()
			Set Multiply = Multiply(objMatrixGenerator.Init(objAnother).Transpose())
		ElseIf TypeName(objAnother) = "Matrix" Then
			Rem Matrix * Matrix

			Assert ColumnCount = objAnother.RowCount, _
				"Dimension mismatch."
			ReDim adblMultiplied(RowCount - 1, objAnother.ColumnCount - 1)
			Dim lngAnotherColumnIndex
			For lngRowIndex = 0 To UBound(Values, 1)
				For lngAnotherColumnIndex = 0 To UBound(objAnother.Values, 2)
					
					adblMultiplied(lngRowIndex, lngAnotherColumnIndex) = _
						RowVector(lngRowIndex).DotProduct( _
						objAnother.ColumnVector(lngAnotherColumnIndex))
				Next
			Next
			Set Multiply = objMatrixGenerator.Init(adblMultiplied)
		End If
	End Function

	Private Function IsZero(ByRef varValue)
		IsZero = Abs(varValue) < 1E-7
	End Function
	
	Public Function Divide(ByVal varAnotherNumber)
		Assert IsNumeric(varAnotherNumber), "Type mismatch."
		Assert Not IsZero(varAnotherNumber), "Division by zero."
		Set Divide = Multiply(1 / varAnotherNumber)
	End Function

	Public Function Append(ByVal objAnotherMatrix)
		Rem Append the matrix to the bottom of the current matrix.

		Assert TypeName(objAnotherMatrix) = "Matrix", "Type mismatch."
		Assert ColumnCount = objAnotherMatrix.ColumnCount, _
			"Dimension mismatch."
		
		Dim adblAppended()
		ReDim adblAppended(RowCount + objAnotherMatrix.RowCount - 1, ColumnCount - 1)
		Dim lngRowIndex
		Dim lngColumnIndex
		For lngRowIndex = 0 To UBound(Values, 1)
			For lngColumnIndex = 0 To UBound(Values, 2)
				adblAppended(lngRowIndex, lngColumnIndex) = Value(lngRowIndex, lngColumnIndex)
			Next
		Next
		For lngRowIndex = 0 To UBound(objAnotherMatrix.Values, 1)
			For lngColumnIndex = 0 To UBound(objAnotherMatrix.Values, 2)
				adblAppended(lngRowIndex + RowCount, lngColumnIndex) = _
					objAnotherMatrix.Value(lngRowIndex, lngColumnIndex)
			Next
		Next
		Set Append = objMatrixGenerator.Init(adblAppended)
	End Function

	Public Function Combine(ByVal objAnotherMatrix)
		Rem Combine the matrix to the right of the current matrix.
		Set Combine = Transpose().Append(objAnotherMatrix.Transpose()).Transpose()
	End Function

	Public Property Get Determinant()
		Assert RowCount = ColumnCount, _
			"Only square matrix has determinant."
		
		If RowCount = 1 Then
			Determinant = Value(0, 0)
		ElseIf RowCount <= 3 Then
			Dim lngColumnIndex
			Determinant = 0
			For lngColumnIndex = 0 To UBound(Values, 2)
				Determinant = Determinant + _
					Value(0, lngColumnIndex) * _
					AlgebraicCofactor(0, lngColumnIndex)
			Next
		Else
			'TODO: gauss elimination
		End If
	End Property

	Public Property Get Rank()
		'TODO 'Rank = 
	End Property

	Public Function RemoveRow(lngRowIndex)
		Rem Remove the specified row from the matrix.
		
		Assert lngRowIndex >= 0 And lngRowIndex < RowCount, _
			"Index out of range."
		
		Dim adblRemoved()
		ReDim adblRemoved(RowCount - 2)
		Dim lngTemporaryRowIndex
		For lngTemporaryRowIndex = 0 To UBound(Values, 1)
			If lngTemporaryRowIndex <> lngRowIndex Then
				adblRemoved(lngTemporaryRowIndex - _
					(Sgn(lngTemporaryRowIndex - lngRowIndex) + 1) / 2) = _
					Row(lngTemporaryRowIndex)
			End If
		Next
		Set RemoveRow = objMatrixGenerator.Init(adblRemoved)
	End Function

	Public Function RemoveColumn(lngColumnIndex)
		Rem Remove the specified column from the matrix.
		Set RemoveColumn = _
			Transpose().RemoveRow(lngColumnIndex).Transpose()
	End Function

	Public Property Get Cofactor(ByVal lngRowIndex, ByVal lngColumnIndex)
		Assert RowCount = ColumnCount, _
			"Only square matrix has cofactor."
		Cofactor = _
			RemoveRow(lngRowIndex).RemoveColumn(lngColumnIndex).Determinant
	End Property

	Public Property Get AlgebraicCofactor(ByVal lngRowIndex, ByVal lngColumnIndex)
		AlgebraicCofactor = _
			(-1) ^ (lngRowIndex + lngColumnIndex) * _
			Cofactor(lngRowIndex, lngColumnIndex)
	End Property
End Class