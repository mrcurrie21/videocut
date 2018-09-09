Attribute VB_Name = "Misc"
Sub RoundedRectangle1_Click()
    FileSelect.Show
End Sub
Function SortArrayAtoZ(myArray As Variant)

Dim i As Long
Dim j As Long
Dim Temp

'Sort the Array A-Z
For i = LBound(myArray) To UBound(myArray) - 1
    For j = i + 1 To UBound(myArray)
        If UCase(myArray(i)) > UCase(myArray(j)) Then
            Temp = myArray(j)
            myArray(j) = myArray(i)
            myArray(i) = Temp
        End If
    Next j
Next i

SortArrayAtoZ = myArray

End Function


