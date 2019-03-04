Attribute VB_Name = "Arrays"
Sub DeclaringArrays()

'Declare Array with range 0,1,2,3
Dim MyArray(0 To 3) As Variant

'Declare Array with range 0,1,2,3
Dim MyArray(3) As Variant

'Declare Array with range 1,2,3
Dim MyArray(1 To 3) As Variant

'Declare Array with range 2,3,4
Dim MyArray(2 To 4) As Variant

'DYNAMIC ARRAYS

'Declare Array with Dynamic Range
Dim MyArray() As Variant
    
'Resize Array with range 0,1,2,3,4
ReDim MyArray(0 To 4)   

'ASSIGN VALUES TO AN ARRAY

MyArray(0) = 100
MyArray(1) = 200
MyArray(2) = 300
MyArray(3) = 400
MyArray(4) = 500

MyArray(5) = 600 '<<< Will Return an error because there is not 5th element.

'LOOP THROUGH ARRAYS

'Using For Loop
Dim i As Long
For i = LBound(MyArray) To UBound(MyArray)
    Debug.Print MyArray(i)
Next

'Using For Each Loop
Dim Elem As Variant
For Each Elem In MyArray
    Debug.Print Elem
Next

'USE ERASE

'Declare Static Array
Dim MyArray(0 To 3) As Long
Erase MyArray  '<<< All Values will be set to 0.

'Declare Dynamic Array
Dim MyArray() As Long
ReDim MyArray(0 To 3)
Erase MyArray  '<<< Array is erased from memory.

'USE REDIM

Dim MyArray() As Variant
MyArray(0) = "MyFirstElement"

'Old Array with "MyFirstElement" is now deleted.
ReDim MyArray(0 To 4)


Dim MyArray() As Variant
MyArray(0) = "MyFirstElement"

'Old Array with "MyFirstElement" is now Resized With Original Content Kept in Place.
ReDim Preserve MyArray(0 To 4)

'USING MULTIDIMENSIONAL ARRAYS

'Declare two dimensional array
Dim MultiDimArray(0 To 3, 0 To 3) As Integer
Dim i, j As Integer

'Assign values to array
For i = LBound(MultiDimArray, 1) To UBound(MultiDimArray, 1)
    For j = LBound(MultiDimArray, 2) To UBound(MultiDimArray, 2)
        MultiDimArray(i, j) = i + j
    Next j
Next i

'Print values from array.
For i = LBound(MultiDimArray, 1) To UBound(MultiDimArray, 1)
    For j = LBound(MultiDimArray, 2) To UBound(MultiDimArray, 2)
        Debug.Print MultiDimArray(i, j)
    Next j
Next i

End Sub
