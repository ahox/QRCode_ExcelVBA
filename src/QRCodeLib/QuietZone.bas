Attribute VB_Name = "QuietZone"
Option Private Module
Option Explicit

Public Function Place(ByRef moduleMatrix() As Variant, Optional ByVal width As Long = 4) As Variant()
    If width < 0 Then Call Err.Raise(5)

    Dim sz As Long
    sz = UBound(moduleMatrix) + width * 2

    Dim ret() As Variant
    ReDim ret(sz)

    Dim rowArray() As Long

    Dim i As Long
    For i = 0 To sz
        ReDim rowArray(sz)
        ret(i) = rowArray
    Next

    Dim r As Long
    Dim c As Long
    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            ret(r + width)(c + width) = moduleMatrix(r)(c)
        Next
    Next

    Place = ret
End Function
