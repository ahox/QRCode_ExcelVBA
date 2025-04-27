Attribute VB_Name = "Separator"
Option Private Module
Option Explicit

Private Const VAL As Long = Values.SEPARATOR_PTN

Public Sub Place(ByRef moduleMatrix() As Variant)
    Dim Offset As Long
    Offset = UBound(moduleMatrix) - 7

    Dim i As Long
    For i = 0 To 7
         moduleMatrix(i)(7) = -VAL
         moduleMatrix(7)(i) = -VAL

         moduleMatrix(Offset + i)(7) = -VAL
         moduleMatrix(Offset + 0)(i) = -VAL

         moduleMatrix(i)(Offset + 0) = -VAL
         moduleMatrix(7)(Offset + i) = -VAL
     Next
End Sub
