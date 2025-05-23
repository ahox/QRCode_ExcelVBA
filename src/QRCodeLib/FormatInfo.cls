VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const VAL As Long = Values.FORMAT_INFO

Private m_values()     As Long
Private m_maskValues() As Long

Private Sub Class_Initialize()
    m_values = LongArray( _
        &H0&, &H537&, &HA6E&, &HF59&, &H11EB&, &H14DC&, &H1B85&, &H1EB2&, &H23D6&, &H26E1&, _
        &H29B8&, &H2C8F&, &H323D&, &H370A&, &H3853&, &H3D64&, &H429B&, &H47AC&, &H48F5&, &H4DC2&, _
        &H5370&, &H5647&, &H591E&, &H5C29&, &H614D&, &H647A&, &H6B23&, &H6E14&, &H70A6&, &H7591&, _
        &H7AC8&, &H7FFF& _
    )

    m_maskValues = LongArray(0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0, 1)
End Sub

Public Sub Place(ByVal ecLevel As ErrorCorrectionLevel, _
                 ByVal maskPattern As Long, _
                 ByRef moduleMatrix() As Variant)
    Dim formatInfoValue As Long
    formatInfoValue = GetValue(ecLevel, maskPattern)

    Dim temp As Long
    Dim v    As Long

    Dim r1 As Long
    r1 = 0

    Dim c1 As Long
    c1 = UBound(moduleMatrix)

    Dim i1 As Long
    For i1 = 0 To 7
        temp = IIf((formatInfoValue And (2 ^ i1)) > 0, 1, 0) Xor m_maskValues(i1)
        v = IIf(temp > 0, VAL, -VAL)
        moduleMatrix(r1)(8) = v
        moduleMatrix(8)(c1) = v

        r1 = r1 + 1
        c1 = c1 - 1

        If r1 = 6 Then
            r1 = r1 + 1
        End If
    Next

    Dim r2 As Long
    r2 = UBound(moduleMatrix) - 6

    Dim c2 As Long
    c2 = 7

    Dim i2 As Long
    For i2 = 8 To 14
        temp = IIf((formatInfoValue And (2 ^ i2)) > 0, 1, 0) Xor m_maskValues(i2)
        v = IIf(temp > 0, VAL, -VAL)
        moduleMatrix(r2)(8) = v
        moduleMatrix(8)(c2) = v

        r2 = r2 + 1
        c2 = c2 - 1

        If c2 = 6 Then
            c2 = c2 - 1
        End If
    Next

    moduleMatrix(UBound(moduleMatrix) - 7)(8) = VAL
End Sub

Public Sub PlaceTempBlank(ByRef moduleMatrix() As Variant)
    Dim i1 As Long
    For i1 = 0 To 8
        If i1 <> 6 Then
            moduleMatrix(8)(i1) = -VAL
            moduleMatrix(i1)(8) = -VAL
        End If
    Next

    Dim i2 As Long
    For i2 = UBound(moduleMatrix) - 7 To UBound(moduleMatrix)
        moduleMatrix(8)(i2) = -VAL
        moduleMatrix(i2)(8) = -VAL
    Next

    moduleMatrix(UBound(moduleMatrix) - 7)(8) = -VAL
End Sub

Private Function GetValue(ByVal ecLevel As ErrorCorrectionLevel, _
                          ByVal maskPattern As Long) As Long
    Dim indicator As Long

    Select Case ecLevel
        Case ErrorCorrectionLevel.l
            indicator = 1
        Case ErrorCorrectionLevel.m
            indicator = 0
        Case ErrorCorrectionLevel.Q
            indicator = 3
        Case ErrorCorrectionLevel.H
            indicator = 2
        Case Else
            Call Err.Raise(5)
    End Select

    GetValue = m_values((indicator * 2 ^ 3) Or maskPattern)
End Function
