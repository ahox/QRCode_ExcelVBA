Attribute VB_Name = "WorksheetFunctionQR"
Option Explicit

Public Function QR(data As String, _
    Optional ecLevel As String = "M", _
    Optional maxVer As Long = Constants.MAX_VERSION, _
    Optional allowStructuredAppend As Boolean = False, _
    Optional charsetName As String = Charset.SHIFT_JIS, _
    Optional fixedSize As Boolean = False) As String
Attribute QR.VB_Description = "���̊֐���QR�R�[�h�𐶐���0/1�̕�����`���̒l��Ԃ��܂��B�e�s�͔��p�󔒂ŋ�؂��܂��B (�e�s�̒�����=FIND("" "",QR������)�œ����܂��B �s���Ƃ̕�����=MID(QR������,�s�ԍ�*�e�s�̒���+1,�e�s�̒���)�ōs�Ȃ����Ƃ��ł��܂��B �e�Z���̒l��=MID(�s���Ƃ̕���,��ԍ�,1)�œ����܂��B)"
Attribute QR.VB_ProcData.VB_Invoke_Func = " \n18"

    Dim p_ecLevel As ErrorCorrectionLevel
    Select Case ecLevel
        Case "L"
            p_ecLevel = l
        Case "M"
            p_ecLevel = m
        Case "H"
            p_ecLevel = H
        Case "Q"
            p_ecLevel = Q
        Case Else
            QR = CVErr(xlErrNA)
            Exit Function
    End Select
    If Not (Constants.MIN_VERSION <= maxVer And maxVer <= Constants.MAX_VERSION) Then
        QR = CVErr(xlErrNA)
        Exit Function
    End If

    Dim charEncoding As New Encoding
    Call charEncoding.Init(charsetName)
    
    If Len(data) = 0 Then Exit Function
    
    Dim sbls As Symbols
    Set sbls = CreateSymbols(p_ecLevel, maxVer, allowStructuredAppend, charsetName, fixedSize)
    Call sbls.AppendText(data)
    
    Dim sbl As Variant
    For Each sbl In sbls
        QR = QR & sbl.GetString() & " "
    Next
        
End Function


