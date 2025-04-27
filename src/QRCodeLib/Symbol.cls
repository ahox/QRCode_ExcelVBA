VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Symbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_parent As Symbols

Private m_position As Long

Private m_currEncoder      As IQRCodeEncoder
Private m_currEncodingMode As EncodingMode
Private m_currVersion      As Long

Private m_dataBitCapacity As Long
Private m_dataBitCounter  As Long

Private m_segments As Collection
Private m_segmentCounter(EncodingMode.NUMERIC To EncodingMode.KANJI) As Long

Private m_index As Long

Friend Sub Init(ByVal parentObj As Symbols)
    Set m_parent = parentObj
    m_index = parentObj.Count

    m_position = parentObj.Count

    Set m_currEncoder = Nothing
    m_currEncodingMode = EncodingMode.UNKNOWN
    m_currVersion = parentObj.MinVersion

    m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
        parentObj.ErrorCorrectionLevel, parentObj.MinVersion)

    m_dataBitCounter = 0

    Set m_segments = New Collection

    If parentObj.StructuredAppend Then
        m_dataBitCapacity = m_dataBitCapacity - StructuredAppend.HEADER_LENGTH
    End If
End Sub

Friend Property Get Parent() As Symbols
    Set Parent = m_parent
End Property

Friend Property Get Index() As Long
    Index = m_index
End Property

Friend Property Get Version() As Long
    Version = m_currVersion
End Property

Friend Property Get CurrentEncodingMode() As EncodingMode
    CurrentEncodingMode = m_currEncodingMode
End Property

Friend Function TryAppend(ByVal c As String) As Boolean
    Dim bitLength As Long
    bitLength = m_currEncoder.GetCodewordBitLength(c)

    Do While (m_dataBitCapacity < m_dataBitCounter + bitLength)
        If m_currVersion >= m_parent.MaxVersion Then
            TryAppend = False
            Exit Function
        End If

        Call SelectVersion
    Loop

    Call m_currEncoder.Append(c)
    m_dataBitCounter = m_dataBitCounter + bitLength
    Call m_parent.UpdateParity(c)

    TryAppend = True
End Function

Friend Function TrySetEncodingMode(ByVal encMode As EncodingMode, _
                                   ByVal c As String) As Boolean
    Dim enc As IQRCodeEncoder

    Select Case encMode
        Case EncodingMode.NUMERIC
            Set enc = New NumericEncoder
        Case EncodingMode.ALPHA_NUMERIC
            Set enc = New AlphanumericEncoder
        Case EncodingMode.EIGHT_BIT_BYTE
            Set enc = New ByteEncoder
            Call enc.Init(m_parent.Encoding)
        Case EncodingMode.KANJI
            If Charset.IsJP(m_parent.Encoding.Charset) Then
                Set enc = New KanjiEncoder
                Call enc.Init(m_parent.Encoding)
            Else
                Call Err.Raise(51)
            End If
        Case Else
            Call Err.Raise(5)
    End Select

    Dim bitLength As Long
    bitLength = enc.GetCodewordBitLength(c)

    Do While (m_dataBitCapacity < _
                m_dataBitCounter + _
                ModeIndicator.Length + _
                CharCountIndicator.GetLength(m_currVersion, encMode) + _
                bitLength)
        If m_currVersion >= m_parent.MaxVersion Then
            TrySetEncodingMode = False
            Exit Function
        End If

        Call SelectVersion
    Loop

    m_dataBitCounter = m_dataBitCounter + _
                       ModeIndicator.Length + _
                       CharCountIndicator.GetLength(m_currVersion, encMode)

    Set m_currEncoder = enc
    Call m_segments.Add(enc)
    m_segmentCounter(encMode) = m_segmentCounter(encMode) + 1
    m_currEncodingMode = encMode

    TrySetEncodingMode = True
End Function

Private Sub SelectVersion()
    Dim num As Long
    Dim encMode As EncodingMode
    For encMode = EncodingMode.NUMERIC To EncodingMode.KANJI
        num = m_segmentCounter(encMode)
        m_dataBitCounter = m_dataBitCounter + _
                           num * CharCountIndicator.GetLength( _
                                m_currVersion + 1, encMode) - _
                           num * CharCountIndicator.GetLength( _
                                m_currVersion + 0, encMode)
    Next

    m_currVersion = m_currVersion + 1
    m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion)
    m_parent.MinVersion = m_currVersion

    If m_parent.StructuredAppend Then
        m_dataBitCapacity = m_dataBitCapacity - StructuredAppend.HEADER_LENGTH
    End If
End Sub

Private Function BuildDataBlock() As Variant()
    Dim dataBytes() As Byte
    dataBytes = GetMessageBytes()

    Dim numPreBlocks As Long
    numPreBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)

    Dim numFolBlocks As Long
    numFolBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)

    Dim ret() As Variant
    ReDim ret(numPreBlocks + numFolBlocks - 1)

    Dim sz As Long
    sz = RSBlock.GetNumberDataCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)

    Dim srcIdx As Long
    srcIdx = 0

    Dim data() As Byte

    Dim i As Long
    For i = 0 To numPreBlocks - 1
        ReDim data(sz - 1)
        Call ArrayUtil.Copy(data, 0, dataBytes, srcIdx, sz)
        srcIdx = srcIdx + sz
        ret(i) = data
    Next

    sz = RSBlock.GetNumberDataCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)

    For i = numPreBlocks To numPreBlocks + numFolBlocks - 1
        ReDim data(sz - 1)
        Call ArrayUtil.Copy(data, 0, dataBytes, srcIdx, sz)
        srcIdx = srcIdx + sz
        ret(i) = data
    Next

    BuildDataBlock = ret
End Function

Private Function BuildErrorCorrectionBlock(ByRef dataBlock() As Variant) As Variant()
    Dim numECCodewords As Long
    numECCodewords = RSBlock.GetNumberECCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion)

    Dim numPreBlocks As Long
    numPreBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)

    Dim numFolBlocks As Long
    numFolBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)

    Dim ret() As Variant
    ReDim ret(numPreBlocks + numFolBlocks - 1)

    Dim eccDataTmp() As Byte
    ReDim eccDataTmp(numECCodewords - 1)

    Dim idx As Long
    For idx = 0 To UBound(ret)
        ret(idx) = eccDataTmp
    Next

    Dim gp() As Long
    gp = GeneratorPolynomials.Item(numECCodewords)

    Dim eccIdx As Long
    Dim data() As Long
    Dim exp    As Long

    Dim blockIdx As Long
    Dim i As Long
    Dim j As Long
    For blockIdx = 0 To UBound(dataBlock)
        ReDim data(UBound(dataBlock(blockIdx)) + UBound(ret(blockIdx)) + 1)
        eccIdx = UBound(data)

        For i = 0 To UBound(dataBlock(blockIdx))
            data(eccIdx) = dataBlock(blockIdx)(i)
            eccIdx = eccIdx - 1
        Next

        For i = UBound(data) To numECCodewords Step -1
            If data(i) > 0 Then
                exp = GaloisField256.ToExp(data(i))
                eccIdx = i

                For j = UBound(gp) To 0 Step -1
                    data(eccIdx) = data(eccIdx) Xor _
                                   GaloisField256.ToInt((gp(j) + exp) Mod 255)
                    eccIdx = eccIdx - 1
                Next
            End If
        Next

        eccIdx = numECCodewords - 1

        For i = 0 To UBound(ret(blockIdx))
            ret(blockIdx)(i) = data(eccIdx)
            eccIdx = eccIdx - 1
        Next
    Next

    BuildErrorCorrectionBlock = ret
End Function

Private Function GetEncodingRegionBytes() As Byte()
    Dim dataBlock() As Variant
    dataBlock = BuildDataBlock()

    Dim ecBlock() As Variant
    ecBlock = BuildErrorCorrectionBlock(dataBlock)

    Dim numCodewords As Long
    numCodewords = Codeword.GetTotalNumber(m_currVersion)

    Dim numDataCodewords As Long
    numDataCodewords = DataCodeword.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion)

    Dim ret() As Byte
    ReDim ret(numCodewords - 1)

    Dim r As Long
    Dim c As Long

    Dim n As Long
    n = 0

    Dim idx As Long
    idx = 0
    Do While idx < numDataCodewords
        r = n Mod (UBound(dataBlock) + 1)
        c = n \ (UBound(dataBlock) + 1)

        If c <= UBound(dataBlock(r)) Then
            ret(idx) = dataBlock(r)(c)
            idx = idx + 1
        End If

        n = n + 1
    Loop

    n = 0

    Do While idx < numCodewords
        r = n Mod (UBound(ecBlock) + 1)
        c = n \ (UBound(ecBlock) + 1)

        If c <= UBound(ecBlock(r)) Then
            ret(idx) = ecBlock(r)(c)
            idx = idx + 1
        End If

        n = n + 1
    Loop

    GetEncodingRegionBytes = ret
End Function

Private Function GetMessageBytes() As Byte()
    Dim bs As New BitSequence

    If m_parent.Count > 1 Then
        Call WriteStructuredAppendHeader(bs)
    End If

    Call WriteSegments(bs)
    Call WriteTerminator(bs)
    Call WritePaddingBits(bs)
    Call WritePadCodewords(bs)

    GetMessageBytes = bs.GetBytes()
End Function

Private Sub WriteStructuredAppendHeader(ByVal bs As BitSequence)
    Call bs.Append(ModeIndicator.STRUCTURED_APPEND_VALUE, _
                   ModeIndicator.Length)
    Call bs.Append(m_position, _
                   SymbolSequenceIndicator.POSITION_LENGTH)
    Call bs.Append(m_parent.Count - 1, _
                   SymbolSequenceIndicator.TOTAL_NUMBER_LENGTH)
    Call bs.Append(m_parent.Parity, _
                   StructuredAppend.PARITY_DATA_LENGTH)
End Sub

Private Sub WriteSegments(ByVal bs As BitSequence)
    Dim data() As Byte
    Dim codewordBitLength As Long

    Dim segment As IQRCodeEncoder
    Dim i As Long
    For Each segment In m_segments
        Call bs.Append(segment.ModeIndicator, ModeIndicator.Length)
        Call bs.Append(segment.CharCount, _
                       CharCountIndicator.GetLength( _
                            m_currVersion, segment.EncodingMode))

        data = segment.GetBytes()

        For i = 0 To UBound(data) - 1
            Call bs.Append(data(i), 8)
        Next

        codewordBitLength = segment.BitCount Mod 8

        If codewordBitLength = 0 Then
            codewordBitLength = 8
        End If

        Call bs.Append(data(UBound(data)) \ _
                       2 ^ (8 - codewordBitLength), codewordBitLength)
    Next
End Sub

Private Sub WriteTerminator(ByVal bs As BitSequence)
    Dim terminatorLength As Long
    terminatorLength = m_dataBitCapacity - m_dataBitCounter

    If terminatorLength > ModeIndicator.Length Then
        terminatorLength = ModeIndicator.Length
    End If

    Call bs.Append(ModeIndicator.TERMINATOR_VALUE, terminatorLength)
End Sub

Private Sub WritePaddingBits(ByVal bs As BitSequence)
    If bs.Length Mod 8 > 0 Then
        Call bs.Append(&H0, 8 - (bs.Length Mod 8))
    End If
End Sub

Private Sub WritePadCodewords(ByVal bs As BitSequence)
    Dim numDataCodewords As Long
    numDataCodewords = DataCodeword.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion)

    Dim flg As Boolean
    flg = True

    Dim cnt As Long
    cnt = ((8 * numDataCodewords) - bs.Length) \ 8

    Dim i As Long
    For i = 1 To cnt
        Call bs.Append(IIf(flg, 236, 17), 8)
        flg = Not flg
    Next
End Sub

Private Function GetModuleMatrix() As Variant()
    Dim numModulesPerSide As Long
    numModulesPerSide = Module.GetNumModulesPerSide(m_currVersion)

    Dim moduleMatrix() As Variant
    ReDim moduleMatrix(numModulesPerSide - 1)

    Dim rowArray() As Long

    Dim i As Long
    For i = 0 To UBound(moduleMatrix)
        ReDim rowArray(numModulesPerSide - 1)
        moduleMatrix(i) = rowArray
    Next

    Call FinderPattern.Place(moduleMatrix)
    Call Separator.Place(moduleMatrix)
    Call TimingPattern.Place(moduleMatrix)

    If m_currVersion >= 2 Then
        Call AlignmentPattern.Place(m_currVersion, moduleMatrix)
    End If

    Call FormatInfo.PlaceTempBlank(moduleMatrix)

    If m_currVersion >= 7 Then
        Call VersionInfo.PlaceTempBlank(moduleMatrix)
    End If

    Call PlaceSymbolChar(moduleMatrix)
    Call RemainderBit.Place(moduleMatrix)

    Call Masking.Apply(m_currVersion, m_parent.ErrorCorrectionLevel, moduleMatrix)

    GetModuleMatrix = QuietZone.Place(moduleMatrix)
End Function

Private Sub PlaceSymbolChar(ByRef moduleMatrix() As Variant)
    Const VAL As Long = Values.WORD

    Dim data() As Byte
    data = GetEncodingRegionBytes()

    Dim r As Long
    r = UBound(moduleMatrix)

    Dim c As Long
    c = UBound(moduleMatrix(0))

    Dim toLeft As Boolean
    toLeft = True

    Dim rowDirection As Long
    rowDirection = -1

    Dim v As Variant
    Dim bitPos As Long
    For Each v In data
        bitPos = 7

        Do While bitPos >= 0
            If moduleMatrix(r)(c) = Values.BLANK Then
                moduleMatrix(r)(c) = IIf((v And 2 ^ bitPos) > 0, VAL, -VAL)
                bitPos = bitPos - 1
            End If

            If toLeft Then
                c = c - 1
            Else
                If (r + rowDirection) < 0 Then
                    r = 0
                    rowDirection = 1
                    c = c - 1

                    If c = 6 Then
                        c = 5
                    End If

                ElseIf ((r + rowDirection) > UBound(moduleMatrix)) Then
                    r = UBound(moduleMatrix)
                    rowDirection = -1
                    c = c - 1

                    If c = 6 Then
                        c = 5
                    End If

                Else
                    r = r + rowDirection
                    c = c + 1
                End If
            End If

            toLeft = Not toLeft
        Loop
    Next
End Sub
Private Function GetMonochromeString() As String
    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()
    
    Dim totalLength As Long
    Dim r As Long, c As Long
    Dim v As Variant
    
    ' 文字列の長さを計算
    totalLength = -1 ' 過剰に加算される末尾のスペース分をあらかじめ引いておく
    For r = LBound(moduleMatrix) To UBound(moduleMatrix)
        totalLength = totalLength + UBound(moduleMatrix(r)) - LBound(moduleMatrix(r)) + 2 '各行の長さ+スペース
    Next
    
    ' 結果文字列を初期化
    Dim result As String
    result = String(totalLength, " ")
    
    ' Mid関数を使って文字列を構築
    Dim pos As Long
    pos = 1
    For r = LBound(moduleMatrix) To UBound(moduleMatrix)
        For c = LBound(moduleMatrix(r)) To UBound(moduleMatrix(r))
            If Values.IsDark(moduleMatrix(r)(c)) Then
                Mid(result, pos, 1) = "1"
            Else
                Mid(result, pos, 1) = "0"
            End If
            pos = pos + 1
        Next
        If r < UBound(moduleMatrix) Then
            Mid(result, pos, 1) = " "
            pos = pos + 1
        End If
    Next
    
    GetMonochromeString = result
    
End Function

Public Function GetString() As String
    If m_dataBitCounter = 0 Then Call Err.Raise(51)
    GetString = GetMonochromeString()
End Function
