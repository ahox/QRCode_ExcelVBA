VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ErrorCorrectionLevel
    l
    m
    Q
    H
End Enum

Public Enum EncodingMode
    UNKNOWN
    NUMERIC
    ALPHA_NUMERIC
    EIGHT_BIT_BYTE
    KANJI
End Enum

