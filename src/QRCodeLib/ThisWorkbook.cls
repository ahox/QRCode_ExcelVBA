VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Application.MacroOptions _
        Macro:="QR", _
        Description:="この関数はQRコードを生成し0/1の文字列形式の値を返します。各行は半角空白で区切られます。 (各行の長さは=FIND("" "",QR文字列)で得られます。 行ごとの分割は=MID(QR文字列,行番号*各行の長さ+1,各行の長さ)で行なうことができます。 各セルの値は=MID(行ごとの分割,列番号,1)で得られます。)", _
        Category:="Web", _
        ArgumentDescriptions:=Array( _
            "QRコードの内容", _
            "エラー訂正レベル (L, M, H, Qのいずれか。デフォルトはM)", _
            "最大バージョン (1から40の整数。デフォルトは40)", _
            "最大バージョンを超えるデータ容量の場合に連結QRコードの生成を許可するか (TRUEで許可。デフォルトはFALSE)", _
            "文字コード名 (UTF-8, Unicode, iso-2022-jp, EUC-JP等を指定。デフォルトはShift_JIS)", _
            "固定サイズにするか (TRUEで固定サイズにする。デフォルトはFALSE)")

End Sub
