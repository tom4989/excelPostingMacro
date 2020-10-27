VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls転記元シート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************************************************************
' 定数
' *********************************************************************************************************************
'
Private Const KEY_転記元情報 = "転記元情報"
Private Const KEY_ファイルパス = "ファイルパス"
Private Const KEY_シート名 = "シート名"
Private Const KEY_ID列 = "ID列"
Private Const KEY_名前列 = "名前列"
Private Const KEY_型列 = "型列"
Private Const KEY_桁列 = "桁列"

' *********************************************************************************************************************
' 変数
' *********************************************************************************************************************
'
Private strファイルパス As String
Private strシート名 As String
Private lngID列 As Long
Private lng名前列 As Long
Private lng型列 As Long
Private lng桁列 As Long

Private varシート内容 As Variant

' *********************************************************************************************************************
' 定数(Public) ※ クラス内では定数をPublicにできないため、Propertyを利用
' *********************************************************************************************************************
'
Private Const C_IDX_ID = 0
Private Const C_IDX_名前 = 1
Private Const C_IDX_型 = 2
Private Const C_IDX_桁 = 3

Private Const IDX_COUNT = 3

Public Property Get IDX_ID() As Long
    IDX_ID = C_IDX_ID
End Property

Public Property Get IDX_名前() As Long
    IDX_名前 = C_IDX_名前
End Property

Public Property Get IDX_型() As Long
    IDX_型 = C_IDX_型
End Property

Public Property Get IDX_桁() As Long
    IDX_桁 = C_IDX_桁
End Property

' *********************************************************************************************************************
' 機能：コンストラクタ ※ 引数は指定できない
' *********************************************************************************************************************
'
Public Sub Class_Initialize()

End Sub

' *********************************************************************************************************************
' 機能：初期化
' *********************************************************************************************************************
'
Public Function 初期化(obj設定値シート As cls設定値シート)

    strファイルパス = obj設定値シート.設定値リスト.Item(KEY_転記元情報).Item(KEY_ファイルパス)
    strシート名 = obj設定値シート.設定値リスト.Item(KEY_転記元情報).Item(KEY_シート名)
    lngID列 = CAlp2Num(obj設定値シート.設定値リスト.Item(KEY_転記元情報).Item(KEY_ID列))
    lng名前列 = CAlp2Num(obj設定値シート.設定値リスト.Item(KEY_転記元情報).Item(KEY_名前列))
    lng型列 = CAlp2Num(obj設定値シート.設定値リスト.Item(KEY_転記元情報).Item(KEY_型列))
    lng桁列 = CAlp2Num(obj設定値シート.設定値リスト.Item(KEY_転記元情報).Item(KEY_桁列))
    
End Function

' *********************************************************************************************************************
' 機能：転記元ブックの転記元シートの内容をVariant変数に格納する
' *********************************************************************************************************************
'
Public Sub シート読込()

    Dim wb転記元 As Workbook
    Set wb転記元 = Workbooks.Open(strファイルパス, UpdateLinks:=0, IgnoreReadOnlyRecommended:=False)
    
    Dim ws転記元 As Worksheet
    Set ws転記元 = wb転記元.Sheets(strシート名)

    Dim lng最終行, lng最終列 As Long
    lng最終行 = 最終行取得(ws転記元)
    lng最終列 = 最終列取得(ws転記元)

    With ws転記元
        varシート内容 = .Range(.Cells(1, 1), .Cells(lng最終行, lng最終列))
    End With

    wb転記元.Close

End Sub

' *********************************************************************************************************************
' 機能：
' *********************************************************************************************************************
'
Public Function 転記元情報取得(str名前 As String) As Variant

    Dim var転記元情報 As Variant
    Dim i As Long
    
    For i = LBound(varシート内容) To UBound(varシート内容)
    
        If str名前 = varシート内容(i, lng名前列) Then
        
            ReDim var転記元情報(IDX_COUNT)
        
            var転記元情報(IDX_ID) = varシート内容(i, lngID列)
            var転記元情報(IDX_名前) = varシート内容(i, lng名前列)
            var転記元情報(IDX_型) = varシート内容(i, lng型列)
            var転記元情報(IDX_桁) = varシート内容(i, lng桁列)
            
            転記元情報取得 = var転記元情報
            Exit Function
            
        End If
    Next
    
    転記元情報取得 = var転記元情報
    
End Function
