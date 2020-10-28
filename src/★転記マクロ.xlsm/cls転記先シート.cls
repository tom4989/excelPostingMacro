VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls転記先シート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************************************************************
' 定数
' *********************************************************************************************************************
'

' *********************************************************************************************************************
' 変数
' *********************************************************************************************************************
'
Private lngID列 As Long
Private lng名前列From As Long
Private lng名前列To As Long
Private lng型列 As Long
Private lng整数桁列 As Long
Private lng小数桁列 As Long

' *********************************************************************************************************************
' 定数(Public) ※ クラス内では定数をPublicにできないため、Propertyを利用
' *********************************************************************************************************************
'

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
Public Function 初期化(obj設定値シート As cls設定値シート, wb転記先ブック As Workbook) As Boolean

    Dim obj転記先設定値 As Object
    Set obj転記先設定値 = obj設定値シート.設定値リスト.Item("転記先情報")
    
    Dim varBookName As Variant
    varBookName = obj転記先設定値.Keys
    
    Dim i As Long
    
    For i = LBound(varBookName) To UBound(varBookName)
    
        If wb転記先ブック.Name Like "*" & varBookName(i) & "*" Then
        
            Dim var名前列 As Variant
            var名前列 = CAlpxAlp2Num(obj転記先設定値.Item(varBookName(i)).Item("名前列"))
        
            lng名前列From = CLng(var名前列(0))
            lng名前列To = CLng(var名前列(1))
            
            lngID列 = CAlp2Num(obj転記先設定値.Item(varBookName(i)).Item("ID列"))
            lng型列 = CAlp2Num(obj転記先設定値.Item(varBookName(i)).Item("型列"))
            lng整数桁列 = CAlp2Num(obj転記先設定値.Item(varBookName(i)).Item("整数桁列"))
            lng小数桁列 = CAlp2Num(obj転記先設定値.Item(varBookName(i)).Item("小数桁列"))
        
            初期化 = True
            Exit Function
        
        End If
    Next
    
    ' 呼び出し元のブックの情報が取得できなかった場合
    MsgBox "対象外のブックです：" & wb転記先ブック.Name
    初期化 = False
    
 End Function
    
' *********************************************************************************************************************
' 機能：転記
' *********************************************************************************************************************
'
Public Function 転記(wb転記先ブック As Workbook, obj転記元シート As cls転記元シート) As Boolean
        
    Dim varエラーメッセージ() As Variant
    Dim selectionRange, selectionRow As Range
    Dim i, lng処理済の行 As Long
    
    With wb転記先ブック.ActiveSheet
    
        ' 転記先をActiveにする
        .Activate
    
        ' 選択された範囲を順番に処理する。
        For Each selectionRange In Selection
    
            ' Selectionは、1セルごとにRangeを返すが、列を複数選択していると同じ行が列数分対象になるため対策する
            If (selectionRange.Row <> lng処理済の行) Then
    
                ' 名前を特定する
                Dim str名前 As String
                str名前 = ""
                
                For i = lng名前列From To lng名前列To
                
                    If .Cells(selectionRange.Row, i).Value <> "" Then
                    
                        str名前 = .Cells(selectionRange.Row, i).Value
                        Exit For
                    
                    End If
                Next
                
                ' 転記元情報の取得
                Dim var転記元情報 As Variant
                var転記元情報 = obj転記元シート.転記元情報取得(str名前)
                
                If IsArrayEx(var転記元情報) <> 1 Then
                
                    Call 一次配列に値を追加(varエラーメッセージ, str名前 & "：転記元に項目が存在しません。")

                Else
                
                If var転記元情報(obj転記元シート.IDX_ID) = "" Then
                    
                    Call 一次配列に値を追加(varエラーメッセージ, str名前 & "：IDが未記入です。")
                    
                End If
            
                    .Cells(selectionRange.Row, lngID列) = var転記元情報(obj転記元シート.IDX_ID)
                    .Cells(selectionRange.Row, lng型列) = var転記元情報(obj転記元シート.IDX_型)
                    .Cells(selectionRange.Row, lng整数桁列) = var転記元情報(obj転記元シート.IDX_整数桁)
                    .Cells(selectionRange.Row, lng小数桁列) = var転記元情報(obj転記元シート.IDX_小数桁)
                
                End If
        
            End If
            
            lng処理済の行 = selectionRange.Row
            
        Next
                
    End With

    ' エラーメッセージの出力
    If IsArrayEx(varエラーメッセージ) = 1 Then
    
        Dim strエラーメッセージ As String
        
        For i = LBound(varエラーメッセージ) To UBound(varエラーメッセージ)
        
            If 0 < i Then
                strエラーメッセージ = strエラーメッセージ & vbCrLf
            End If
            
            strエラーメッセージ = strエラーメッセージ & varエラーメッセージ(i)
        Next
        
        Call MsgBox(strエラーメッセージ, vbExclamation)
    
    End If

End Function
