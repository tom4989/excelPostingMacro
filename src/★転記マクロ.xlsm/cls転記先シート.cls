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
Private Const KEY_キー列 = "キー列"

' *********************************************************************************************************************
' 変数
' *********************************************************************************************************************
'
Private dic列情報 As Object

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
        
            Set dic列情報 = obj転記先設定値.Item(varBookName(i))

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
                Dim varキー列 As Variant
                varキー列 = CAlpxAlp2Num(dic列情報.Item(KEY_キー列))
        
                Dim lng名前列From, lng名前列To As Long
                lng名前列From = CLng(varキー列(0))
                lng名前列To = CLng(varキー列(1))

                Dim txt名前 As String
                txt名前 = ""
                
                For i = lng名前列From To lng名前列To
                
                    If .Cells(selectionRange.Row, i).Value <> "" Then
                    
                        txt名前 = .Cells(selectionRange.Row, i).Value
                        Exit For
                    
                    End If
                Next
                
                If txt名前 = "" Then
                    
                    Call 一次配列に値を追加(varエラーメッセージ, _
                        selectionRange.Row & "行目" & vbTab & _
                        vbTab & _
                        "キー(" & dic列情報.Item(KEY_キー列) & "列)の値が未記入です。")
                End If
                
                ' 転記元情報の取得
                Dim dic転記元行 As Object
                Set dic転記元行 = obj転記元シート.転記元情報取得(txt名前)
                
                If dic転記元行 Is Nothing Then
                
                    Call 一次配列に値を追加(varエラーメッセージ, _
                        selectionRange.Row & "行目" & vbTab & _
                        txt名前 & vbTab & _
                        "転記元に項目が存在しません。")
                Else
                
                    Dim txt列名 As Variant
                
                    For Each txt列名 In dic転記元行
                    
                        If txt列名 <> KEY_キー列 Then
                        
                            .Cells(selectionRange.Row, CAlp2Num(dic列情報.Item(txt列名))) = _
                                dic転記元行.Item(txt列名)
                        
                        End If
                    
                    Next txt列名
                
                End If
        
            End If
            
            lng処理済の行 = selectionRange.Row
            
        Next
                
    End With

    ' エラーメッセージの出力
    If IsArrayEx(varエラーメッセージ) = 1 Then
    
        Dim txtエラーメッセージ As String
        
        For i = LBound(varエラーメッセージ) To UBound(varエラーメッセージ)
        
            If 0 < i Then
                txtエラーメッセージ = txtエラーメッセージ & vbCrLf
            End If
            
            txtエラーメッセージ = txtエラーメッセージ & varエラーメッセージ(i)
        Next
        
        Call MsgBox(txtエラーメッセージ, vbExclamation)
    
    End If

End Function
