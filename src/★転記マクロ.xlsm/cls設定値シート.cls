VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls設定値シート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************************************************************
' 定数
' *********************************************************************************************************************
'
Const cnstシート名 = "★設定値"


' *********************************************************************************************************************
' 変数
' *********************************************************************************************************************
'
Private obj設定値リスト As Object

Public Property Get 設定値リスト() As Object
    Set 設定値リスト = obj設定値リスト
End Property

' *********************************************************************************************************************
' 機能：コンストラクタ ※ 引数は指定できない
' *********************************************************************************************************************
'
Public Sub Class_Initialize()

    ' 辞書初期化
    Set obj設定値リスト = CreateObject("Scripting.Dictionary")

    Dim lng最終行 As Long
    lng最終行 = 最終行取得(ThisWorkbook.Sheets(cnstシート名))
    
    Dim lng最終列 As Long
    lng最終列 = 最終行取得(ThisWorkbook.Sheets(cnstシート名))
        
    Dim varシートの値全体 As Variant
    
    With ThisWorkbook.Sheets(cnstシート名)

        varシートの値全体 = .Range(.Cells(1, 1), .Cells(lng最終行, lng最終列))
        
    End With
    
    Dim i, j1, j2 As Long
    
    For i = LBound(varシートの値全体) To UBound(varシートの値全体)
    
        ' 要素数を数える
        For j1 = LBound(varシートの値全体, 2) To UBound(varシートの値全体, 2)
        
            If varシートの値全体(i, j1) = "" Then
                Exit For
            End If
        Next
    
        ' 空行でない場合
        If j1 > 1 Then
        
            ' 1行に記載された情報を解析する
            For j2 = LBound(varシートの値全体, 2) To UBound(varシートの値全体, 2)
        
                Dim objDic As Object
        
                ' 処理中の列が、最初の列の場合
                If j2 = 1 Then
                    
                    Set objDic = obj設定値リスト
                    
                End If
        
                ' 処理中の列が、最後の列でない場合
                If j2 < j1 - 2 Then
                
                    If objDic.Exists(varシートの値全体(i, j2)) Then
                    
                        Set objDic = objDic.Item(varシートの値全体(i, j2))
                    
                    Else
                        Dim objNewDic As Object
                        Set objNewDic = CreateObject("Scripting.Dictionary")
                        Call objDic.Add(varシートの値全体(i, j2), objNewDic)
                        
                        Set objDic = objNewDic

                    End If
                    
                ' 処理中の列が、最後の列の1つ前の場合
                Else
                    Call objDic.Add(varシートの値全体(i, j2), varシートの値全体(i, j2 + 1))
                    
                    Exit For
                    
                End If
            Next
        End If
    Next

End Sub
