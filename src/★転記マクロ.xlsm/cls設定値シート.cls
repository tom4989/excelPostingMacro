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
Private dic設定値リスト As Object

Public Property Get 設定値リスト() As Object
    Set 設定値リスト = dic設定値リスト
End Property

' *********************************************************************************************************************
' 機能：コンストラクタ ※ 引数は指定できない
' *********************************************************************************************************************
'
Public Sub Class_Initialize()

    ' 辞書初期化
    Set dic設定値リスト = CreateObject("Scripting.Dictionary")

    Dim varシートの値全体 As Variant
    varシートの値全体 = シート内容取得(ThisWorkbook.Sheets(cnstシート名))
    
    Dim i As Long
    
    For i = LBound(varシートの値全体) To UBound(varシートの値全体)
            
        ' A列に★が設定されている行のみ対象
        If varシートの値全体(i, 1) = "★" Then
        
            Dim j, lng開始列, lng最終列 As Long
            lng開始列 = -1
            lng最終列 = -1

            ' 要素数を数える
            For j = LBound(varシートの値全体, 2) To UBound(varシートの値全体, 2)
        
                If varシートの値全体(i, j) <> "" Then
                
                    If lng開始列 = -1 Then
                    
                        lng開始列 = j
                    
                    End If
                
                    lng最終列 = j
                End If
            Next j
    
            ' 設定値行が2階層以上設定されている場合
            If lng開始列 < lng最終列 Then
        
                ' 1行に記載された情報を解析する
                For j = lng開始列 To lng最終列
        
                    Dim objDic As Object
        
                    ' 処理中の列が、開始列の場合
                    If j = lng開始列 Then
                    
                        Set objDic = dic設定値リスト
                    
                    End If
        
                    ' 処理中の列が、最後の列でない場合
                    If j < lng最終列 Then
                
                        If objDic.Exists(varシートの値全体(i, j)) Then
                    
                            Set objDic = objDic.Item(varシートの値全体(i, j))
                    
                        Else
                            Dim objNewDic As Object
                            Set objNewDic = CreateObject("Scripting.Dictionary")
                            Call objDic.Add(varシートの値全体(i, j), objNewDic)
                        
                            Set objDic = objNewDic

                        End If
                    
                    ' 処理中の列が、最後の列の1つ前の場合
                    Else
                        Call objDic.Add(varシートの値全体(i, j), varシートの値全体(i, j + 1))
                    
                        Exit For
                    
                    End If
                Next j
            End If
        End If
    Next i
    
    
End Sub
