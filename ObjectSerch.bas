Option Explicit

'------------------------------------------------------------------------------------------------------------------------
' グローバル変数定義
'------------------------------------------------------------------------------------------------------------------------
Dim searchedShapes As Collection    ' 検索図形コレクション
Dim currentShapeIndex As Integer    ' 図形インデックス
Dim searchBook As Workbook          ' 検索ブック
Dim searchSheet As Worksheet        ' 検索シート

' 図形の検索範囲モード
Enum RangeMode
    sheet
    book
End Enum


'------------------------------------------------------------------------------------------------------------------------
' 図形検索
' 内容：検索文字列が含まれる図形を抽出してコレクションする
' 引数1：rangeMode 図形の検索範囲モード
' 引数2：searchString 検索する文字列
'------------------------------------------------------------------------------------------------------------------------
Sub SearchShapes(rangeMode As RangeMode, searchString As String)
  
    Dim shape As Shape              ' 図形
    Dim sheet As Worksheet          ' シート

    On Error GoTo ErrorSearchShapes
    '検索範囲モードによって処理を分岐
    Select Case rangeMode

        ' シート内検索
        Case RangeMode.sheet
            Set searchSheet = ActiveSheet           ' アクティブシートを取得
            For Each shape In ActiveSheet.Shapes    ' シート内の図形を取得
                If InStr(1, shape.TextFrame.Characters.Text, searchString, vbTextCompare) > 0 Then    ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
                    searchedShapes.Add shape        ' コレクションに図形を追加
                End If
            Next shape

        ' ブック全体検索
        Case RangeMode.book
            Set searchBook = ActiveWorkbook         ' アクティブブックを取得
            For Each sheet In ThisWorkbook.Sheets   ' ブック内のシートを取得
                For Each shape In sheet.Shapes      ' シート内の図形を取得
                    If InStr(1, shape.TextFrame.Characters.Text, searchString, vbTextCompare) > 0 Then    ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
                        searchedShapes.Add shape    ' コレクションに図形を追加
                    End If
                Next shape
            Next sheet

        ' 上記以外
        Case Else
            ' エラー処理
            MsgBox "検索範囲が不正です。"
            Exit Sub
    End Select
    Exit Sub

ErrorSearchShapes:
    MsgBox "図形検索でエラーが発生しました"
End Sub



'------------------------------------------------------------------------------------------------------------------------
' 図形表示
' 内容：選択図形を画面に表示する
'------------------------------------------------------------------------------------------------------------------------
Sub ShowShape()

    Dim shapeRow As Integer     ' 図形左上の行番号
    Dim shapeColumn As Integer  ' 図形左上の列番号

    On Error GoTo ErrorShowShape
    shapeRow = searchedShapes(currentShapeIndex).TopLeftCell.Row            ' 図形の左上の行番号を取得
    shapeColumn = searchedShapes(currentShapeIndex).TopLeftCell.Column      ' 図形の左上の列番号を取得
    Application.Goto Cells(shapeRow, shapeColumn), True  ' 図形の左上が画面の左上に来るように画面移動
    Exit Sub
    
ErrorShowShape:
    MsgBox "図形表示でエラーが発生しました"

End Sub



'------------------------------------------------------------------------------------------------------------------------
' ハイライト付与
' 内容：図形内の検索文字列にハイライトをつける
' 引数1：searchString 検索する文字列
'------------------------------------------------------------------------------------------------------------------------
sub HighlightShapeString(searchString As String)

    Dim index As Integer        ' 先頭位置

    On Error GoTo ErrorHighlightShapeString
    index = InStr(1, searchedShapes(currentShapeIndex).TextFrame.Characters.Text, searchString, vbTextCompare)      ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
    Do While index > 0          ' 文字列が見つからなくなるまで
        searchedShapes(currentShapeIndex).TextFrame.TextRange.Characters(index, Len(searchString)).Font.Fill.BackColor.RGB = RGB(255, 255, 0)     ' 背景色を設定
        index = InStr(index + 1, searchedShapes(currentShapeIndex).TextFrame.Characters.Text, searchString)         ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
    Loop

ErrorHighlightShapeString:
    MsgBox "ハイライトの付与でエラーが発生しました"
End SUb



'------------------------------------------------------------------------------------------------------------------------
' ハイライトクリア
' 内容：図形内の文字列のハイライトをすべてクリアする
'------------------------------------------------------------------------------------------------------------------------
sub ClearHighlightShape()

    On Error GoTo ErrorClearHighlightShape
    searchedShapes(currentShapeIndex).TextFrame.TextRange.Font.Fill.BackColor.RGB = RGB(255, 255, 255)     ' 背景色を設定

ErrorClearHighlightShape:
    MsgBox "ハイライトのクリアでエラーが発生しました"
End SUb



