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
    none
    sheet
    book
End Enum

' 検索方向
Enum SearchMode
    nextShape
    prevShape
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

    currentShapeIndex = 1

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
        searchedShapes(currentShapeIndex).TextFrame.TextRange.Characters(index, Len(searchString)).Font.Glow.Radius = 3 ' 光彩の半径を設定
        index = InStr(index + 1, searchedShapes(currentShapeIndex).TextFrame.Characters.Text, searchString)         ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
    Loop
    Exit Sub

ErrorHighlightShapeString:
    MsgBox "ハイライトの付与でエラーが発生しました"
End SUb


'------------------------------------------------------------------------------------------------------------------------
' ハイライトクリア
' 内容：図形内の文字列のハイライトをすべてクリアする
'------------------------------------------------------------------------------------------------------------------------
sub ClearHighlightShape()

    On Error GoTo ErrorClearHighlightShape

    searchedShapes(currentShapeIndex).TextFrame.TextRange.Font.Glow.Radius = 0 ' 光彩の半径を設定
    Exit Sub

ErrorClearHighlightShape:
    MsgBox "ハイライトのクリアでエラーが発生しました"
End SUb


'------------------------------------------------------------------------------------------------------------------------
' 検索対象の図形を変更する
' 内容：次または１つ前の図形を検索対象とする
' 引数1：targetshape 検索する図形
'------------------------------------------------------------------------------------------------------------------------
Sub SearchNextShape(targetshape As SearchMode)

    On Error GoTo ErrorSearchNextShape
    
    ' 検索方向によって分岐
    Select Case targetshape

        ' 次の図形を検索する
        Case SearchMode.nextShape
            If currentShapeIndex == searchedShapes.Count Then   ' 図形インデックスが最後の値のとき
                currentShapeIndex = 1                           ' 図形インデックスを先頭に戻す
            ElseIf
                currentShapeIndex = currentShapeIndex + 1       ' 図形インデックスの値を１つ増やす
            End If

        ' １つ前の図形を検索する
        Case SearchMode.prevShape
            If currentShapeIndex == 1 Then                      ' 図形インデックスが先頭のとき
                currentShapeIndex = searchedShapes.Count        ' 図形インデックスを最後の値にする
            ElseIf
                currentShapeIndex = currentShapeIndex - 1       ' 図形インデックスの値を１つ減らす
            End If

    ' 上記以外
        Case Else
            ' エラー処理
            MsgBox "検索対象の図形が不正です。"
            Exit Sub
    End Select
    Exit Sub

ErrorSearchNextShape:
    MsgBox "検索対象の図形の変更でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' 文字列を置換する
' 内容：検索文字列を置換文字列に置換する
' 引数1：searchString 検索文字列
' 引数2：replaceString 置換文字列
'------------------------------------------------------------------------------------------------------------------------
Sub ReplaceShapeText(searchString As String, replaceString As String)

    On Error GoTo ErrorReplaceShapeText
    searchedShapes(currentShapeIndex).TextFrame.Characters.Text = Replace(searchedShapes(currentShapeIndex).TextFrame.Characters.Text, searchString, replaceString) '文字列の置換
    Exit Sub

ErrorReplaceShapeText:
    MsgBox "文字の置換でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' すべての図形で文字列を置換する
' 内容：すべての図形で検索文字列を置換文字列に置換する
' 引数1：searchString 検索文字列
' 引数2：replaceString 置換文字列
'------------------------------------------------------------------------------------------------------------------------
Sub ReplaceAllText(searchString As String, replaceString As String)

    On Error GoTo ErrorReplaceAllText

    For shapeIndex To searchedShapes.Count                      ' 検索図形コレクションの回数実行
        currentShapeIndex = shapeIndex                          ' 図形インデックスを更新
        Call ReplaceShapeText(searchString, replaceString)      ' 文字列の置換
    Next shapeIndex
    Exit Sub

ErrorReplaceAllText:
    MsgBox "図形の参照でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' 検索ボタン押下時処理
' 内容：検索ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub btnSearchAll_Click()

    Dim searchString As String  ' 検索文字列

    On Error GoTo ErrorbtnSearchAll_Click

    searchString = txtSearchString.Text
    Call SearchShapes(cmbSearchRange.ListIndex, searchString)     ' 図形検索
    Call ShowShape()        ' 図形に移動
    Call HighlightShapeString(searchString)     ' 文字列にハイライト付与
    Exit Sub

ErrorbtnSearchAll_Click:
    MsgBox "検索ボタンの処理でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' 次を検索ボタン押下時処理
' 内容：次を検索ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub btnSearchNextShape_Click()

    Dim searchString As String  ' 検索文字列

    On Error GoTo ErrorbtnSearchNextShape_Click

    searchString = txtSearchString.Text

    Call ClearHighlightShape()      ' 文字列のハイライトをクリア
    Call SearchNextShape(SearchMode.nextShape)    ' 次の図形を対象に変更
    Call ShowShape()        ' 図形に移動
    Call HighlightShapeString(searchString)     ' 文字列にハイライト付与
    Exit Sub

ErrorbtnSearchNextShape_Click:
    MsgBox "次を検索ボタンの処理でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' １つ前を検索ボタン押下時処理
' 内容：１つ前を検索ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub btnSearchPrevShape_Click()

    Dim searchString As String  ' 検索文字列

    On Error GoTo ErrorbtnSearchPrevShape_Click

    searchString = txtSearchString.Text

    Call ClearHighlightShape()      ' 文字列のハイライトをクリア
    Call SearchNextShape(SearchMode.prevShape)    ' 次の図形を対象に変更
    Call ShowShape()        ' 図形に移動
    Call HighlightShapeString(searchString)     ' 文字列にハイライト付与
    Exit Sub

ErrorbtnSearchPrevShape_Click:
    MsgBox "１つ前を検索ボタンの処理でエラーが発生しました"
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 置換ボタン押下時処理
' 内容：置換ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub btnReplaceShape_Click()

    Dim searchString As String  ' 検索文字列
    Dim replaceString As String ' 置換文字列

    On Error GoTo ErrorbtnReplaceShape_Click

    searchString = txtSearchString.Text
    replaceString = txtReplaceString.Text

    Call ClearHighlightShape()      ' 文字列のハイライトをクリア
    Call ReplaceShapeText(searchString, replaceString)  ' 文字列を置換
    Call HighlightShapeString(searchString)     ' 文字列にハイライト付与
    Exit Sub

ErrorbtnReplaceShape_Click:
    MsgBox "置換ボタンの処理でエラーが発生しました"
End Sub

