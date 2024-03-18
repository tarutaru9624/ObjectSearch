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

' 検索方向
Enum SearchMode
    nextShape
    prevShape
End Enum


'------------------------------------------------------------------------------------------------------------------------
' 図形検索
' 内容：検索文字列が含まれる図形を抽出してコレクションする
' 引数1：rangeMode 図形の検索範囲モード
' 引数2：searchText 検索する文字列
'------------------------------------------------------------------------------------------------------------------------
Sub SearchShapes(rangeSpecified As RangeMode, searchText As String)
  
    Dim shape As Shape              ' 図形
    Dim sheet As Worksheet          ' シート
    Dim item As shape               ' グループ内図形

    On Error GoTo ErrorSearchShapes
    '検索範囲モードによって処理を分岐
    Select Case rangeSpecified

        ' シート内検索
        Case rangeMode.sheet
            Set searchedShapes = New Collection             ' 検索図形コレクションの生成
            Set searchSheet = ActiveSheet                   ' アクティブシートを取得
            For Each shape In searchSheet.Shapes            ' シート内の図形を取得
                If shape.Type = msoGroup Then               ' 図形がグループ化されている時
                    For Each item In shape.GroupItems
                        If InStr(1, item.TextFrame.Characters.Text, searchText, vbTextCompare) > 0 Then    ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
                            searchedShapes.Add item         ' コレクションに図形を追加
                        End If
                    Next item
                Else                                        ' 図形がグループ化されていない時
                    If InStr(1, shape.TextFrame.Characters.Text, searchText, vbTextCompare) > 0 Then    ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
                        searchedShapes.Add shape            ' コレクションに図形を追加
                    End If
                End If
                
            Next shape

        ' ブック全体検索
        Case rangeMode.book
            Set searchedShapes = New Collection             ' 検索図形コレクションの生成
            Set searchBook = ActiveWorkbook                 ' アクティブブックを取得
            For Each sheet In searchBook.Sheets             ' ブック内のシートを取得
                For Each shape In sheet.Shapes              ' シート内の図形を取得
                    If shape.Type = msoGroup Then           ' 図形がグループ化されている時
                        For Each item In shape.GroupItems
                            If InStr(1, item.TextFrame.Characters.Text, searchText, vbTextCompare) > 0 Then    ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
                                searchedShapes.Add item     ' コレクションに図形を追加
                            End If
                        Next item
                    Else                                    ' 図形がグループ化されていない時
                        If InStr(1, shape.TextFrame.Characters.Text, searchText, vbTextCompare) > 0 Then        ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
                            searchedShapes.Add shape        ' コレクションに図形を追加
                        End If
                    End If
                Next shape
            Next sheet


        ' 上記以外
        Case Else
            ' エラー処理
            MsgBox "検索範囲が不正です。"
            Exit Sub
    End Select

    If searchedShapes.Count = 0 Then
        MsgBox "見つかりませんでした。"
        End
    End If
    
    currentShapeIndex = 1
    lblShapesNum.Caption = currentShapeIndex & " / " & searchedShapes.Count       ' 図形数表示
    Call isEnableButton(False)               ' すべて検索ボタンを無効化
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
    Dim shapeSheet As Worksheet ' 図形のシート

    On Error GoTo ErrorShowShape

    shapeRow = searchedShapes(currentShapeIndex).TopLeftCell.Row            ' 図形の左上の行番号を取得
    shapeColumn = searchedShapes(currentShapeIndex).TopLeftCell.Column      ' 図形の左上の列番号を取得
    Set shapeSheet = searchedShapes(currentShapeIndex).TopLeftCell.Worksheet    ' 図形のシートを取得
    Application.Goto shapeSheet.Cells(shapeRow, shapeColumn), True  ' 図形の左上が画面の左上に来るように画面移動
    Exit Sub
    
ErrorShowShape:
    MsgBox "図形表示でエラーが発生しました"

End Sub


'------------------------------------------------------------------------------------------------------------------------
' ハイライト付与
' 内容：図形内の検索文字列にハイライトをつける
' 引数1：searchText 検索する文字列
'------------------------------------------------------------------------------------------------------------------------
sub HighlightShapeString(searchText As String)

    Dim index As Integer        ' 先頭位置

    On Error GoTo ErrorHighlightShapeString

    index = InStr(1, searchedShapes(currentShapeIndex).TextFrame.Characters.Text, searchText, vbTextCompare)      ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
    Do While index > 0          ' 文字列が見つからなくなるまで
        With searchedShapes(currentShapeIndex).TextFrame2.TextRange.Characters(index, Len(searchText)).Font.Glow
            .Color.ObjectThemeColor = msoThemeColorAccent4  ' 光彩の色を設定
            .Radius = 18 ' 光彩の半径を設定
        End With
        index = InStr(index + 1, searchedShapes(currentShapeIndex).TextFrame.Characters.Text, searchText)         ' 文字列検索 検索文字列が含まれる場合、戻り値に位置番号
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

    searchedShapes(currentShapeIndex).TextFrame2.TextRange.Characters.Font.Glow.Radius = 0 ' 光彩の半径を設定
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
            If currentShapeIndex = searchedShapes.Count Then   ' 図形インデックスが最後の値のとき
                currentShapeIndex = 1                           ' 図形インデックスを先頭に戻す
            Else
                currentShapeIndex = currentShapeIndex + 1       ' 図形インデックスの値を１つ増やす
            End If

        ' １つ前の図形を検索する
        Case SearchMode.prevShape
            If currentShapeIndex = 1 Then                      ' 図形インデックスが先頭のとき
                currentShapeIndex = searchedShapes.Count        ' 図形インデックスを最後の値にする
            Else
                currentShapeIndex = currentShapeIndex - 1       ' 図形インデックスの値を１つ減らす
            End If

    ' 上記以外
        Case Else
            ' エラー処理
            MsgBox "検索対象の図形が不正です。"
            Exit Sub
    End Select

    lblShapesNum.Caption = currentShapeIndex & " / " & searchedShapes.Count       ' 図形数表示
    Exit Sub

ErrorSearchNextShape:
    MsgBox "検索対象の図形の変更でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' 文字列を置換する
' 内容：検索文字列を置換文字列に置換する
' 引数1：searchText 検索文字列
' 引数2：replaceText 置換文字列
'------------------------------------------------------------------------------------------------------------------------
Sub ReplaceShapeText(searchText As String, replaceText As String)

    On Error GoTo ErrorReplaceShapeText
    searchedShapes(currentShapeIndex).TextFrame.Characters.Text = Replace(searchedShapes(currentShapeIndex).TextFrame.Characters.Text, searchText, replaceText) '文字列の置換
    Exit Sub

ErrorReplaceShapeText:
    MsgBox "文字の置換でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' すべての図形で文字列を置換する
' 内容：すべての図形で検索文字列を置換文字列に置換する
' 引数1：searchText 検索文字列
' 引数2：replaceText 置換文字列
'------------------------------------------------------------------------------------------------------------------------
Sub ReplaceAllText(searchText As String, replaceText As String)

    Dim shapeIndex As Integer    ' 図形番号

    On Error GoTo ErrorReplaceAllText

    For shapeIndex = 1 To searchedShapes.Count                      ' 検索図形コレクションの回数実行
        currentShapeIndex = shapeIndex                          ' 図形インデックスを更新
        Call ReplaceShapeText(searchText, replaceText)      ' 文字列の置換
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

    Dim searchText As String  ' 検索文字列

    On Error GoTo ErrorbtnSearchAll_Click

    searchText = txtSearchText.Text

    If searchText = "" Then
        Exit Sub
    End If

    Call SearchShapes(cmbSearchRange.ListIndex, searchText)     ' 図形検索
    If searchedShapes.Count > 0 Then
        Call ShowShape()        ' 図形に移動
        Call HighlightShapeString(searchText)     ' 文字列にハイライト付与
    End If
    Exit Sub

ErrorbtnSearchAll_Click:
    MsgBox "検索ボタンの処理でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' 次を検索ボタン押下時処理
' 内容：次を検索ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub btnSearchNextShape_Click()

    Dim searchText As String  ' 検索文字列

    On Error GoTo ErrorbtnSearchNextShape_Click

    searchText = txtSearchText.Text

    If searchText = "" Then
        Exit Sub
    End If

    If searchedShapes.Count > 0 Then
        Call ClearHighlightShape()      ' 文字列のハイライトをクリア
        Call SearchNextShape(SearchMode.nextShape)    ' 次の図形を対象に変更
        Call ShowShape()        ' 図形に移動
        Call HighlightShapeString(searchText)     ' 文字列にハイライト付与
    End If
    Exit Sub

ErrorbtnSearchNextShape_Click:
    MsgBox "次を検索ボタンの処理でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' １つ前を検索ボタン押下時処理
' 内容：１つ前を検索ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub btnSearchPrevShape_Click()

    Dim searchText As String  ' 検索文字列

    On Error GoTo ErrorbtnSearchPrevShape_Click

    searchText = txtSearchText.Text

    If searchText = "" Then
        Exit Sub
    End If

    If searchedShapes.Count > 0 Then
        Call ClearHighlightShape()      ' 文字列のハイライトをクリア
        Call SearchNextShape(SearchMode.prevShape)    ' 次の図形を対象に変更
        Call ShowShape()        ' 図形に移動
        Call HighlightShapeString(searchText)     ' 文字列にハイライト付与
    End If
    Exit Sub

ErrorbtnSearchPrevShape_Click:
    MsgBox "１つ前を検索ボタンの処理でエラーが発生しました"
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 置換ボタン押下時処理
' 内容：置換ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub btnReplaceShape_Click()

    Dim searchText As String  ' 検索文字列
    Dim replaceText As String ' 置換文字列

    On Error GoTo ErrorbtnReplaceShape_Click

    searchText = txtSearchText.Text
    replaceText = txtReplaceText.Text

    If searchText = "" Then
        Exit Sub
    End If

    If searchedShapes.Count > 0 Then
        Call ClearHighlightShape()      ' 文字列のハイライトをクリア
        Call ReplaceShapeText(searchText, replaceText)  ' 文字列を置換
        Call HighlightShapeString(searchText)     ' 文字列にハイライト付与
    End If
    Exit Sub

ErrorbtnReplaceShape_Click:
    MsgBox "置換ボタンの処理でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' まとめて置換ボタン押下時処理
' 内容：まとめて置換ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub btnReplaceAll_Click()

    Dim searchText As String  ' 検索文字列
    Dim replaceText As String ' 置換文字列

    On Error GoTo ErrorbtnReplaceAll_Click

    searchText = txtSearchText.Text
    replaceText = txtReplaceText.Text

    If searchText = "" Then
        Exit Sub
    End If

    Call SearchShapes(cmbSearchRange.ListIndex, searchText)     ' 図形検索
    If searchedShapes.Count > 0 Then
        Call ClearHighlightShape()      ' 文字列のハイライトをクリア
        Call ReplaceAllText(searchText, replaceText)  ' 文字列を置換
        Call isEnableButton(True)               ' すべて検索ボタンを有効化
    End If
    Exit Sub

ErrorbtnReplaceAll_Click:
    MsgBox "まとめて置換ボタンの処理でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' ユーザーフォームを閉じるときの処理
' 内容：ユーザーフォームの終了ボタン押下時の処理
'------------------------------------------------------------------------------------------------------------------------
Private Sub FormObjectSearch_QueryClose()

    On Error GoTo ErrorFormObjectSearch_QueryClose
    Call ClearHighlightShape()      ' 文字列のハイライトをクリア
    Exit Sub

ErrorFormObjectSearch_QueryClose:
    MsgBox "終了処理でエラーが発生しました"
End Sub


'------------------------------------------------------------------------------------------------------------------------
' ユーザーフォームの初期化処理
' 内容：ユーザーフォームの初期化処理
'------------------------------------------------------------------------------------------------------------------------
Sub FormObjectSearch_Initialize()

    On Error GoTo ErrorFormObjectSearch_Initialize
    Call isEnableButton(True)               ' すべて検索ボタンを有効化
    With cmbSearchRange
        .AddItem "シート"
        .AddItem "ブック"
        .ListIndex = 0                      ' インデックス0番を初期値とする
    End With
    Exit Sub

ErrorFormObjectSearch_Initialize:
    MsgBox "初期化処理でエラーが発生しました"
End Sub

'------------------------------------------------------------------------------------------------------------------------
' ボタンの有効/無効切替
' 内容：ボタンの有効/無効を切り替える
' 引数：searchAllEnable すべて検索ボタンの有効/無効 
'------------------------------------------------------------------------------------------------------------------------
Sub isEnableButton(searchAllEnable As Boolean)

    On Error GoTo ErrorisEnableButton
    If searchAllEnable = True Then
        btnSearchAll.Enabled = True             ' すべて検索ボタンを有効化
        btnSearchNextShape.Enabled = False      ' 次を検索ボタンを無効化
        btnSearchPrevShape.Enabled = False      ' 前を検索ボタンを無効化
        btnReplaceShape.Enabled = False         ' 置換ボタンを無効化
    Else
        btnSearchAll.Enabled = False            ' すべて検索ボタンを無効化
        btnSearchNextShape.Enabled = True       ' 次を検索ボタンを有効化
        btnSearchPrevShape.Enabled = True       ' 前を検索ボタンを有効化
        btnReplaceShape.Enabled = True          ' 置換ボタンを有効化
    Endif
    Exit Sub

ErrorisEnableButton:
    MsgBox "ボタンの有効/無効切替でエラーが発生しました"
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 検索文字列変更時処理
' 内容：検索文字列が変更された時、ボタンの有効/無効を切り替える
'------------------------------------------------------------------------------------------------------------------------
Sub txtSearchText_Change()
    On Error GoTo ErrortxtSearchText_Change
    Call isEnableButton(True)               ' すべて検索ボタンを有効化
    Exit Sub

ErrortxtSearchText_Change:
    MsgBox "検索文字列変更時処理でエラーが発生しました"
End Sub

'------------------------------------------------------------------------------------------------------------------------
' ユーザーフォーム起動時イベント
' 内容：ユーザーフォームが立ち上がったタイミングで発生するイベント
'------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Call FormObjectSearch_Initialize
End Sub

'------------------------------------------------------------------------------------------------------------------------
' ユーザーフォーム終了時イベント
' 内容：ユーザーフォームを終了するタイミングで発生するイベント
'------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call FormObjectSearch_QueryClose
End Sub