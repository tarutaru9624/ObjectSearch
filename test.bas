Option Explicit

Dim searchedShapes As Collection
Dim currentShapeIndex As Integer

Sub SearchShapes()
    Dim searchText As String
    Dim searchRange As String
    Dim shape As Shape
    Dim shp As Shape
    Dim i As Integer
    Dim totalFound As Integer
    
    ' ポップアップで検索文字を取得
    searchText = InputBox("検索したい文字列を入力してください:", "文字列検索")
    If searchText = "" Then Exit Sub
    
    ' ポップアップで検索範囲を選択
    searchRange = InputBox("検索範囲を選択してください (シート内またはブック全体):", "検索範囲選択", "シート内")
    
    ' 検索範囲に応じてシートまたはブック全体を設定
    If LCase(searchRange) = "シート内" Then
        Set searchedShapes = New Collection
        For Each shp In ActiveSheet.Shapes
            If InStr(1, shp.TextFrame.Characters.Text, searchText, vbTextCompare) > 0 Then
                searchedShapes.Add shp
            End If
        Next shp
    ElseIf LCase(searchRange) = "ブック全体" Then
        Set searchedShapes = New Collection
        For Each shp In ThisWorkbook.Sheets
            For Each shape In shp.Shapes
                If InStr(1, shape.TextFrame.Characters.Text, searchText, vbTextCompare) > 0 Then
                    searchedShapes.Add shape
                End If
            Next shape
        Next shp
    Else
        MsgBox "無効な選択です。", vbExclamation
        Exit Sub
    End If
    
    ' ヒット数を表示
    totalFound = searchedShapes.Count
    If totalFound = 0 Then
        MsgBox "検索文字が見つかりませんでした。", vbInformation
        Exit Sub
    Else
        MsgBox totalFound & " 個の図形が見つかりました。", vbInformation
        currentShapeIndex = 1
        ShowCurrentShape
    End If
End Sub

Sub ShowCurrentShape()
    Dim shp As Shape
    
    If searchedShapes.Count > 0 Then
        For Each shp In ActiveSheet.Shapes
            shp.Visible = False
        Next shp
        searchedShapes(currentShapeIndex).Visible = True
        MsgBox "見ている図形: " & currentShapeIndex & vbCrLf & "ヒットした図形の合計: " & searchedShapes.Count, vbInformation
    End If
End Sub

Sub NextShape()
    If searchedShapes.Count > 0 Then
        currentShapeIndex = currentShapeIndex + 1
        If currentShapeIndex > searchedShapes.Count Then
            currentShapeIndex = 1
        End If
        ShowCurrentShape
    End If
End Sub

Sub PreviousShape()
    If searchedShapes.Count > 0 Then
        currentShapeIndex = currentShapeIndex - 1
        If currentShapeIndex < 1 Then
            currentShapeIndex = searchedShapes.Count
        End If
        ShowCurrentShape
    End If
End Sub

Sub ReplaceText()
    Dim replaceText As String
    
    If searchedShapes.Count > 0 Then
        replaceText = InputBox("置換後の文字列を入力してください:", "文字列置換")
        If replaceText = "" Then Exit Sub
        
        searchedShapes(currentShapeIndex).TextFrame.Characters.Text = Replace(searchedShapes(currentShapeIndex).TextFrame.Characters.Text, searchText, replaceText)
        MsgBox "文字列が置換されました。", vbInformation
    End If
End Sub