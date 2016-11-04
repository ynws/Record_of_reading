Sub bookToMD()
    ' UTF8使うためにいろいろする
    Dim Stream As Object
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Charset = "UTF-8"
    Stream.Type = 2 '2=テキスト形式
    Stream.Open

    '変数宣言
    Dim filePath As String
    Dim maxRow As Long

    '初期値設定
    filePath = ActiveWorkbook.Path & "\rowdata.md"
    maxRow = Range("A1").End(xlDown).Row '最終行取得

    Stream.WriteText "|No.|like|Date|Title|" & vbCrLf
    Stream.WriteText "|:---|:---|:---|:---|" & vbCrLf

    For i = 1 To maxRow
        '通し番号
        Stream.WriteText "|" & i
        '色つきの本は★をつける
        If Cells(i, 2).Font.ColorIndex <> 1 Then
            Stream.WriteText "|" & "★"
        Else
            Stream.WriteText "|" & "  "
        End If
        Stream.WriteText "|" & Cells(i, 1) & "|" & Cells(i, 2) & "|" & vbCrLf
    Next i

    Stream.SaveToFile (filePath), 2 '2=上書き保存
    Stream.Close
    Set Stream = Nothing
End Sub
