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
    filePath = ActiveWorkbook.Path & "\Record_of_reading\rowdata.md"
    maxRow = Range("A1").End(xlDown).Row '最終行取得

    Stream.WriteText "|No.|like|Date|ISBN|Title|" & vbCrLf
    Stream.WriteText "|:---|:---|:---|:---|:---|" & vbCrLf

    For n = 1 To maxRow
        Dim i
        i = maxRow - n + 1
        '通し番号
        Stream.WriteText "|" & i
        '色つきの本は★をつける
        If Cells(i, 3).Font.ColorIndex <> 1 Then
            Stream.WriteText "|" & "★"
        Else
            Stream.WriteText "|" & "  "
        End If
        Stream.WriteText "|" & Cells(i, 1) & "|" & Cells(i, 2) & "|" & Cells(i, 3) & "|" & vbCrLf
    Next n

    Stream.SaveToFile (filePath), 2 '2=上書き保存
    Stream.Close
    Set Stream = Nothing
End Sub

