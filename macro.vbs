Sub bookToMD()
    ' UTF8�g�����߂ɂ��낢�낷��
    Dim Stream As Object
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Charset = "UTF-8"
    Stream.Type = 2 '2=�e�L�X�g�`��
    Stream.Open

    '�ϐ��錾
    Dim filePath As String
    Dim maxRow As Long
    Dim fileNo As Integer

    '�����l�ݒ�
    filePath = ActiveWorkbook.Path & "\rowdata.md"
    maxRow = Range("A1").End(xlDown).Row '�ŏI�s�擾
    fileNo = FreeFile 'FreeFile�֐��Ŏg�p�\�ȃt�@�C���ԍ��擾

    Stream.WriteText "|No.|Date|Title|" & vbCrLf
    Stream.WriteText "|:---|:---|:---|" & vbCrLf

    For i = 1 To maxRow
        Stream.WriteText "|" & i & "|" & Cells(i, 1) & "|" & Cells(i, 2) & "|"  & vbCrLf
    Next i

    Stream.SaveToFile (filePath), 2 '2=�㏑���ۑ�
    Stream.Close
    Set Stream = Nothing
End Sub

