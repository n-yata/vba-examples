Option Explicit

Sub メイン()
    Dim wb As Workbook
    Dim wbPath As String
    Dim sheet As Worksheet
    Dim outPath As String

    'ファイル開く
    wbPath = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
    If wbPath = "False" Then
        GoTo Fin
    End If
    Set wb = Workbooks.Open(wbPath)
    
    'シート走査
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        Set sheet = wb.Worksheets(i)
        
        If sheet.Name = "マスタ" Then
            'マスタシートは読み取り対象外のためcontinue
            GoTo Continue
        End If
        
        Dim sql As String
        
        ' 行走査
        Dim j As Long
        For j = 2 To sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
            sql = SQL生成(sql, sheet, j)
        Next
        
        '出力ファイルパスを指定
        outPath = wb.Path & "\test_" & i & ".sql"
        
        Call ファイル出力(sql, outPath)
Continue:
    Next
    
    MsgBox "ファイル読み込み完了"
    wb.Save
    Call wb.Close
Fin:
End Sub

Function SQL生成(sql As String, sheet As Worksheet, rowIdx As Long)

    sql = sql & _
        "INSERT INTO table VALUES (" & _
        ");" & vbCrLf

    SQL生成 = sql
End Function

Sub ファイル出力(sql As String, outPath As String)
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText sql
        .SaveToFile outPath, 2
        .Close
    End With
        
End Sub
