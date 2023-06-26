Sub テスト()
    Dim wb As Workbook
    Dim wbPath As String
    Dim sheet As Worksheet

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
        
        ' 行走査
        Dim j As Long
        For j = 2 To sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
            sheet.Cells(j, 2) = "banana" & j
            
        Next
Continue:
    Next
    
    MsgBox "ファイル読み込み完了"
    wb.Save
    Call wb.Close
Fin:
End Sub
