
Sub 更新()
'
' 更新 Macro

        Dim forCheck As New RegExp
        forCheck.Pattern = "^.{2}[0-9]{1,2}年[0-9]{1,2}月$"
        'RegExpは、VBScriptに正規表現として用意されているオブジェクト
        'Patternは変数の正規表現で使用するパターンを設定
        ' increaseSheet Macro
        '

    
    sheetName = ActiveSheet.Name
    If forCheck.test(sheetName) Then
    ' "Test"メソッド。正規表現によるマッチングを行う。一致したらtrue
        Dim nowMonth As Date
        nowMonth = DateValue(sheetName)
        'DateValue("2013/6/9") ・・・ 2013/06/09みたいな
        Dim newMonth As Date
        newMonth = DateAdd("m", 1, nowMonth)
        'Debug.Print newMonth
        
        sheetName = Format(newMonth, "ggge年m月")
        sheetNameMonth = Format(newMonth, "m月")
        'Debug.Print sheetNameMonth
        
        
        
        
        Debug.Print sheetName
        ActiveSheet.Copy After:=ActiveSheet

        
        Dim i As Long
        Dim SheetsCnt As Long

        SheetsCnt = ThisWorkbook.Sheets.Count
        For i = 1 To SheetsCnt
            If Sheets(i).Name = sheetName Then
                Exit Sub        'Sub自体をイグジット！！
            End If
        Next i
        ActiveSheet.Name = sheetName
    Else
        Debug.Print sheetName
    End If
 
 
 
 
 
    
    
    'ActiveSheet.Copy after:=ActiveSheet
    
    Range("E17:O19").Copy
    ActiveSheet.Paste Range("D17")
    Range("N17").AutoFill Destination:=Range("N17:O17"), Type:=xlFillDefault
    
    Range("O18:O19").ClearContents
    Range("E30:O32").Copy
    ActiveSheet.Paste Range("D30")
    Range("N30").AutoFill Destination:=Range("N30:O30"), Type:=xlFillDefault
    Range("O31:O32").ClearContents
    
    
    
End Sub
