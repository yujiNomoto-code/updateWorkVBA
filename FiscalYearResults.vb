Sub increaseSheet()

'Cells.EntireColumn.Hidden = False

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
        
        Range("AC3").Value = sheetName
        
        
        Debug.Print sheetName
        ActiveSheet.Copy Before:=Sheets(1)
        
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
    
If sheetNameMonth <> "4月" Then



    Dim findCell As Range
    Set findCell = Range("p5:aa5").Find(sheetNameMonth, lookAt:=xlWhole)
    
    Dim findCellAddress As Integer
    'findCellAddressOffset = findCell.Offset(0, 1).Address
    'Debug.Print findCellAddressOffset
    findCellAddress = findCell.Column
    'Debug.Print findCellAddress
    
    'Debug.Print findCellAddress
    '----------------------------------↓数値で列指定
    'Columns(2).Resize(, 3).Select
    '----------------------------------
    
    '----------------------------
    'Dim findCellClmn As Integer
    
    'findCellClmn = Split(findCellAddressOffset, "$")(1)
    'Debug.Print findCellClmn
    
    'findCellClmn = Val(findCell.Column) + 1

    Cells.EntireColumn.Hidden = False
    
    
       
    
    Dim basicNum As Integer
    Dim hiddenClmnAmount As Integer
    basicNum = 27
    '4月のカラムが'O'で数字にすれば15。年度末3月は'Z’で26。つまり各月のカラムの数値を26から引いていけば右に何行を非表示にするか行数がわかる
    hiddenClmnAmount = basicNum - findCellAddress
    Debug.Print findCellAddress
    Debug.Print hiddenClmnAmount
    
    If hiddenClmnAmount <> 0 Then
    Columns(findCellAddress + 1).Resize(, hiddenClmnAmount).Hidden = True
    End If
    
    Dim pastClmn As Integer
    pastClmn = findCellAddress - 14
    '↑14なのが意味わからん！
    'わかった。下はどこからどこ、ではなく、どこから何個右にいくか。つまり4月の際はOでカラム15。なので-14で差が1なので、カラムcから右に一個で辻褄が合う。
    Columns(3).Resize(, pastClmn).Hidden = True
    
    
    
    
    
    
    
    

        
    
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")   'オブジェクト作成　正規表現のオブジェクト
    With reg
        .Pattern = "\-[0-9]+" '雛形にマッチしているかどうか。英数字1文字[abc]?、[abc]*、[a-zA-Z0-9]、 /^[0-9]{3}\-[0-9]{4}$/
        '//メモ　コンパイル言語とスクリプト言語　VBAはコンパイル言語。正規表現を学習する上では不向きなことも。 「/   /」スラッシュはVBAでいらない。
        .IgnoreCase = True '大文字と小文字を区別するFalseか、しないTrueか設定
        .Global = False       '文字列全体を検索するTrueか、しないFalseか設定
    End With



    Dim copyPer As Range
    Set copyPer = Range("AC6:AC37")
    'パーセンテージの計算式が入った列を代入
    
    For Each c In copyPer 'セル縦列を１個ずつ回す
        Dim cNum
        cNum = Val(reg.Execute(c.FormulaR1C1)(0).Value) 'c.FormulaR1C1でｃ(AC6)の座標。つまりAC6のときはAC6の数式を取得して、設定したパターンとすり合わせる
        'RC[-○○]/RC[-1]の○○の部分を取得して、その数値をcNumに代入

        'Debug.Print reg.Execute(c.FormulaR1C1)(0)
        
        cNum = cNum + 1
        '取得したcNumに1足すことによって一つ右のカラムを取得するための数値を取得
        c.FormulaR1C1 = reg.Replace(c.FormulaR1C1, Str(cNum))
        '例えば-10だったものを-9というようにして相対参照のカラムを一つ右へ移動
    
    Next
    
    
    '
    
    Dim copyPer1 As Range
    Set copyPer1 = Range("AC41:AC75")
    For Each c In copyPer1
        'Dim cNum
        cNum = Val(reg.Execute(c.FormulaR1C1)(0).Value)
        'Debug.Print cNum
        cNum = cNum + 1
        c.FormulaR1C1 = reg.Replace(c.FormulaR1C1, Str(cNum))
        'Debug.Print c.FormulaR1C1
    
    Next
    
    Dim copyPer2 As Range
    Set copyPer2 = Range("AC77:AC82")
    For Each c In copyPer2
        'Dim cNum
        cNum = Val(reg.Execute(c.FormulaR1C1)(0).Value)
        cNum = cNum + 1
        c.FormulaR1C1 = reg.Replace(c.FormulaR1C1, Str(cNum))
    
    Next
    
    '------------------------------------------------------------------------------------------------------------
    
    Dim previous As Range
    Set previous = Range("AB6:AB37")
    '前年同月の計算式が入ったセルを挿入
    
    For Each c In previous 'セル縦列を１個ずつ回す
        Dim previous1yNum
        previous1yNum = Val(reg.Execute(c.FormulaR1C1)(0).Value) 'c.FormulaR1C1でｃ(AB6)の座標。つまりAB6のときはAB6の数式を取得して、設定したパターンとすり合わせる
        'RC[-○○]/RC[-1]の○○の部分を取得して、その数値をcNumに代入
        'Debug.Print (c.FormulaR1C1)
        'Debug.Print reg.Execute(c.FormulaR1C1)(0)
        
        previous1yNum = previous1yNum + 1
        '取得したcNumに1足すことによって一つ右のカラムを取得するための数値を取得
        c.FormulaR1C1 = reg.Replace(c.FormulaR1C1, Str(previous1yNum))
        '例えば-10だったものを-9というようにして相対参照のカラムを一つ右へ移動
    
    Next
    
    
    Dim previous1 As Range
    Set previous1 = Range("AB41:AB75")
    '前年同月の計算式が入ったセルを挿入
    
    For Each c In previous1 'セル縦列を１個ずつ回す
        'Dim previous1yNum
        previous1yNum = Val(reg.Execute(c.FormulaR1C1)(0).Value) 'c.FormulaR1C1でｃ(AB6)の座標。つまりAB6のときはAB6の数式を取得して、設定したパターンとすり合わせる
        'RC[-○○]/RC[-1]の○○の部分を取得して、その数値をcNumに代入
        'Debug.Print (c.FormulaR1C1)
        'Debug.Print reg.Execute(c.FormulaR1C1)(0)
        
        previous1yNum = previous1yNum + 1
        '取得したcNumに1足すことによって一つ右のカラムを取得するための数値を取得
        c.FormulaR1C1 = reg.Replace(c.FormulaR1C1, Str(previous1yNum))
        '例えば-10だったものを-9というようにして相対参照のカラムを一つ右へ移動
    
    Next
    
    
    
    Dim previous2 As Range
    Set previous2 = Range("AB77:AB82")
    '前年同月の計算式が入ったセルを挿入
    
    For Each c In previous2 'セル縦列を１個ずつ回す
        'Dim previous1yNum
        previous1yNum = Val(reg.Execute(c.FormulaR1C1)(0).Value) 'c.FormulaR1C1でｃ(AB6)の座標。つまりAB6のときはAB6の数式を取得して、設定したパターンとすり合わせる
        'RC[-○○]/RC[-1]の○○の部分を取得して、その数値をcNumに代入
        'Debug.Print (c.FormulaR1C1)
        'Debug.Print reg.Execute(c.FormulaR1C1)(0)
        
        previous1yNum = previous1yNum + 1
        '取得したcNumに1足すことによって一つ右のカラムを取得するための数値を取得
        c.FormulaR1C1 = reg.Replace(c.FormulaR1C1, Str(previous1yNum))
        '例えば-10だったものを-9というようにして相対参照のカラムを一つ右へ移動
    
    Next
    
    
    
    
    '--------------------------------------------------------------------------------------------------------------
    Dim reg2 '今年度累計の参照範囲を右に一つ拡張
    Set reg2 = CreateObject("VBScript.RegExp")   'オブジェクト作成　正規表現のオブジェクト
    With reg2
        .Pattern = "\:[A-Z]+\[\-[0-9]+\]" '雛形にマッチしているかどうか。英数字1文字[abc]?、[abc]*、[a-zA-Z0-9]、 /^[0-9]{3}\-[0-9]{4}$/
        '.Pattern = "\-[0-9]+" '雛形にマッチしているかどうか。英数字1文字[abc]?、[abc]*、[a-zA-Z0-9]、 /^[0-9]{3}\-[0-9]{4}$/
        '//メモ　コンパイル言語とスクリプト言語　VBAはコンパイル言語。正規表現を学習する上では不向きなことも。 「/   /」スラッシュはVBAでいらない。
        .IgnoreCase = True '大文字と小文字を区別するFalseか、しないTrueか設定
        .Global = False       '文字列全体を検索するTrueか、しないFalseか設定
    End With
    
    Dim extendingReference As Range
    Set extendingReference = Range("AD6:AD10")
    For Each c In extendingReference
        Dim clmnNum As Variant
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        Debug.Print cNum
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    
    Next
    
    
        Dim extendingReference1 As Range
    Set extendingReference = Range("AD12:AD18")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    
    Next
    
    
        Dim extendingReference2 As Range
    Set extendingReference = Range("AD24:AD36")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    
    Next
    
    
    
            Dim extendingReference3 As Range
    Set extendingReference = Range("AD41:AD45")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    Next
    
    
    
            Dim extendingReference4 As Range
    Set extendingReference = Range("AD47:AD52")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    
    Next
    
    
    Dim extendingReference5 As Range
    Set extendingReference = Range("AD58:AD74")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    
    Next
    
    
    'Dim extendingReferencePast As Range
    Set extendingReference = Range("AE6:AE10")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    Next
    
    'Dim extendingReferencePast As Range
    Set extendingReference = Range("AE12:AE18")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    Next
    
    'Dim extendingReferencePast As Range
    Set extendingReference = Range("AE24:AE36")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    Next
    
    'Dim extendingReferencePast As Range
    Set extendingReference = Range("AE41:AE45")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    Next
    
    'Dim extendingReferencePast As Range
    Set extendingReference = Range("AE47:AE52")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    Next
    
    'Dim extendingReferencePast As Range
    Set extendingReference = Range("AE58:AE74")
    For Each c In extendingReference
        clmnNum = (reg2.Execute(c.FormulaR1C1)(0))
        cNum = Val(reg.Execute(clmnNum)(0).Value)
        cNum = cNum + 1
        clmnNum = reg.Replace(clmnNum, Str(cNum))
        c.FormulaR1C1 = (reg2.Replace(c.FormulaR1C1, clmnNum))
    Next
    
    
    
    'Range("AF6").Offset(0, 1).Value = "( 0, 1)"→これは隣のセル等、相対的な位置に入力
    '---------------------------------------------------------------------------------------------------------------
    
'！！！！！アクティブシートが3月の場合(つまり4月の分を作る場合)、従来の処理と分岐して前年度実績をコピーする必要がある！！！！！

   
    '----------------------------
    'Columns(findCellClmn, basicNum).Hidden = True
    'Columns("5:7").Hidden = True
    
    'Range(findCellClmn, basicNum).EntireColumn.Hidden = True
    
    '1a
    '2b
    '3c
    '4d
    '5e
    '6f
    '7g
    '8h
    '9i
    '10j
    '11k
    '12l
    '13m
    '14n
    '15o
    '16p
    '17q
    '18r
    '19s
    '20t
    '21u
    '22v
    '23w
    '24x
    '25y
    '26z
    '
    '
    '
Else
    Debug.Print sheetNameMonth
                                    'ここに年度替わりの際の処理を書く
    Debug.Print sheetNameMonth
End If


End Sub
