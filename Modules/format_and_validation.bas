Attribute VB_Name = "format_and_validation"
Option Explicit

Public Sub convertCellsValueToDbl(ByVal Target As Range)
' -------------------------------------------------------------------------------- '
' Преобразует значения всех ячеек из диапазона target в числа с плавающей запятой.
' Значения ячеек предварительно очищаются от нечисловых символов
' -------------------------------------------------------------------------------- '

    Dim cell As Range
    Dim regexp As Object
    
    On Error GoTo ErrorHandler
    Set regexp = CreateObject("vbscript.regexp")
    
    ' Пропускаем ячеейки, данные в которых невозможно преобразовать
    On Error Resume Next
    With regexp
        .Global = True
        ' шаблон соответствует всем нечисловым символам, кроме последней точки
        ' или запятой, которая считается разделителем целой и дробной части
        .Pattern = "[^\d\.\,]+|[^\d]+(?=.*[\.\,].*$)"

        For Each cell In Target.Cells
            If Not cell.HasFormula Then
                If Not IsNumeric(cell.Value2) Then
                    ' если значение ячейки нельзя преобразовать в число, то очищаем от
                    ' нечисловых символов и заменяем десятичный разделитель на запятую
                    cell.Value2 = CDbl(Replace(.Replace(cell.Value2, vbNullString), ".", ","))
                    If (err.number <> 0) Then: cell.Value2 = 0
                Else
                    ' если значение ячейки можно преобразовать в число, то делаем это
                    cell.Value2 = CDbl(cell.Value2)
                End If
            End If
        Next cell
    End With
    
CleanExit:
    Set cell = Nothing
    Set regexp = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description & vbCrLf & vbCrLf _
            & "Не удалось создать объект vbscript.regexp. Операция не завершена."
    Resume CleanExit
End Sub

Public Function convertPriceToText(number As Double, currencyType As String) As String
    Dim rank As Single
    Dim i As Single
    Dim fraction As String
    Dim str As String, temp As String
    Dim treeDigitNumber As Long
    
    'rank = numberOfDigits(number)
    
    On Error Resume Next
    str = Mid(CStr(number), 1, InStr(1, CStr(number), ",", vbTextCompare) - 1)
    If err.number <> 0 Then
        str = CStr(number)
    Else
        fraction = Left(Mid(CStr(number), InStr(1, CStr(number), ",", vbTextCompare) + 1), 2)
    End If
    On Error GoTo ErrorHandler
    
    If Len(fraction) < 2 Then: fraction = fraction & "0"
    str = Mid(str, InStr(1, str, "-", vbTextCompare) + 1)
    rank = Len(str)
    
    i = 0
    Do
        i = i + 1
        
        If i * 3 < rank Then
            treeDigitNumber = CLng(Right(str, 3))
            str = Mid(str, 1, Len(str) - 3)
        Else
            treeDigitNumber = CLng(str)
        End If
        
        Select Case i
            Case 1: temp = convertThreeDigitsNumberToText(treeDigitNumber, "integers", currencyType) & temp
            Case 2: temp = convertThreeDigitsNumberToText(treeDigitNumber, "thousands", currencyType) & temp
            Case 3: temp = convertThreeDigitsNumberToText(treeDigitNumber, "millions", currencyType) & temp
            Case 4: temp = convertThreeDigitsNumberToText(treeDigitNumber, "billions", currencyType) & temp
        End Select
    Loop While i * 3 < rank
    
    convertPriceToText = temp & convertThreeDigitsNumberToText(CLng(fraction), "decimal", currencyType)
    
CleanExit:
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    convertPriceToText = vbNullString
    Resume CleanExit
End Function
Private Function convertThreeDigitsNumberToText(number As Long, kind As String, currencyType As String) As String
    Dim integerPart As Single
    Dim str As String
    
    str = vbNullString
    integerPart = number \ 100
    If integerPart > 0 Then
        Select Case integerPart
            Case 1: str = str + "сто "
            Case 2: str = str + "двести "
            Case 3: str = str + "триста "
            Case 4: str = str + "четыреста "
            Case 5: str = str + "пятьсот "
            Case 6: str = str + "шестьсот "
            Case 7: str = str + "семьсот "
            Case 8: str = str + "восемьсот "
            Case 9: str = str + "девятьсот "
            Case Else: str = "странно :-/"
        End Select
    End If
    
    integerPart = (number Mod 100) \ 10
    If integerPart > 0 Then
        Select Case integerPart
            Case 1
                Select Case number Mod 100
                    Case 10: str = str + "десять "
                    Case 11: str = str + "одиннадцать "
                    Case 12: str = str + "двенадцать "
                    Case 13: str = str + "тринадцать "
                    Case 14: str = str + "четырнадцать "
                    Case 15: str = str + "пятнадцать "
                    Case 16: str = str + "шестнадцать "
                    Case 17: str = str + "семнадцать "
                    Case 18: str = str + "восемнадцать "
                    Case 19: str = str + "девятнадцать "
                    Case Else: str = "странно :-/"
                End Select
            Case 2: str = str + "двадцать "
            Case 3: str = str + "тридцать "
            Case 4: str = str + "сорок "
            Case 5: str = str + "пятьдесят "
            Case 6: str = str + "шестьдесят "
            Case 7: str = str + "семьдесят "
            Case 8: str = str + "восемьдесят "
            Case 9: str = str + "девяносто "
            Case Else: str = "странно :-/"
        End Select
    End If
    
    integerPart = number Mod 10
    If integerPart > 0 And ((number Mod 100) \ 10 <> 1) Then
        If kind = "thousands" Or kind = "decimal" Then
            Select Case integerPart
                Case 1: str = str + "одна "
                Case 2: str = str + "две "
            End Select
        Else
            Select Case integerPart
                Case 1: str = str + "один "
                Case 2: str = str + "два "
            End Select
        End If
        
        Select Case integerPart
            Case 3: str = str + "три "
            Case 4: str = str + "четыре "
            Case 5: str = str + "пять "
            Case 6: str = str + "шесть "
            Case 7: str = str + "семь "
            Case 8: str = str + "восемь "
            Case 9: str = str + "девять "
        End Select
    ElseIf number = 0 Then
        str = "ноль "
    End If
    
    Select Case kind
        Case "decimal"
            Select Case Application.Match(currencyType, Application.Range(CURRENCIES_ARRAY_NAME), 0)
                Case 2
                    Select Case True
                        Case number Mod 10 Like "[056789]" Or number Mod 100 Like "*1#"
                            str = str + "копеек"
                        Case number Mod 10 Like "[234]"
                            str = str + "копейки"
                        Case number Mod 10 = 1
                            str = str + "копейка"
                    End Select
                Case Else
                    Select Case True
                        Case number Mod 10 Like "[056789]" Or number Mod 100 Like "*1#"
                            str = str + "центов"
                        Case number Mod 10 Like "[234]"
                            str = str + "цента"
                        Case number Mod 10 = 1
                            str = str + "цент"
                    End Select
            End Select
        Case "integers"
            Select Case Application.Match(currencyType, Application.Range(CURRENCIES_ARRAY_NAME), 0)
                Case 1
                    str = str + "евро "
                Case 2
                    Select Case True
                        Case number Mod 10 Like "[056789]" Or number Mod 100 Like "*1#"
                            str = str + "рублей "
                        Case number Mod 10 Like "[234]"
                            str = str + "рубля "
                        Case number Mod 10 = 1
                            str = str + "рубль "
                    End Select
                Case 3
                    Select Case True
                        Case number Mod 10 Like "[056789]" Or number Mod 100 Like "*1#"
                            str = str + "долларов США "
                        Case number Mod 10 Like "[234]"
                            str = str + "доллара США "
                        Case number Mod 10 = 1
                            str = str + "доллар США "
                    End Select
            End Select
        Case "thousands"
            Select Case True
                Case number Mod 10 Like "[056789]" Or number Mod 100 Like "*1#"
                    str = str + "тысяч "
                Case number Mod 10 Like "[234]"
                    str = str + "тысячи "
                Case number Mod 10 = 1
                    str = str + "тысяча "
            End Select
        Case "millions"
            Select Case True
                Case number Mod 10 Like "[056789]" Or number Mod 100 Like "*1#"
                    str = str + "миллионов "
                Case number Mod 10 Like "[234]"
                    str = str + "миллиона "
                Case number Mod 10 = 1
                    str = str + "миллион "
            End Select
        Case "billions"
            Select Case True
                Case number Mod 10 Like "[056789]" Or number Mod 100 Like "*1#"
                    str = str + "миллиардов "
                Case number Mod 10 Like "[234]"
                    str = str + "миллиарда "
                Case number Mod 10 = 1
                    str = str + "миллиард "
            End Select
    End Select

    convertThreeDigitsNumberToText = str
End Function

Private Function numberOfDigits(number As Double) As Single
' возвращает количество целых разрядов
    On Error Resume Next
    numberOfDigits = Len(Mid(CStr(number), 1, InStr(1, CStr(number), ",", vbTextCompare) - 1))
    If err.number <> 0 Then: numberOfDigits = Len(CStr(number))
    On Error GoTo 0
    If InStr(1, CStr(number), "-", vbTextCompare) <> 0 Then: numberOfDigits = numberOfDigits - 1
End Function


Public Sub delAllFormatConditions(ByVal Target As Range)
' -------------------------------------------------------------------------------- '
' Процедура удаляет все правила условного форматирования, применяемые к диапазону
' target
' -------------------------------------------------------------------------------- '
    Dim i As Long

    ' Т.к. FormatConditions представляет из себя коллекцию, то удаление элемента
    ' в её начале будет сдвигать осташиеся элементы и изменять индексы, поэтому
    ' обходим её начиная с конца
    For i = Target.FormatConditions.Count To 1 Step -1
        Target.FormatConditions(i).Delete
    Next i
    
    Set Target = Nothing
End Sub

Public Sub createPriceFormatConditions(ByVal priceRange As Range, ByVal isect As Range)
    Dim formatRUR As String, formatEUR As String, formatUSD As String
    
    formatRUR = "# ##0,00\ [$" & ChrW(8381) & "-ru-RU]_-;[Красный]-# ##0,00\ [$" & ChrW(8381) & "-ru-RU]_-;""-""??\ [$" & ChrW(8381) & "-ru-RU]_-"
    formatEUR = "# ##0,00\ [$€-x-euro1]_-;[Красный]-# ##0,00\ [$€-x-euro1]_-;""-""??\ [$€-x-euro1]_-"
    formatUSD = "# ##0,00\ $_-;[Красный]-# ##0,00\ $_-;""-""??\ $_-"

    With priceRange
        ' Правила условного форматирования для колонок с ценами прайс-листа и цен закупки
        With .FormatConditions.Add(xlExpression, , "=" & Range(PURCHASE_TABLE_NAME) _
                              .columns(PurchaseColumns.GPL_CURRENCY).Cells(1).Address(False, False, xlR1C1, , .Cells(1)) & "=""RUR""")
            .NumberFormat = formatRUR
            .StopIfTrue = False
        End With
            
        With .FormatConditions _
            .Add(xlExpression, , "=" & Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.GPL_CURRENCY).Cells(1) _
                                                                     .Address(False, False, xlR1C1, , .Cells(1)) & "=""EUR""")
            .NumberFormat = formatEUR
            .StopIfTrue = False
        End With
            
        With .FormatConditions _
            .Add(xlExpression, , "=" & Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.GPL_CURRENCY).Cells(1) _
                                                                     .Address(False, False, xlR1C1, , .Cells(1)) & "=""USD""")
            .NumberFormat = formatUSD
            .StopIfTrue = False
        End With
            
        ' Меняем форматирование ячейки на случай, если пользователь вставил ячейки с сохранением форматов из
        ' неконтроллируемого источника
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
            
        With isect
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            
            With .Font
                .name = "Calibri"
                .Size = 10
                .Color = RGB(0, 0, 0)
                .Italic = False
                .Underline = xlUnderlineStyleNone
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With ' .Font
        End With ' isect
    End With ' priceRange
    
    Set priceRange = Nothing
    Set isect = Nothing
End Sub

Public Sub createProfitFormatCondition(ByVal Target As Range)
    With Target.FormatConditions.AddColorScale(xlColorScale)
        .ColorScaleCriteria(1).Type = xlConditionValueNumber
        .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 80, 80)
        .ColorScaleCriteria(1).Value = 0
        
        .ColorScaleCriteria(2).Type = xlConditionValueNumber
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 153)
        .ColorScaleCriteria(2).Value = 0.1
        
        .ColorScaleCriteria(3).Type = xlConditionValueNumber
        .ColorScaleCriteria(3).FormatColor.Color = RGB(51, 153, 102)
        .ColorScaleCriteria(3).Value = 0.3

    End With

    Set Target = Nothing
End Sub

Public Sub createFormatCondition(ByVal Target As Range, typename As String)
    With Target.FormatConditions
        Select Case typename
            Case "profit_type"
                .Add Type:=xlTextString, TextOperator:=xlContains, String:=Range(PROFIT_TYPE_ARRAY_NAME).Cells(1).Value2
                .Add Type:=xlTextString, TextOperator:=xlContains, String:=Range(PROFIT_TYPE_ARRAY_NAME).Cells(2).Value2
                
                .Item(1).Interior.Color = RGB(240, 255, 240)
                .Item(2).Interior.Color = RGB(240, 240, 255)
                
            Case "calc_source"
                .Add Type:=xlTextString, TextOperator:=xlContains, String:=Range(CALC_SOURCE_ARRAY_NAME).Cells(1).Value2
                .Add Type:=xlTextString, TextOperator:=xlContains, String:=Range(CALC_SOURCE_ARRAY_NAME).Cells(2).Value2
                
                .Item(1).Interior.Color = RGB(240, 255, 240)
                .Item(2).Interior.Color = RGB(240, 240, 255)
        
        End Select
    End With

    Set Target = Nothing
End Sub

Private Sub createValidation(ByVal Target As Range, typename As String)
    On Error GoTo ErrorHandler
    With Target.Validation
        .Delete
        Select Case typename
            Case "profit_type": .Add Type:=xlValidateList, Formula1:="=" & PROFIT_TYPE_ARRAY_NAME
            Case "calc_source": .Add Type:=xlValidateList, Formula1:="=" & CALC_SOURCE_ARRAY_NAME
        End Select
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
CleanExit:
    Set Target = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    'changeUpdatingState True
    Resume CleanExit
End Sub

Public Sub trimAndClearRange(ByVal Target As Range)
' -------------------------------------------------------------------------------- '
' Процедура очищает значения ячеек от двойных пробелов и непечатаемых символов
' с помощью встроенных функций ПЕЧСИМВ(CLEAN) и СЖПРОБЕЛЫ (TRIM)
' -------------------------------------------------------------------------------- '
    Dim cell As Range

    For Each cell In Target.Cells
        cell.Value2 = Application.WorksheetFunction.Clean(cell.Value2)
        cell.Value2 = Application.WorksheetFunction.Trim(cell.Value2)
    Next cell
    
    Set Target = Nothing
    Set cell = Nothing
End Sub

Public Function replaceCyrillicWithLatin(ByVal Target As Range) As Long
' -------------------------------------------------------------------------------- '
' Функция ищет кириллические символы в значениях ячеек диапазона target, меняет
' их на похожие по начертанию символы латинского алфавита и возвращаем количество
' произведённых замен
' -------------------------------------------------------------------------------- '
    Dim cell As Range
    Dim i As Long
    Dim c1 As String, c2 As String, _
        rus As String, eng As String

    rus = "асекорхуАВСЕНКМОРТХУ"
    eng = "acekopxyABCEHKMOPTXY"
    replaceCyrillicWithLatin = 0
      
    For Each cell In Target.Cells
        For i = 1 To Len(cell.Value2)
            c1 = Mid(cell, i, 1)
            ' [строка] - поиск в наборе символов
            If c1 Like "[" & rus & "]" Then
                c2 = Mid(eng, InStr(1, rus, c1), 1)
                ' производим замену только первого найденного символа, что  позволит
                ' корректно произвести подсчёт всех замен, а не только замен уникальных
                ' символов
                cell.Value2 = Replace(cell, c1, c2, , 1)
             
                replaceCyrillicWithLatin = replaceCyrillicWithLatin + 1
            End If
        Next i
    Next cell
    
    Set Target = Nothing
    Set cell = Nothing
End Function

Public Function replaceLatinWithCyrillic(ByVal Target As Range) As Long
' -------------------------------------------------------------------------------- '
' Функция ищет символы латинского алфавита в значениях ячеек диапазона target, меняет
' их на похожие по начертанию символы кириллицы и возвращаем количество
' произведённых замен
' -------------------------------------------------------------------------------- '
    Dim cell As Range
    Dim i As Long
    Dim c1 As String, c2 As String, _
        rus As String, eng As String

    rus = "асекорхуАВСЕНКМОРТХУ"
    eng = "acekopxyABCEHKMOPTXY"
    replaceLatinWithCyrillic = 0
  
    For Each cell In Target.Cells
        For i = 1 To Len(cell.Value2)
            c1 = Mid(cell, i, 1)
            ' [строка] - поиск в наборе символов
            If c1 Like "[" & eng & "]" Then
                c2 = Mid(rus, InStr(1, eng, c1), 1)
                ' производим замену только первого найденного символа, что  позволит
                ' корректно произвести подсчёт всех замен, а не только замен уникальных
                ' символов
                cell.Value2 = Replace(cell, c1, c2, , 1)
         
                replaceLatinWithCyrillic = replaceLatinWithCyrillic + 1
            End If
        Next i
    Next cell
    
    Set Target = Nothing
    Set cell = Nothing
End Function

Public Sub hideSalesColumns()
    Dim shape As shape
    Dim column As Long
    
    'Application.EnableEvents = False

    For Each shape In Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_GROUP_NAME).GroupItems
        If shape.FormControlType = xlCheckBox Then
            column = findColNumber(getDesiredColumns(), shape.AlternativeText) + COLUMN_OFFSET
            
            With Application.Worksheets(SALES_SHEET_NAME).Cells.columns(column)
                If shape.OLEFormat.Object.Value = xlOff Then
                    If .Hidden = False Then: .Hidden = True
                ElseIf shape.OLEFormat.Object.Value = xlOn Then
                    If .Hidden = True Then: .Hidden = False
                End If
            End With
        End If
    Next shape
    
    Set shape = Nothing
End Sub

Public Sub formatRangeAsType(ByVal Target As Range, Optional typename As String = "basic")
' -------------------------------------------------------------------------------- '
' Форматирует ячейки в соответствии с переданным процедуре типом. Если значение
' типа не передано, то ячейки форматируются по умолчанию
' -------------------------------------------------------------------------------- '
    Select Case typename
        Case "basic"
            formatRangeBasic Target
        
        Case "wo-zeros"
            formatRangeBasic Target
            formatAsTextWithoutZeros Target
        
        Case "subgroup"
            ' создаём подгруппу без строк подзаголовка и подытогов
            Target.Offset(1).Resize(Target.Rows.Count - 2).EntireRow.OutlineLevel = INDEX_RANK_QTY - 1
            
            formatRangeBasic Target.Rows(1)
            formatRangeBasic Target.Rows(Target.Rows.Count)
            formatAsPrice Target.Rows(Target.Rows.Count)
            Target.Rows(1).Font.Bold = True
            Target.Rows(Target.Rows.Count).Font.Bold = True
            Target.Rows(Target.Rows.Count).HorizontalAlignment = xlRight
        
        Case "kit"
            ' создаём подгруппу без строки подзаголовка
            Target.Offset(1).Resize(Target.Rows.Count - 1).EntireRow.OutlineLevel = INDEX_RANK_QTY
            
            ' по умолчанию считается, что имя сборки всегда идёт сразу после колонки
            ' с индексами
            formatRangeBasic Target(1, 2)
            Target(1, 2).Font.Bold = True
            'Stop
            
        Case "price"
            formatRangeBasic Target
            Target.HorizontalAlignment = xlRight
            formatAsPrice Target
        
        Case "percent"
            formatRangeBasic Target
            Target.HorizontalAlignment = xlRight
            formatAsPercent Target
            
        Case "center"
            formatRangeBasic Target
            Target.HorizontalAlignment = xlCenter
            
        Case "header"
            formatRangeBasic Target
            Target.HorizontalAlignment = xlCenter
            Target.Font.Bold = True
            
        Case "footer"
            formatRangeBasic Target
            Target.HorizontalAlignment = xlRight
            Target.Font.Bold = True
            formatAsPrice Target
        
        Case "profit_type"
            formatRangeBasic Target
            Target.HorizontalAlignment = xlCenter
            createValidation Target, typename
            createFormatCondition Target, typename
            
        Case "calc_source"
            formatRangeBasic Target
            Target.HorizontalAlignment = xlCenter
            createValidation Target, typename
            createFormatCondition Target, typename
    End Select
    
    Set Target = Nothing
End Sub

Private Sub formatRangeBasic(ByVal Target As Range)
' -------------------------------------------------------------------------------- '
' Форматирование ячеек коммерческого предложения с параметрами по умолчанию
' -------------------------------------------------------------------------------- '
    With Target
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        With .Font
            .name = DEFAULT_FONT
            .Size = 10
            .Color = RGB(0, 0, 0)
            .Italic = False
            .Underline = xlUnderlineStyleNone
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    Set Target = Nothing
End Sub

Private Sub formatAsTextWithoutZeros(ByVal Target As Range)
    Target.NumberFormat = "0;;;@"
    
    Set Target = Nothing
End Sub

Private Sub formatAsPrice(ByVal Target As Range)
    Target.NumberFormatLocal = "# ##0,00;[Красный]-# ##0,00;""-""??\ "
    
    Set Target = Nothing
End Sub

Private Sub formatAsPercent(ByVal Target As Range)
    Target.NumberFormatLocal = "0,00%"
    
    Set Target = Nothing
End Sub

Public Sub adjustingSalesRange(ByVal Target As Range, desiredColumns As Collection)
    Dim i As Long
    
    For i = 1 To desiredColumns.Count
        With Target.columns(i)
            Select Case desiredColumns.Item(i)
                Case SalesColumns.INDEX_NUMBER: .ColumnWidth = 9
                Case SalesColumns.MANUFACTURER: .ColumnWidth = 15
                Case SalesColumns.PN: .ColumnWidth = 15
                Case SalesColumns.NAME_AND_DESCRIPTION: .ColumnWidth = 55
                Case SalesColumns.QTY: .ColumnWidth = 6
                Case SalesColumns.Unit: .ColumnWidth = 7
                Case SalesColumns.Price: .ColumnWidth = 12
                Case SalesColumns.total: .ColumnWidth = 12
                Case SalesColumns.VAT: .ColumnWidth = 12
                Case SalesColumns.DELIVERY_TIME: .ColumnWidth = 12
                Case SalesColumns.BLANK: .ColumnWidth = 5
                Case SalesColumns.Row: .ColumnWidth = 6
                Case SalesColumns.PROFIT_TYPE: .ColumnWidth = 10
                Case SalesColumns.CALC_SOURCE: .ColumnWidth = 10
                Case SalesColumns.PROFIT: .ColumnWidth = 10
                Case Else: .ColumnWidth = 10
            End Select
        End With
    Next i
    
    Target.EntireRow.AutoFit
    Target.EntireColumn.AutoFit
    
    For i = 1 To desiredColumns.Count
        With Target.columns(i)
            Select Case desiredColumns.Item(i)
                Case SalesColumns.Row: .Hidden = True
                Case SalesColumns.PROFIT_TYPE: .ColumnWidth = 9
            End Select
        End With
    Next i
    
End Sub

Public Sub resetFormulasInPurchaseTable(Optional column As Long = 0)
    Dim cell1 As String
    Dim cell2 As String
    Dim cell3 As String
    Dim cell4 As String
    Dim str As String
    
    With Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME)
        Select Case column
            Case PurchaseColumns.PRICE_GPL_RECALCULATED
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.GPL_CURRENCY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_GPL) & "]]"
                cell3 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.VAT) & "]]"
                cell4 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.VAT_PURCHASE) & "]]"
                
                .columns(PurchaseColumns.PRICE_GPL_RECALCULATED).FormulaR1C1 = "=" & cell2 & "*" & "IFERROR(INDEX(" & _
                        CALC_CURRENCIES_ARRAY_NAME & ",MATCH(" & CALC_CURRENCY_CELL_NAME & "," & _
                        CURRENCIES_ARRAY_NAME & ",0),MATCH(" & cell1 & "," & CURRENCIES_ARRAY_NAME & ",0)),0)" & "*" & _
                        "IFERROR(INDEX(" & CALC_VAT_ARRAY_NAME & ",MATCH(" & cell4 & "," & _
                        VAT_ARRAY_NAME & ",0),MATCH(" & cell3 & "," & VAT_ARRAY_NAME & ",0)),0)"
            Case PurchaseColumns.TOTAL_GPL_RECALCULATED
                cell1 = "RC[" & CStr(PurchaseColumns.QTY - PurchaseColumns.TOTAL_GPL_RECALCULATED) & "]"
                cell2 = "RC[" & CStr(PurchaseColumns.PRICE_GPL_RECALCULATED - PurchaseColumns.TOTAL_GPL_RECALCULATED) & "]"
                
                .columns(column).FormulaR1C1 = "=" & cell1 & "*" & cell2
            Case PurchaseColumns.DISCOUNT
                cell1 = "RC[" & CStr(PurchaseColumns.PRICE_GPL_RECALCULATED - PurchaseColumns.DISCOUNT) & "]"
                cell2 = "RC[" & CStr(PurchaseColumns.PRICE_PURCHASE_RECALCULATED - PurchaseColumns.DISCOUNT) & "]"
            
                .columns(column).FormulaR1C1 = "=IFERROR((" & cell1 & "-" & cell2 & ")/" & cell1 & ","""")"
            Case PurchaseColumns.PRICE_PURCHASE_RECALCULATED
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PURCHASE_CURRENCY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_PURCHASE) & "]]"
                cell3 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.VAT) & "]]"
                cell4 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.VAT_PURCHASE) & "]]"
                
                .columns(PurchaseColumns.PRICE_PURCHASE_RECALCULATED).FormulaR1C1 = "=" & cell2 & "*" & "IFERROR(INDEX(" & _
                        CALC_CURRENCIES_ARRAY_NAME & ",MATCH(" & CALC_CURRENCY_CELL_NAME & "," & _
                        CURRENCIES_ARRAY_NAME & ",0),MATCH(" & cell1 & "," & CURRENCIES_ARRAY_NAME & ",0)),0)" & "*" & _
                        "IFERROR(INDEX(" & CALC_VAT_ARRAY_NAME & ",MATCH(" & cell4 & "," & _
                        VAT_ARRAY_NAME & ",0),MATCH(" & cell3 & "," & VAT_ARRAY_NAME & ",0)),0)"
                    
            Case PurchaseColumns.TOTAL_PURCHASE_RECALCULATED
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_PURCHASE_RECALCULATED) & "]]"
                
                .columns(column).FormulaR1C1 = "=" & cell1 & "*" & cell2
            Case PurchaseColumns.TOTAL_WEIGHT
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.UNIT_WEIGHT) & "]]"
                
                .columns(column).FormulaR1C1 = "=" & cell1 & "*" & cell2
            Case PurchaseColumns.TOTAL_VOLUME
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.UNIT_VOLUME) & "]]"
                
                .columns(column).FormulaR1C1 = "=" & cell1 & "*" & cell2
            Case PurchaseColumns.TOTAL_GPL
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_GPL) & "]]"
                
                .columns(column).FormulaR1C1 = "=" & cell1 & "*" & cell2
            Case PurchaseColumns.TOTAL_PURCHASE
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_PURCHASE) & "]]"
                
                .columns(column).FormulaR1C1 = "=" & cell1 & "*" & cell2
            Case 0
                ' PurchaseColumns.PRICE_GPL_RECALCULATED
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.GPL_CURRENCY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_GPL) & "]]"
                cell3 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.VAT) & "]]"
                cell4 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.VAT_PURCHASE) & "]]"
                
                .columns(PurchaseColumns.PRICE_GPL_RECALCULATED).FormulaR1C1 = "=" & cell2 & "*" & "IFERROR(INDEX(" & _
                        CALC_CURRENCIES_ARRAY_NAME & ",MATCH(" & CALC_CURRENCY_CELL_NAME & "," & _
                        CURRENCIES_ARRAY_NAME & ",0),MATCH(" & cell1 & "," & CURRENCIES_ARRAY_NAME & ",0)),0)" & "*" & _
                        "IFERROR(INDEX(" & CALC_VAT_ARRAY_NAME & ",MATCH(" & cell4 & "," & _
                        VAT_ARRAY_NAME & ",0),MATCH(" & cell3 & "," & VAT_ARRAY_NAME & ",0)),0)"
                
                ' PurchaseColumns.TOTAL_GPL_RECALCULATED
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_GPL_RECALCULATED) & "]]"
                .columns(PurchaseColumns.TOTAL_GPL_RECALCULATED).FormulaR1C1 = "=" & cell1 & "*" & cell2
                
                ' PurchaseColumns.DISCOUNT
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_GPL_RECALCULATED) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_PURCHASE_RECALCULATED) & "]]"
                .columns(PurchaseColumns.DISCOUNT).FormulaR1C1 = "=IFERROR((" & cell1 & "-" & cell2 & ")/" & cell1 & ","""")"
                
                ' PurchaseColumns.PRICE_PURCHASE_RECALCULATED
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PURCHASE_CURRENCY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_PURCHASE) & "]]"
                cell3 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.VAT) & "]]"
                cell4 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.VAT_PURCHASE) & "]]"
                
                .columns(PurchaseColumns.PRICE_PURCHASE_RECALCULATED).FormulaR1C1 = "=" & cell2 & "*" & "IFERROR(INDEX(" & _
                        CALC_CURRENCIES_ARRAY_NAME & ",MATCH(" & CALC_CURRENCY_CELL_NAME & "," & _
                        CURRENCIES_ARRAY_NAME & ",0),MATCH(" & cell1 & "," & CURRENCIES_ARRAY_NAME & ",0)),0)" & "*" & _
                        "IFERROR(INDEX(" & CALC_VAT_ARRAY_NAME & ",MATCH(" & cell4 & "," & _
                        VAT_ARRAY_NAME & ",0),MATCH(" & cell3 & "," & VAT_ARRAY_NAME & ",0)),0)"
                
                ' PurchaseColumns.TOTAL_PURCHASE_RECALCULATED
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_PURCHASE_RECALCULATED) & "]]"
                .columns(PurchaseColumns.TOTAL_PURCHASE_RECALCULATED).FormulaR1C1 = "=" & cell1 & "*" & cell2
            
                ' PurchaseColumns.TOTAL_WEIGHT
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.UNIT_WEIGHT) & "]]"
                .columns(PurchaseColumns.TOTAL_WEIGHT).FormulaR1C1 = "=" & cell1 & "*" & cell2
                
                ' PurchaseColumns.TOTAL_VOLUME
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.UNIT_VOLUME) & "]]"
                .columns(PurchaseColumns.TOTAL_VOLUME).FormulaR1C1 = "=" & cell1 & "*" & cell2
                
                ' PurchaseColumns.TOTAL_GPL
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_GPL) & "]]"
                .columns(PurchaseColumns.TOTAL_GPL).FormulaR1C1 = "=" & cell1 & "*" & cell2
                
                ' PurchaseColumns.TOTAL_PURCHASE
                cell1 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.QTY) & "]]"
                cell2 = "[@[" & getTableColumnName(PURCHASE_TABLE_NAME, PurchaseColumns.PRICE_PURCHASE) & "]]"
                .columns(PurchaseColumns.TOTAL_PURCHASE).FormulaR1C1 = "=" & cell1 & "*" & cell2
        End Select
    End With
End Sub

Private Function getTableColumnName(tableName As String, columnNumber As Long) As String
    getTableColumnName = Application.Range(tableName & "[#headers]").Value2(1, columnNumber)
    'getTableColumnName = Application.ListObjects(tableName).HeaderRowRange.Value2(1, columnNumber)
    'HeaderRowRange
End Function
