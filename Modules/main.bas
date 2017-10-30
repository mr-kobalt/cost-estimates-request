Attribute VB_Name = "main"
Option Explicit
Option Base 1

Sub createSalesOffer()
    Dim desiredColumns As Collection, purchColumns As Collection
    Dim salesRange As Range, salesHeader As Range, salesFooter As Range
    
    'Dim StartTime As Double
    'StartTime = Timer
    
    On Error GoTo ErrorHandler
    Application.StatusBar = "Start making sales offer..."
    changeUpdatingState False
    
    ' Активируем и чистим лист, на котором будем формировать КП
    ' Оставляем нетронутыми первые (ROW_OFFSET - 1) строк
    On Error GoTo ErrorHandler2
    Application.StatusBar = "Очистка листа""" & SALES_SHEET_NAME & """"
    Sheets(SALES_SHEET_NAME).Activate
    
    On Error GoTo ErrorHandler
    initializeShapes
    Application.ActiveSheet.[A1].Select ' deselecting all shapes which were selected before creating Sales Offer
    
    With Application.ActiveSheet
        .UsedRange.EntireRow.Delete
        .UsedRange.ClearFormats
        .UsedRange.ClearOutline
        .UsedRange.columns.Hidden = False
        .UsedRange.Rows.Hidden = False
        .columns(COLUMN_OFFSET + 1).NumberFormat = "@"
    End With
    
    Application.StatusBar = "Считываем требуемые для отображения в КП колонки"
    Set desiredColumns = getDesiredColumns() ' 0.002
    
    Application.StatusBar = "Копируем требуемые колонки из расчёта в КП и создаём новые"
    Set salesRange = makeSalesTable(desiredColumns) ' 0.4
    
    Application.StatusBar = "Удаляем из КП все строки, первая ячейка которых не содержит данных"
    Set salesRange = delEmptyRows(salesRange) ' 0.13
    
    If Not (salesRange Is Nothing) Then
        Application.StatusBar = "Добавляем шапку"
        Set salesHeader = insHeader(salesRange, desiredColumns) ' 0.05
        
        Application.StatusBar = "Парсим столбец с индексами, корректируем"
        correctIndexColumn salesRange.columns(SalesColumns.INDEX_NUMBER) ' 0.07
        
        Application.StatusBar = "Сортируем по столбцу с индексами"
        salesRange.Sort key1:=salesRange(1, SalesColumns.INDEX_NUMBER), Order1:=xlAscending ' 0.004
        
        Application.StatusBar = "Парсим таблицу КП, создаём подгруппы, сборки, форматируем"
        Set salesRange = parseSalesRange(salesRange, desiredColumns) ' 0.3
        
        Application.StatusBar = "Добавляем строку итогов"
        Set salesFooter = insFooter(salesRange, desiredColumns) ' 0.015
    
        Application.Calculation = xlCalculationAutomatic
        adjustingSalesRange salesRange, desiredColumns ' 0.6
        
        Application.StatusBar = "Прячем колонки, отмеченные чекбоксами в группе " & CHECKBOXES_GROUP_NAME
        hideSalesColumns ' 0.16
    Else
        MsgBox "Добавьте товар и проставьте индексы на листе ""Расчёт закупки"""
    End If
    
    'salesRange.Select
    
    'salesRange.EntireRow.AutoFit
    
    'salesRange.Select
    'formatRangeAsType salesRange, "basic"
    'salesRange.Rows.Ungroup
CleanExit:
    Application.StatusBar = "Убираем за собой"
    Set desiredColumns = Nothing
    Set purchColumns = Nothing
    Set salesRange = Nothing
    Set salesHeader = Nothing
    Set salesFooter = Nothing
    
    changeUpdatingState True
    Application.StatusBar = False
    'MsgBox Timer - StartTime
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
    
ErrorHandler2:
    MsgBox "Error " & err.number & ": " & err.Description & vbCrLf & vbCrLf _
            & "Не найден лист с расчётом, таблица расчёта или колонка в таблице расчёта. Операция не завершена."
    Resume CleanExit
End Sub


Public Function getDesiredColumns() As Collection
    Dim checkboxes As Collection
    Dim shape As shape
    Dim topCheckbox As shape
    Dim minTop As Single
    Dim i As Variant
    
    Set getDesiredColumns = New Collection
    Set checkboxes = New Collection
    minTop = MAXSINGLE
    
    For Each shape In Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_GROUP_NAME).GroupItems
        If shape.FormControlType = xlCheckBox Then
            checkboxes.Add shape, shape.AlternativeText
        End If
    Next shape
    
    If checkboxes.Count > 0 Then
        Do
            minTop = MAXSINGLE
            For Each shape In checkboxes
                If shape.Top < minTop Then
                    Set topCheckbox = shape
                    minTop = topCheckbox.Top
                End If
            Next shape
            checkboxes.Remove topCheckbox.AlternativeText
            
            getDesiredColumns.Add CLng(topCheckbox.AlternativeText), topCheckbox.AlternativeText
        Loop While checkboxes.Count > 0
    End If
    
    
    For i = SalesColumns.[_MIDDLE] + 1 To SalesColumns.[_LAST] - 1
        getDesiredColumns.Add i, CStr(i)
    Next i
End Function

Public Function findColNumber(columns As Collection, query As String) As Variant
' -------------------------------------------------------------------------------- '
' Ищет в коллекции колонок, которые задал пользователь, строку query
' В случае нахождения возвращает её индекс в коллекции
' В противном случае возвращает Nothing
' -------------------------------------------------------------------------------- '
    Dim i As Long
    
    For i = 1 To columns.Count
        If CStr(columns.Item(i)) = query Then
            findColNumber = i
            Exit Function
        End If
    Next i
    
    findColNumber = Null
End Function

Private Function findCheckboxName(query As String) As Variant
' -------------------------------------------------------------------------------- '
' Ищет имя чекбокса по альтернативному тексту в группе CHECKBOXES_GROUP_NAME
' на листе с расчётом продажи
' -------------------------------------------------------------------------------- '
    Dim shape As shape
    findCheckboxName = Null
    
    For Each shape In Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_GROUP_NAME).GroupItems
        If shape.FormControlType = xlCheckBox And shape.AlternativeText = query Then
            findCheckboxName = shape.OLEFormat.Object.Text
            Exit Function
        End If
    Next shape
End Function

Private Function insHeader(ByVal salesRange As Range, desiredColumns As Collection) As Range
    Dim i As Long
    Dim name As Variant
    
    salesRange.Cells(1).EntireRow.Insert
    'Set insHeader = salesRange.Resize(1, findColNumber(desiredColumns, SalesColumns.BLANK) - 1).Offset(-1)
    Set insHeader = salesRange.Resize(1).Offset(-1)
    
    For i = 1 To insHeader.Cells.Count
        name = findCheckboxName(CStr(desiredColumns.Item(i)))
        If Not IsNull(name) Then
            insHeader.Cells(i).Value2 = name
        Else
            'insHeader.Cells(i).Value2 = vbNullString
            insHeader.Cells(i).Value2 = Application.Worksheets(SERVICE_SHEET_NAME).Range(SERVICE_COLUMNS_ARRAY_NAME) _
                                                   .Cells(CDbl(desiredColumns.Item(i)) - 99).Value2
        End If

        Select Case True
            Case i = SalesColumns.Price Or i = SalesColumns.total
                insHeader.Cells(i).FormulaR1C1 = "=""" & insHeader.Cells(i).Value2 & """&IF(" & _
                                                 INCLUDE_VAT_CELL_NAME & "<>INDEX(" & VAT_ARRAY_NAME & _
                                                 ",3),"" ""&" & INCLUDE_VAT_CELL_NAME & ","""")&"", ""&INDEX(" & _
                                                 CURRENCIES_HEADER_ARRAY_NAME & ",MATCH(" & CALC_CURRENCY_CELL_NAME & _
                                                 "," & CURRENCIES_ARRAY_NAME & ",0))"
                                                 
            Case i = SalesColumns.VAT:
                insHeader.Cells(i).FormulaR1C1 = "=""" & insHeader.Cells(i).Value2 & ", ""&INDEX(" & _
                                                 CURRENCIES_HEADER_ARRAY_NAME & ",MATCH(" & _
                                                 CALC_CURRENCY_CELL_NAME & "," & CURRENCIES_ARRAY_NAME & ",0))"
        End Select
    Next
    
    Set insHeader = Union(salesRange.Resize(1, findColNumber(desiredColumns, SalesColumns.BLANK) - 1).Offset(-1), _
                          salesRange.Resize(1, findColNumber(desiredColumns, SalesColumns.[_LAST] - 1) - findColNumber(desiredColumns, SalesColumns.BLANK) - 1).Offset(-1, _
                          findColNumber(desiredColumns, SalesColumns.BLANK) + 1))
    formatRangeAsType insHeader, "header"
End Function

Private Function insFooter(ByVal salesRange As Range, desiredColumns As Collection) As Range
    Dim i As Long
    Dim columnTotal As Long
    Dim columnVAT As Long
    
    'salesRange.Rows(salesRange.Rows.Count).EntireRow.Insert
    Set insFooter = salesRange.Resize(1, findColNumber(desiredColumns, SalesColumns.BLANK) - 1).Offset(salesRange.Rows.Count)
    
    columnTotal = findColNumber(desiredColumns, SalesColumns.total)
    columnVAT = findColNumber(desiredColumns, SalesColumns.VAT)
    
    insFooter.columns(columnTotal).FormulaR1C1 = "=subtotal(9," & salesRange(1, columnTotal) _
                                                .Address(False, False, xlR1C1, , insFooter.columns(columnTotal).Cells(1)) & _
                                                ":R[-1]C)"
    insFooter.columns(columnVAT).FormulaR1C1 = insFooter.columns(columnTotal).FormulaR1C1
    
    If isExistNamedRange(REVENUE_CELL_NAME) Then
        Application.Names.Item(REVENUE_CELL_NAME).RefersTo = insFooter.columns(columnTotal)
    Else
        Application.Names.Add name:=REVENUE_CELL_NAME, RefersTo:=insFooter.columns(columnTotal)
    End If
    
    If isExistNamedRange(VAT_AMOUNT_CELL_NAME) Then
        Application.Names.Item(VAT_AMOUNT_CELL_NAME).RefersTo = insFooter.columns(columnVAT)
    Else
        Application.Names.Add name:=VAT_AMOUNT_CELL_NAME, RefersTo:=insFooter.columns(columnVAT)
    End If
    
    formatRangeAsType insFooter, "footer"
    
    With Application.Worksheets(SALES_SHEET_NAME).Range(insFooter.Cells(1), insFooter.Cells(columnTotal - 1))
        .ClearContents
        .Merge
        .FormulaR1C1 = "=""" & TEXTS_TOTAL & " ""&IF(" & INCLUDE_VAT_CELL_NAME & "<>INDEX(" & VAT_ARRAY_NAME & _
                       ",3)," & INCLUDE_VAT_CELL_NAME & ",""(" & TEXTS_NOT_SUBJECT_VAT & ")"")&"", ""&INDEX(" & _
                       CURRENCIES_HEADER_ARRAY_NAME & ",MATCH(" & CALC_CURRENCY_CELL_NAME & "," & CURRENCIES_ARRAY_NAME & _
                       ",0))&"":"""
    End With
End Function

Private Function makeSalesTable(desiredColumns As Collection)
' -------------------------------------------------------------------------------- '
' Копирует из диапазона "Расчёт" необходимые для КП колонки, а также
' добавляет новые, такие как Цена, НДС и другие.
'
' Если колонка из desiredColumns не найдена, то заполняет ячейки новой колонки #N/A
' -------------------------------------------------------------------------------- '
    Dim i As Long
    Dim newColumn As Range
    Dim columnValue As Variant
    Dim shape As shape
    Dim tempAddress1 As String, tempAddress2 As String, total As String
    
    With Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME)
        For i = 1 To desiredColumns.Count
            Set newColumn = Range(Cells(ROW_OFFSET + 1, COLUMN_OFFSET + i), _
                                  Cells(ROW_OFFSET + .Rows.Count, COLUMN_OFFSET + i))
            columnValue = desiredColumns.Item(i)
            
            Select Case True
                Case columnValue = SalesColumns.INDEX_NUMBER
                    newColumn.Value2 = .columns(columnValue).Value2
                    formatRangeAsType newColumn
                    
                Case columnValue = SalesColumns.PN
                    newColumn.Formula = "='" & PURCHASE_SHEET_NAME & "'!" & .columns(PurchaseColumns.PN).Cells(1).Address(False, True, xlA1)
                    formatRangeAsType newColumn, "wo-zeros"
                    
                Case columnValue = SalesColumns.NAME_AND_DESCRIPTION
                    newColumn.Formula = "='" & PURCHASE_SHEET_NAME & "'!" & .columns(PurchaseColumns.NAME_AND_DESCRIPTION).Cells(1).Address(False, True, xlA1)
                    formatRangeAsType newColumn, "wo-zeros"
                    
                Case columnValue = SalesColumns.QTY
                    newColumn.Formula = "='" & PURCHASE_SHEET_NAME & "'!" & .columns(PurchaseColumns.QTY).Cells(1).Address(False, True, xlA1)
                    formatRangeAsType newColumn, "center"
                
                Case columnValue = SalesColumns.Unit
                    newColumn.Formula = "='" & PURCHASE_SHEET_NAME & "'!" & .columns(PurchaseColumns.Unit).Cells(1).Address(False, True, xlA1)
                    formatRangeAsType newColumn, "center"
                    
                Case columnValue = SalesColumns.Price
                    tempAddress1 = "'" & PURCHASE_SHEET_NAME & "'!" & .columns(PurchaseColumns.PRICE_GPL_RECALCULATED). _
                                            Cells(1).Address(False, True, xlR1C1, , newColumn.Cells(1))
                    tempAddress2 = "'" & PURCHASE_SHEET_NAME & "'!" & .columns(PurchaseColumns.PRICE_PURCHASE_RECALCULATED). _
                                            Cells(1).Address(False, True, xlR1C1, , newColumn.Cells(1))
                    
                    If Application.Worksheets(SALES_SHEET_NAME).shapes(GPL_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
                        total = TOTAL_GPL_CELL_NAME
                    ElseIf Application.Worksheets(SALES_SHEET_NAME).shapes(NET_PRICE_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
                        total = TOTAL_COST_CELL_NAME
                    End If
                    
                    newColumn.FormulaR1C1 = "=ROUND(IF(INDEX(" & CALC_SOURCE_ARRAY_NAME & ",1)=RC" & _
                                            (findColNumber(desiredColumns, SalesColumns.CALC_SOURCE) + COLUMN_OFFSET) & _
                                            "," & tempAddress1 & "," & tempAddress2 & ")*IF(" & INCLUDE_DELIVERY_CELL_NAME & _
                                            "=""да""" & ",(1+" & DELIVERY_COST_CELL_NAME & "/" & total & _
                                            "*INDEX(" & CALC_CURRENCIES_ARRAY_NAME & ",MATCH(" & CALC_CURRENCY_CELL_NAME & _
                                            "," & CURRENCIES_ARRAY_NAME & ",0),2)*INDEX(" & CALC_VAT_ARRAY_NAME & _
                                            ",MATCH(" & INCLUDE_VAT_CELL_NAME & "," & VAT_ARRAY_NAME & _
                                            ",0),1)),1)*IF(INDEX(" & PROFIT_TYPE_ARRAY_NAME & ",1)=RC" & _
                                            (findColNumber(desiredColumns, SalesColumns.PROFIT_TYPE) + COLUMN_OFFSET) & _
                                            ",(1+RC" & (findColNumber(desiredColumns, SalesColumns.PROFIT) + COLUMN_OFFSET) & _
                                            "),1/(1-RC" & (findColNumber(desiredColumns, SalesColumns.PROFIT) + COLUMN_OFFSET) & _
                                            "))," & PRICE_ROUNDING_UP_TO_QTY & ")"
                    formatRangeAsType newColumn, "price"
                    
                Case columnValue = SalesColumns.total
                    newColumn.FormulaR1C1 = "=RC" & (findColNumber(desiredColumns, SalesColumns.QTY) + COLUMN_OFFSET) & _
                                            "*RC" & (findColNumber(desiredColumns, SalesColumns.Price) + COLUMN_OFFSET)
                    formatRangeAsType newColumn, "price"
                    
                Case columnValue = SalesColumns.VAT
                    newColumn.FormulaR1C1 = "=ROUND(RC" & _
                                            (findColNumber(desiredColumns, SalesColumns.total) + COLUMN_OFFSET) & _
                                            "*IF(" & INCLUDE_VAT_CELL_NAME & "=INDEX(" & VAT_ARRAY_NAME & ",1),0.18/1.18,IF(" & _
                                            INCLUDE_VAT_CELL_NAME & "=INDEX(" & VAT_ARRAY_NAME & ",2),0.18,0))," & _
                                            PRICE_ROUNDING_UP_TO_QTY & ")"

                    formatRangeAsType newColumn, "price"
                
                Case columnValue = SalesColumns.DELIVERY_TIME
                    newColumn.Formula = "='" & PURCHASE_SHEET_NAME & "'!" & .columns(PurchaseColumns.DELIVERY_TIME).Cells(1).Address(False, True, xlA1)
                    formatRangeAsType newColumn, "wo-zeros"
                
                Case columnValue = SalesColumns.BLANK
                    newColumn.Value2 = vbNullString
                
                Case columnValue = SalesColumns.Row
                    newColumn.Cells(1).Value2 = Mid(.Cells(1).Address(ReferenceStyle:=xlR1C1), 1, _
                                                    InStr(.Cells(1).Address(ReferenceStyle:=xlR1C1), "C") - 1)
                    newColumn.Cells(1).AutoFill newColumn, xlFillSeries
                    formatRangeAsType newColumn, "center"
                
                Case columnValue = SalesColumns.PROFIT_TYPE
                    If Application.Worksheets(SALES_SHEET_NAME).shapes(MARKUP_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
                        Set shape = Application.Worksheets(SALES_SHEET_NAME).shapes(MARKUP_SHAPE_NAME)
                    ElseIf Application.Worksheets(SALES_SHEET_NAME).shapes(MARGIN_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
                        Set shape = Application.Worksheets(SALES_SHEET_NAME).shapes(MARGIN_SHAPE_NAME)
                    End If
                    
                    newColumn.Value2 = shape.AlternativeText
                    formatRangeAsType newColumn, "profit_type"
                    
                Case columnValue = SalesColumns.CALC_SOURCE
                    If Application.Worksheets(SALES_SHEET_NAME).shapes(GPL_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
                        Set shape = Application.Worksheets(SALES_SHEET_NAME).shapes(GPL_SHAPE_NAME)
                    ElseIf Application.Worksheets(SALES_SHEET_NAME).shapes(NET_PRICE_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
                        Set shape = Application.Worksheets(SALES_SHEET_NAME).shapes(NET_PRICE_SHAPE_NAME)
                    End If
                    
                    newColumn.Value2 = shape.AlternativeText
                    formatRangeAsType newColumn, "calc_source"
                    
                Case columnValue = SalesColumns.PROFIT
                    newColumn.Value2 = CDbl(Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_BUTTON_SHAPE_NAME) _
                                            .OLEFormat.Object.Caption) / 100
                    formatRangeAsType newColumn, "percent"
                    createProfitFormatCondition newColumn
                    
                Case IsNumeric(columnValue) And columnValue > PurchaseColumns.[_FIRST] And columnValue < PurchaseColumns.[_LAST]
                    newColumn.Formula = "='" & PURCHASE_SHEET_NAME & "'!" & .columns(columnValue).Cells(1).Address(False, True, xlA1)
                    formatRangeAsType newColumn, "wo-zeros"
                    
                Case Else
                    newColumn.Value2 = CVErr(xlErrNA)
            End Select
        Next
    
        Set makeSalesTable = Application.Worksheets(SALES_SHEET_NAME) _
                                        .Range(Cells(ROW_OFFSET + 1, COLUMN_OFFSET + 1), _
                                               Cells(ROW_OFFSET + .Rows.Count, COLUMN_OFFSET + desiredColumns.Count))
    End With
End Function

Private Function delEmptyRows(ByVal toClearRange As Range)
' -------------------------------------------------------------------------------- '
' Удаляет все строки из диапазона, крайняя левая ячейка которых не содержит
' символов, и возвращает новый диапазон.
'
' Возвращает Nothing, если были удалены все строки в диапазоне
' -------------------------------------------------------------------------------- '
    Dim delRange As Range
    Dim isError As Boolean
    Dim rc As Long
    
    On Error Resume Next
    Set delRange = toClearRange.columns(1).SpecialCells(xlCellTypeBlanks)
    isError = (err.number <> 0)
    On Error GoTo 0
    
    If isError Then
        Set delEmptyRows = toClearRange
    Else
        rc = toClearRange.Rows.Count - delRange.Cells.Count
        delRange.EntireRow.Delete
        
        If rc <= 0 Then
            Set delEmptyRows = Nothing
        Else
            Set delEmptyRows = toClearRange.Resize(rc)
        End If
    End If
End Function

Public Sub correctIndexColumn(ByVal column As Range)
' -------------------------------------------------------------------------------- '
' корректирует колонку с индексами строк КП - очищает от лишних нулей, нецифровых
' и спецсимволов; корректируем последовательность индексов (начиная с единицы
' с инкриментом 1)
' -------------------------------------------------------------------------------- '
    Dim i As Long, j As Long
    Dim indexArray() As Variant
    Dim temp As Variant
    Dim regexp As Object
    
    ReDim indexArray(column.Cells.Count, INDEX_RANK_QTY)
    Set regexp = CreateObject("vbscript.regexp")
    
    Call correctIndexRange(column)
    For i = 1 To column.Cells.Count
        With regexp
            .Global = True
            ' Формируем двумерный массив индексов, каждая строка которого хранит индекс строки в КП -
            ' одно или несколько чисел (их количество ограничено INDEX_RANK_QTY), разделённых
            ' по столбцам массива. Если количество разрядов в индексе меньше INDEX_RANK_QTY,
            ' то отсутствующие значения заполняются пустой строкой - "" (vbNullString)
            .Pattern = "\d+"
            For j = LBound(indexArray, 2) To UBound(indexArray, 2)
                If j <= .Execute(column.Cells(i).Value2).Count Then
                    indexArray(i, j) = .Execute(column.Cells(i).Value2)(j - 1)
                Else
                    indexArray(i, j) = vbNullString
                End If
            Next j
        End With
    Next i
    
    ' исправляем индексы
    Call shrinkColumnsIndices(indexArray)
    
    ' записываем исправленные индексы обратно в колонку
    For i = 1 To column.Cells.Count
        temp = vbNullString
        For j = 1 To UBound(indexArray, 2)
            If indexArray(i, j) = vbNullString Then: Exit For ' TODO: исправить костыль с принудительным выходом их цикла
            temp = temp & indexArray(i, j) & "."
        Next j
        column.Cells(i).Value2 = Mid(temp, 1, Len(temp) - 1) ' TODO: исправить костыль с отсечением последней точки
    Next i
    
    Erase indexArray
End Sub

Public Sub correctIndexRange(ByVal indexRange As Range)
    Dim regexp As Object
    Dim temp As Variant
    Dim indexArray() As Variant
    Dim i As Long
    
    If isArrayEmpty(indexRange.Value2) Then
        ReDim indexArray(1, 1)
        indexArray(1, 1) = indexRange.Value2
    Else
        indexArray = indexRange.Value2
    End If
    
    Set regexp = CreateObject("vbscript.regexp")
    For i = LBound(indexArray, 1) To UBound(indexArray, 1)
        temp = CStr(indexArray(i, 1))
        With regexp
            .Global = True
            ' удаляем символы и незначащие нули в начале и конце строки
            .Pattern = "^([^\d]*0*)(?=\d)|[^\d]+$": temp = .Replace(temp, vbNullString)
            ' заменяем символы и незначащие нули на точки
            .Pattern = "[^\d]+0*(?=\d)":            temp = .Replace(temp, ".")
            ' оставляем в строке первые INDEX_RANK_QTY разряда чисел разделённых точками
            ' Например, в строке "13.0.87.1.12" пять разрядов; при INDEX_RANK_QTY = 3, и
            ' после преобразования, строка примет вид - "13.0.87"
            .Pattern = "(\d+\.){0," & INDEX_RANK_QTY - 1 & "}\d+"
            If .test(temp) Then temp = .Execute(temp)(0) Else temp = vbNullString
        End With
        
        indexArray(i, 1) = temp
    Next i
    
    indexRange.Cells.Value2 = indexArray
End Sub

Private Function shrinkColumnsIndices(arr() As Variant, Optional column As Long = 1, Optional rowIndices As Collection) As Variant
' -------------------------------------------------------------------------------- '
' arr() - двумерный массив, каждая строка которого хранит индекс строки в КП -
' одно или несколько чисел (их количество ограничено INDEX_RANK_QTY), разделённых
' по столбцам массива. Если количество чисел в индексе меньше INDEX_RANK_QTY,
' то отсутствующие значения должны быть заполнены пустой строкой - "" (vbNullString).
' column - столбец массива, в котором производится замена индексов в текущей итерации
' rowIndices - коллекция строковых индексов массива, в которых будет производится
' замена индексов в текущей итерации
' -------------------------------------------------------------------------------- '
' Функция рекурсивно исправляет индексы в полученном двумерном массиве, для чего:
' - ищет минимальное значение minValue индекса в текущей колонке column среди
' заданных строковых индексов rowIndices
' - находит все строковые индексы элементов массива среди заданных rowIndices,
' которые равны найденному минимальному значению
' - заменяет значения всех найденных элементов на newIndex
' - удаляет найденные индексы из коллекции rowIndices
' - рекурсивно вызывает себя, передавая в качестве аргумента массив, номер
' следующей колонки и найденные в шаге 2 индексы
' - увеличиваем значение индекса newIndex на 1 и повторяем цикл, пока rowIndices
' не пуст
' -------------------------------------------------------------------------------- '
' Пример работы функции:
'   10  3   5        3   1   1
'   8   1            1   2
'   8   0            1   1
'   9           -->  2
'   10  10  10       3   2   2
'   10  10  8        3   2   1
' -------------------------------------------------------------------------------- '
    Dim minValue As Variant         ' хранит минимальное значение в колонке column строк с индексами rowIndices
    Dim maxValue As Variant         ' хранит максимальное значение в колонке column строк с индексами rowIndices
    Dim uniqueIntValues As Variant  ' хранит кол-во уникальных значений в колонке column строк с индексами rowIndices
    Dim newIndex As Variant         ' новый индекс для текущей колонки; заменяет значение всех элементов равных
                                    ' minValue в диапазоне rowIndices
    
    Dim temp As String
    Dim i As Variant
    Dim slicedIndices As Collection ' используется для хранения индексов строк с элементами равными minValue
    
    ' если индексов не было передано в функцию, то формируем свою коллекцию индексов,
    ' включающую все индексы массива arr()
    If rowIndices Is Nothing Then
        Set rowIndices = New Collection
        
        For i = LBound(arr) To UBound(arr)
            ' каждому элементу коллекции присваиваем ключ в виде значения элемента преобразованного в строку,
            ' что позволить удалять элементы коллекции в цикле, не боясь смещения индексов
            rowIndices.Add i, CStr(i)
        Next i
    End If
    
    If column <= INDEX_RANK_QTY Then
        
        newIndex = 1 ' по умолчанию заполняем индексы целыми числами начиная с единицы
        maxValue = maxValueInColumn(arr, column, rowIndices)
        uniqueIntValues = uniqueIntValuesInColumn(arr, column, rowIndices)
        
        Do While rowIndices.Count > 0 ' пока есть необработанные элементы массива
            minValue = minValueInColumn(arr, column, rowIndices)
            ' если числового минимального значения не найдено, то обнуляем коллекцию индексов
            ' и завершаем работы функции
            If IsNull(minValue) Then
                Set rowIndices = New Collection
            Else
                Set slicedIndices = New Collection
                ' если минимальное значение индекса равно нулю и это не первый разряд индекса,
                ' то начинаем заполнять новыми индексами начиная с нуля
                If minValue = 0 And column > 1 Then: newIndex = 0
                ' находит все строковые индексы элементов массива среди заданных rowIndices,
                ' которые равны найденному минимальному значению
                For Each i In rowIndices
                    If arr(i, column) = CStr(minValue) Then
                        'temp = String(Len(CStr(maxValue)) - Len(CStr(newIndex)), "0")
                        temp = String(Len(CStr(uniqueIntValues)) - Len(CStr(newIndex)), "0")
                        
                        arr(i, column) = temp & CStr(newIndex) ' заменяет значения всех найденных элементов на newIndex
                        slicedIndices.Add i, CStr(i) ' сохраняет найденные индексы
                        rowIndices.Remove CStr(i) ' удаляет найденные индексы из коллекции rowIndices
                    End If
                Next i
                
                ' рекурсивно вызывает себя, передавая в качестве аргумента массив, номер
                ' следующей колонки и найденные индексы
                
                If slicedIndices.Count > 0 Then
                    Call shrinkColumnsIndices(arr, column + 1, slicedIndices)
                End If
                
                newIndex = newIndex + 1
            End If
        Loop
    End If
End Function

Public Function minValueInColumn(arr() As Variant, column As Long, Optional rowIndices As Collection)
' -------------------------------------------------------------------------------- '
' Обёртка функции extremumValueInColumn для поиска минимального значения в колонке
' массива
' -------------------------------------------------------------------------------- '
    minValueInColumn = extremumValueInColumn(arr, column, "<", rowIndices)
End Function

Public Function maxValueInColumn(arr() As Variant, column As Long, Optional rowIndices As Collection)
' -------------------------------------------------------------------------------- '
' Обёртка функции extremumValueInColumn для поиска максимального значения в колонке
' массива
' -------------------------------------------------------------------------------- '
    maxValueInColumn = extremumValueInColumn(arr, column, ">", rowIndices)
End Function

Private Function extremumValueInColumn(arr() As Variant, column As Long, op As String, Optional rowIndices As Collection)
' -------------------------------------------------------------------------------- '
' Ищет минимальное или максимальное значение (в зависимости от переменной op)
' в заданной колонке массива (второе измерение). Опционально возможно передать
' коллекцию индексов строк, в которых будет производиться поиск.
' Возвращает Nothing, если массив пуст или передано .
' -------------------------------------------------------------------------------- '
    Dim i As Variant
    
    If isArrayEmpty(arr) Then
        extremumValueInColumn = Null
    Else
        ' в начале цикла минимальное или максимальное значение инициализируется
        ' константами MAXLONG или MINLONG
        'extremumValueInColumn = CLng(arr(LBound(arr, 1), column))
        If op = "<" Then
            extremumValueInColumn = MAXLONG
        Else
            extremumValueInColumn = MINLONG
        End If
    
        ' Проходим либо по всем строкам столбца, либо по строкам с индексами,
        ' полученными из коллекции rowIndices
        If rowIndices Is Nothing Then
            For i = LBound(arr) To UBound(arr)
                ' Если среди значений столбца нет чисел, то будет выполнено сравнение
                ' строк.
                ' https://msdn.microsoft.com/ru-ru/library/215yacb6.aspx
                If IsNumeric(extremumValueInColumn) And IsNumeric(arr(i, column)) Then
                    If Application.Evaluate(CStr(arr(i, column)) & op & CStr(extremumValueInColumn)) Then
                        extremumValueInColumn = CLng(arr(i, column))
                    End If
                End If
            Next i
        'ElseIf rowIndices.Count = 0 Then
        '    extremumValueInColumn = Null
        Else
            If rowIndices.Count <> 0 Then
                For Each i In rowIndices
                    ' Если среди значений столбца нет чисел, то будет выполнено сравнение
                    ' строк.
                    ' https://msdn.microsoft.com/ru-ru/library/215yacb6.aspx
                    If IsNumeric(extremumValueInColumn) And IsNumeric(arr(i, column)) Then
                        If Application.Evaluate(CStr(arr(i, column)) & op & CStr(extremumValueInColumn)) Then
                            extremumValueInColumn = CLng(arr(i, column))
                        End If
                    End If
                Next i
            End If
        End If
        
        If extremumValueInColumn = MINLONG Or extremumValueInColumn = MAXLONG Then
            extremumValueInColumn = Null
        End If
    End If
End Function

Private Function uniqueIntValuesInColumn(arr() As Variant, column As Long, Optional rowIndices As Collection)
' -------------------------------------------------------------------------------- '
' Вычисляет кол-во уникальных значений в колонке column массива arr
' -------------------------------------------------------------------------------- '
    Dim i As Variant
    Dim dict As New Collection ' не словарь, но что-то похожее
    
    If isArrayEmpty(arr) Then
        uniqueIntValuesInColumn = Null
    Else
        ' Проходим либо по всем строкам столбца, либо по строкам с индексами,
        ' полученными из коллекции rowIndices
        If rowIndices Is Nothing Then
            For i = LBound(arr) To UBound(arr)
                On Error Resume Next
                If IsNumeric(arr(i, column)) Then
                    dict.Add arr(i, column), arr(i, column)
                End If
                On Error GoTo 0
            Next i
        Else
            If rowIndices.Count <> 0 Then
                For Each i In rowIndices
                    On Error Resume Next
                    If IsNumeric(arr(i, column)) Then
                        dict.Add arr(i, column), arr(i, column)
                    End If
                    On Error GoTo 0
                Next i
            End If
        End If
        
        uniqueIntValuesInColumn = dict.Count
    End If
End Function

Private Function parseSalesRange(ByVal salesRange As Range, desiredColumns As Collection) As Range
    Dim indexColumn() As Variant
    If isArrayEmpty(salesRange.columns(PurchaseColumns.INDEX_NUMBER).Value2) Then
        ReDim indexColumn(1, 1)
        indexColumn(1, 1) = salesRange.columns(PurchaseColumns.INDEX_NUMBER).Value2
    Else
        indexColumn = salesRange.columns(PurchaseColumns.INDEX_NUMBER).Value2
    End If
    
    Dim groupCount As Collection: Set groupCount = getIndexGroupCount(indexColumn)
    Dim regexp As Object: Set regexp = CreateObject("vbscript.regexp")
    Dim tempRegExp As Object
    Dim tempValue As Variant
    Dim currentGroup As String
    Dim kitRange As Range
    Dim i As Long, j As Long

    
    With regexp
        .Global = True
        .Pattern = "\d+"
        
        i = 1
        Do While i <= salesRange.Rows.Count
            Set tempRegExp = .Execute(salesRange(i, 1).Value2)
            
            '
            If tempRegExp.Count = 1 And i <= salesRange.Rows.Count Then
                If currentGroup <> tempRegExp(0) Then
                    currentGroup = tempRegExp(0)
                    For j = i + 1 To salesRange.Rows.Count
                        Set tempRegExp = .Execute(salesRange(j, 1).Value2)
                        If tempRegExp.Count > 0 Then
                            If currentGroup <> tempRegExp(0) Then: Exit For
                        End If
                    Next j
                    
                    If (j - 1) - i > 0 Then: Set salesRange = makeSubGroup(salesRange, i, j, desiredColumns)
                End If
                ElseIf tempRegExp.Count > 1 Then
                ' Преобразуем последнюю числовую группу в число. Если оно равно нулю, то
                ' делаем сборку. Пример:
                ' "123.12.000" -> "000" -> 0 = 0
                ' "123.12.010" -> "010" -> 10 != 0
                If CLng(tempRegExp(tempRegExp.Count - 1)) = 0 Then
                    ' Обрезаем нули в конце строки. Пример:
                    ' "123.12.000" -> "123.12."
                    .Pattern = "(\d+\.){1,}"
                    tempValue = .Execute(salesRange(i, 1).Value2)(0)
                    ' Обрезаем оконечную точку. Пример:
                    ' "123.12." -> "123.12"
                    salesRange(i, 1).Value2 = Mid(tempValue, 1, Len(tempValue) - 1)
                    
                    For j = i + 1 To salesRange.Rows.Count
                        Set tempRegExp = .Execute(salesRange(j, 1).Value2)
                        If tempRegExp.Count > 0 Then
                            If InStr(tempRegExp(0), tempValue) <> 1 Then: Exit For
                        Else
                            Exit For
                        End If
                    Next j
                    
                    Set kitRange = Application.Worksheets(SALES_SHEET_NAME).Range(salesRange(i, 1), salesRange(j - 1, salesRange.columns.Count))
                    makeKit kitRange, desiredColumns
                    
                    ' пропускаем все строки, включённые в сборку
                    i = j - 1
                    
                    .Pattern = "\d+"
                End If
            End If
            
            i = i + 1
        Loop
    End With
    
    Set parseSalesRange = salesRange
End Function

Private Function makeSubGroup(ByRef salesRange As Range, firstRow As Long, lastRow As Long, desiredColumns As Collection) As Range
    Dim regexp As Object: Set regexp = CreateObject("vbscript.regexp")
    Dim tempRegExp As Object
    Dim tempFormula As Variant
    Dim column As Long
    
    With regexp
        .Global = True
        .Pattern = "\d+"
        Set tempRegExp = .Execute(salesRange(firstRow, PurchaseColumns.INDEX_NUMBER).Value2)
        
        tempFormula = salesRange(firstRow, findColNumber(desiredColumns, SalesColumns.NAME_AND_DESCRIPTION)).FormulaR1C1
        
        With Application.Worksheets(SALES_SHEET_NAME).Range(salesRange(firstRow, SalesColumns.INDEX_NUMBER + 1), _
                                                            salesRange(firstRow, findColNumber(desiredColumns, SalesColumns.BLANK) - 1))
            .ClearContents
            .Merge
            .FormulaR1C1 = tempFormula
        End With
        
        salesRange.Rows(lastRow).EntireRow.Insert xlShiftDown
        
        If lastRow > salesRange.Rows.Count Then
            Set salesRange = salesRange.Resize(salesRange.Rows.Count + 1, salesRange.columns.Count)
        End If
        
        column = findColNumber(desiredColumns, SalesColumns.total)
        formatRangeAsType Range(salesRange(firstRow, 1), salesRange(lastRow, findColNumber(desiredColumns, SalesColumns.BLANK) - 1)), "subgroup"
        
        With Application.Worksheets(SALES_SHEET_NAME).Range(salesRange(lastRow, SalesColumns.INDEX_NUMBER), _
                                                            salesRange(lastRow, column - 1))
            .ClearContents
            .Merge
            .FormulaR1C1 = "=""" & TEXTS_SUBTOTAL & """&IF(" & INCLUDE_VAT_CELL_NAME & "<>INDEX(" & VAT_ARRAY_NAME & _
                           ",3),"" ""&" & INCLUDE_VAT_CELL_NAME & ","""")&"", ""&INDEX(" & CURRENCIES_HEADER_ARRAY_NAME & _
                           ",MATCH(" & CALC_CURRENCY_CELL_NAME & "," & CURRENCIES_ARRAY_NAME & ",0))" & "&"":"""
        End With
        With salesRange(lastRow, column)
            .FormulaR1C1 = "=subtotal(9," & _
                            salesRange(firstRow, column) _
                            .Address(False, False, xlR1C1, , .Cells(1)) & _
                           ":R[-1]C)"
            salesRange(lastRow, findColNumber(desiredColumns, SalesColumns.VAT)).FormulaR1C1 = .FormulaR1C1
        End With
    End With
    
    Set makeSubGroup = salesRange
End Function

Private Sub makeKit(ByVal kitRange As Range, desiredColumns As Collection)
    Dim i As Long, j As Long
    Dim column As Range
    Dim columnName As Variant, columnNumberQty As Variant, columnNumberName As Variant
    Dim tempAddress1 As String, tempAddress2 As String, tempAddress3 As String, tempFormula As String
    Dim total As String
    
    On Error GoTo ErrorHandler
    If kitRange.Rows.Count > 1 Then
        columnNumberQty = findColNumber(desiredColumns, SalesColumns.QTY)
        
        For i = 1 To findColNumber(desiredColumns, SalesColumns.BLANK) - 1
            With kitRange.columns(i)
                columnName = desiredColumns.Item(i)
                Select Case True
                    Case columnName = SalesColumns.NAME_AND_DESCRIPTION
                        Set column = kitRange.columns(i).Resize(kitRange.Rows.Count - 1).Offset(1)
                        
                        tempAddress1 = Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME) _
                                     .columns(PurchaseColumns.QTY).Cells(1).Address(False, True, xlR1C1, , column.Cells(1))
                        tempAddress2 = Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME) _
                                     .columns(PurchaseColumns.Unit).Cells(1).Address(False, True, xlR1C1, , column.Cells(1))
                        
                        For j = 1 To column.Rows.Count
                            tempFormula = column.Cells(j).FormulaR1C1
                            column.Cells(j).FormulaR1C1 = tempFormula & "&"" - ""&" & _
                                                     Mid(tempFormula, 2, InStr(tempFormula, "C") - 1) & _
                                                     Mid(tempAddress1, InStr(tempAddress1, "C") + 1) & _
                                                     "/" & kitRange(1, columnNumberQty).Address(ReferenceStyle:=xlR1C1) & _
                                                     "&"" ""&" & _
                                                     Mid(tempFormula, 2, InStr(tempFormula, "C") - 1) & _
                                                     Mid(tempAddress2, InStr(tempAddress2, "C") + 1)
                        Next j
        
                    Case columnName = SalesColumns.QTY
                        If .Cells(1).Value2 = 0 Then
                            tempAddress1 = Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME) _
                                     .columns(PurchaseColumns.QTY).Cells(1).Address(False, True, xlR1C1, , .Cells(1))
                            tempAddress1 = Mid(tempAddress1, InStr(tempAddress1, "C") + 1)
                            tempFormula = "=GCD("
                            
                            For j = 1 To .Rows.Count
    
                                tempFormula = tempFormula & "'" & PURCHASE_SHEET_NAME & "'!" & _
                                              Application.Worksheets(PURCHASE_SHEET_NAME) _
                                              .Cells(CLng(Mid(kitRange(j, findColNumber(desiredColumns, SalesColumns.Row)).Value2, 2)), _
                                               CLng(tempAddress1)).Address(False, True, xlR1C1, , .Cells(1)) & ","
                            Next j
    
                            .Cells(1).FormulaR1C1 = tempFormula & ")"
                        End If
                        
                        .Resize(kitRange.Rows.Count - 1).Offset(1).ClearContents
                        .Merge
                        
                    Case columnName = SalesColumns.Unit
                        .ClearContents
                        .Merge
                        .Cells(1).FormulaR1C1 = "=INDEX(ед_изм,2)"
                    
                    Case columnName = SalesColumns.Price
                        tempAddress1 = Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME) _
                                     .columns(PurchaseColumns.TOTAL_GPL_RECALCULATED).Cells(1).Address(False, True, xlR1C1, , .Cells(1))
                        tempAddress2 = Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME) _
                                     .columns(PurchaseColumns.TOTAL_PURCHASE_RECALCULATED).Cells(1).Address(False, True, xlR1C1, , .Cells(1))
                        
                        If Application.Worksheets(SALES_SHEET_NAME).shapes(GPL_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
                            total = TOTAL_GPL_CELL_NAME
                        ElseIf Application.Worksheets(SALES_SHEET_NAME).shapes(NET_PRICE_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
                            total = TOTAL_COST_CELL_NAME
                        End If
                        tempFormula = "=ROUND(IF(" & INCLUDE_DELIVERY_CELL_NAME & "=""да""" & ",(1+" & _
                                        DELIVERY_COST_CELL_NAME & "/" & total & "*INDEX(" & _
                                        CALC_CURRENCIES_ARRAY_NAME & ",MATCH(" & CALC_CURRENCY_CELL_NAME & _
                                        "," & CURRENCIES_ARRAY_NAME & ",0),2)*INDEX(" & CALC_VAT_ARRAY_NAME & _
                                        ",MATCH(" & INCLUDE_VAT_CELL_NAME & "," & VAT_ARRAY_NAME & ",0),1)),1)*SUM("
                                        
                        
                        For j = 1 To .Rows.Count
                            tempFormula = tempFormula & "IF(INDEX(" & CALC_SOURCE_ARRAY_NAME & ",1)=R[" & CStr(j - 1) & "]C" & _
                                           (findColNumber(desiredColumns, SalesColumns.CALC_SOURCE) + COLUMN_OFFSET) & _
                                           "," & "'" & PURCHASE_SHEET_NAME & "'!" & _
                                           kitRange(j, findColNumber(desiredColumns, SalesColumns.Row)).Value2 & _
                                           Mid(tempAddress1, InStr(tempAddress1, "C")) & "," & "'" & PURCHASE_SHEET_NAME & "'!" & _
                                           kitRange(j, findColNumber(desiredColumns, SalesColumns.Row)).Value2 & _
                                           Mid(tempAddress2, InStr(tempAddress2, "C")) & ")" & _
                                           "*IF(INDEX(" & PROFIT_TYPE_ARRAY_NAME & ",1)=R[" & CStr(j - 1) & "]C" & _
                                           (findColNumber(desiredColumns, SalesColumns.PROFIT_TYPE) + COLUMN_OFFSET) & _
                                           ",(1+R[" & CStr(j - 1) & "]C" & (findColNumber(desiredColumns, SalesColumns.PROFIT) + COLUMN_OFFSET) & _
                                           "),1/(1-R[" & CStr(j - 1) & "]C" & (findColNumber(desiredColumns, SalesColumns.PROFIT) + COLUMN_OFFSET) & _
                                           ")),"
                        Next j
    
                        .ClearContents
                        .Merge
                        .Cells(1).FormulaR1C1 = Mid(tempFormula, 1, Len(tempFormula) - 1) & ")/" & _
                                                kitRange(1, columnNumberQty).Address(False, True, xlR1C1, , .Cells(1)) & _
                                                "," & PRICE_ROUNDING_UP_TO_QTY & ")"
                    
                    Case columnName = SalesColumns.total
                        .Resize(kitRange.Rows.Count - 1).Offset(1).ClearContents
                        .Merge
                        
                    Case columnName = SalesColumns.VAT
                        .Resize(kitRange.Rows.Count - 1).Offset(1).ClearContents
                        .Merge
                End Select
            End With
        Next i
        
        tempFormula = kitRange(1, findColNumber(desiredColumns, SalesColumns.NAME_AND_DESCRIPTION)).FormulaR1C1
        With Application.Worksheets(SALES_SHEET_NAME).Range(kitRange(1, SalesColumns.INDEX_NUMBER + 1), kitRange(1, findColNumber(desiredColumns, SalesColumns.QTY) - 1))
                .ClearContents
                .Merge
                .FormulaR1C1 = tempFormula
        End With
        
        formatRangeAsType kitRange, "kit"
    End If
    
CleanExit:
    Set column = Nothing
    Set kitRange = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Public Function isArrayEmpty(anArray As Variant) As Boolean
' -------------------------------------------------------------------------------- '
' Проверяет пустоту массива запросом верней границы первого измерения и отлову
' возможной ошибки
'
' http://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
' -------------------------------------------------------------------------------- '
    Dim i As Integer
    
    On Error Resume Next
    i = UBound(anArray, 1) ' Just try it. If it fails, Err.Number will be nonzero.
    isArrayEmpty = (err.number <> 0)
    err.Clear
End Function

Private Function isExist(col As Collection, key As Variant) As Boolean
' -------------------------------------------------------------------------------- '
' Проверяет наличие элемента с ключём key в коллекции
' -------------------------------------------------------------------------------- '
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    isExist = (err.number = 0)
    err.Clear
End Function

Private Function isExistNamedRange(rangeName As String) As Boolean
' -------------------------------------------------------------------------------- '
' Проверяет наличие именованного диапазона с именем rangeName
' -------------------------------------------------------------------------------- '
    On Error Resume Next
    Application.Names.Item (rangeName)  ' Just try it. If it fails, Err.Number will be nonzero.
    isExistNamedRange = (err.number = 0)
    err.Clear
End Function

Public Function isExistShape(shapeName As String, sheetName As String) As Boolean
' -------------------------------------------------------------------------------- '
' Проверяет существование формы shapeName на листе sheetName
' -------------------------------------------------------------------------------- '
    Dim shape As shape
    On Error Resume Next
    Set shape = Application.Worksheets(sheetName).shapes(shapeName)
    isExistShape = (err.number = 0)
    
    err.Clear
    Set shape = Nothing
End Function

Public Function customGCD(arr() As Variant) As Variant ' не пригодилась :(
    Dim i As Variant
    Dim j As Variant
    
    For i = customMin(arr) To 2 Step -1
        For j = LBound(arr, 1) To UBound(arr, 1)
            If arr(j, 1) Mod i <> 0 Then: Exit For
        Next j
        
        If j > UBound(arr) Then: Exit For
    Next i
    
    customGGC = i
End Function

Public Function customMin(arr() As Variant) As Variant ' не пригодилась :(
    Dim i As Variant
    customMin = MAXLONG
    
    For Each i In arr
        If IsNumeric(i) Then
            If customMin > CLng(i) And Not i = Empty Then: customMin = CLng(i)
        Else
            Set customMin = Nothing
            Exit For
        End If
    Next i
End Function

Private Function getIndexGroupCount(indexColumn() As Variant) As Collection ' не пригодилась :(
    Dim i As Long, _
        j As Long
    Dim minRankCount As Long: minRankCount = INDEX_RANK_QTY
    Dim maxRankCount As Long: maxRankCount = 0
    Dim temp As Object
    Dim regexp As Object: Set regexp = CreateObject("vbscript.regexp")
    
    Set getIndexGroupCount = New Collection
    
    With regexp
        .Global = True
        .Pattern = "\d+"
        For i = LBound(indexColumn, 1) To UBound(indexColumn, 1)
            If .test(indexColumn(i, 1)) Then
                Set temp = .Execute(indexColumn(i, 1))
                If temp.Count < minRankCount Then: minRankCount = temp.Count
                If temp.Count > maxRankCount Then: maxRankCount = temp.Count
            End If
        Next i
    End With
    
    getIndexGroupCount.Add minRankCount, key:="MIN"
    getIndexGroupCount.Add maxRankCount, key:="MAX"
    
End Function

Private Function vStack2D(arr1() As Variant, arr2() As Variant) ' не пригодилась :(
' -------------------------------------------------------------------------------- '
' Аналог numpy.vstack для двух двумерных массивов VBA.
' Сцепляет два массива вертикально (построчно). Количество колонок у массивов
' должно совпадать. Если количество колонок не совпадает, а также если пусты оба
' полученных массива, то возвращает пустой объект Variant()
' -------------------------------------------------------------------------------- '
    Dim temp() As Variant
    Dim i As Long, _
        j As Long, _
        k As Long
    Dim cols1 As Long, cols2 As Long, _
        rows1 As Long, rows2 As Long
    
    Select Case True
        Case isArrayEmpty(arr1)
            temp = arr2
        Case isArrayEmpty(arr2)
            temp = arr1
        Case Else
            cols1 = UBound(arr1, 2) - LBound(arr1, 2) + 1
            cols2 = UBound(arr2, 2) - LBound(arr2, 2) + 1
            rows1 = UBound(arr1, 1) - LBound(arr1, 1) + 1
            rows2 = UBound(arr2, 1) - LBound(arr2, 1) + 1
    
            If cols1 = cols2 Then
                ReDim temp(rows1 + rows2, cols1)
                k = 1
                
                For i = LBound(arr1, 1) To UBound(arr1, 1)
                    For j = LBound(arr1, 2) To UBound(arr1, 2)
                        temp(k, j) = arr1(i, j)
                    Next j
                    
                    k = k + 1
                Next i
                
                For i = LBound(arr2, 1) To UBound(arr2, 1)
                    For j = LBound(arr2, 2) To UBound(arr2, 2)
                        temp(k, j) = arr2(i, j)
                    Next j
                    
                    k = k + 1
                Next i
            End If
    End Select
    
    vStack2D = temp
End Function
