VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
' -------------------------------------------------------------------------------- '
' Обработка внесение изменений в ячейки листа пользователем
' 1. Вырезание ячеек запрещаем (побочный, но полезный эффект)
' 2. Любые значения вставляются как данные и формулы, форматирование не копируется
' 3. Любые значение вставляемые в столбцы с ценами (прайс и вход) пытаемся
'    пытаемся преобразовать в число с плавающей точкой (см. convertCellsValueToDbl)
' 4. Любое изменение ячеек с датой по ЦБ РФ и курсов USD/EUR на эту дату ведёт к
'    обновлению значения этих ячеек с сайта ЦБ РФ
' -------------------------------------------------------------------------------- '
    Dim isect As Range, priceRange As Range, cell As Range
    Dim undoControl As Object
    Dim i As Long

    On Error GoTo ErrorHandler

    'Application.EnableEvents = False
    changeUpdatingState False

    ' Перехватываем вставку и вставляем в ячейки только значения и формулы, но не форматирование и пр.
    ' Адаптированное решение подсмотренное здесь:
    ' http://www.siddharthrout.com/2011/08/15/vba-excelallow-paste-special-only/

    ' Сохраняем контрол отвечающий за отмену последних действий. 128 - его индекс по умолчанию
    Set undoControl = Application.CommandBars("Standard").FindControl(ID:=128)

    ' Если невозможно найти Control по индексу 128 или нет действий для отмены, то завершаем процедуру
    If undoControl.Enabled = True Then
        If undoControl.ListCount <> 0 Then
            ' Ищем тип последнего действия по его тексовому описанию. Должно работать для русской (точно)
            ' и английской (метод не тестировался, но судя по ссылке выше должен работать) локализации.
            ' TODO: найти более надёжный метод определения последнего действия
            If Left(undoControl.list(1), 5) = "Paste" Or Left(undoControl.list(1), 7) = "вставку" Then
                ' Отменяем последнее действие без очистки буфера обмена
                Application.Undo
                ' Выделяем область, в которую необходимо произвести вставку
                Target.Select
                ' Вставляем данные без форматирования
                On Error GoTo CutHandler
                Select Case Application.CutCopyMode
                    Case 0
                        ' Если значения скопированы не ячейки Excel, то пытаемся вставить как HTML без форматирования,
                        ' а в случае неудачи - просто вставляем как есть
                        On Error Resume Next
                        ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
                        If err.number <> 0 Then: Application.ActiveSheet.Paste
                        err.Clear
                    Case Else
                        ' Если Excel в режиме копирования или вырезания ячеек, то вставляем только значения. Если ячейки
                        ' были были вырезаны, то PasteSpecial критически завершит работу и управление будет передано в
                        ' обработчик CutHandler. Т.е. полный запрет на вставку вырезанных ячеек с любых листов и книг.
                        Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone
                End Select
            ElseIf Left(undoControl.list(1), 4) = "Fill" Or Left(undoControl.list(1), 10) = "автозаполнение" Then
                Application.Undo
                MsgBox "Заполнение ячеек на этом листе запрещено. Используйте копирование, вставку и удаление"
                GoTo CleanExit
            End If
        End If
    End If

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME)
        On Error GoTo ErrorHandler2
        With .Range(PURCHASE_TABLE_NAME)
            Set priceRange = Union(.columns(PurchaseColumns.PRICE_GPL), _
                                   .columns(PurchaseColumns.PRICE_PURCHASE))
        End With
        On Error GoTo ErrorHandler

        ' Если вставляем или вводим данные в колонки с ценами, то предварительно
        ' пытаемся преобразовать их в число
        Set isect = Application.Intersect(Target, priceRange)
        If Not (isect Is Nothing) Then: convertCellsValueToDbl isect


        ' Если вставляем или вводим данные в ячейку с Валютой расчёта (SALES_CURRENCY_CELL_NAME),
        ' то копируем новое значение в выпадающий список на листе Расчёт продажи (SALES_SHEET_NAME)
'        Set isect = Application.Intersect(Target, .Range(SALES_CURRENCY_CELL_NAME))
'        If Not (isect Is Nothing) Then
'            ThisWorkbook.Sheets(SALES_SHEET_NAME).shapes(CURRENCY_SHAPE_NAME).OLEFormat.Object.Value = _
'                        Application.Match(.Range(SALES_CURRENCY_CELL_NAME).Value2, _
'                        ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_ARRAY_NAME).Value2, 0)
'        End If

        ' Если вставляем или вводим данные в ячейку-индикатор включения НДС в расчёт (INCLUDE_VAT_CELL_NAME),
        ' то копируем новое значение в выпадающий список на листе Расчёт продажи (SALES_SHEET_NAME) и во все
        ' ячейки колонки "НДС продажи" на листе расчёта, а также изменяем валидацию ячеек этой  колонки
        Set isect = Application.Intersect(Target, .Range(INCLUDE_VAT_CELL_NAME))
        If Not (isect Is Nothing) Then
'            ThisWorkbook.Sheets(SALES_SHEET_NAME).Shapes(VAT_SHAPE_NAME).OLEFormat.Object.Value = _
'                        Application.Match(.Range(INCLUDE_VAT_CELL_NAME).Value2, _
'                        ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Value2, 0)
            createValidation Target:=.Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.VAT_SALES), typename:="VAT"
            With .Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.VAT_SALES)
                For i = 1 To .Cells.Count
                    If .Cells(i).Value2 <> ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Cells(3).Value2 And _
                                            .Cells(i).Value2 <> vbNullString Then
                        .Cells(i) = isect.Value2
                    End If
                Next i
            End With

        End If

        ' Если вставляем или вводим данные в ячейку-индикатор включения стоимости доставки расчёт (INCLUDE_DELIVERY_CELL_NAME),
        ' то правим чекбокс на листе Расчёт продажи (SALES_SHEET_NAME)
'        Set isect = Application.Intersect(Target, .Range(INCLUDE_DELIVERY_CELL_NAME))
'        If Not (isect Is Nothing) Then
'            If isect.Value2 = YES Then
'                ThisWorkbook.Sheets(SALES_SHEET_NAME).shapes(DELIVERY_SHAPE_NAME).OLEFormat.Object.Value = xlOn
'            Else
'                ThisWorkbook.Sheets(SALES_SHEET_NAME).shapes(DELIVERY_SHAPE_NAME).OLEFormat.Object.Value = xlOff
'            End If
'        End If

        ' Если меняется ячейка с датой курса, то подгружаем значение с сайта ЦБ РФ
        Set isect = Application.Intersect(Target, .Range(CURRENT_RATE_DATE_CELL_NAME))
        If Not (isect Is Nothing) Then: isect.Value2 = "Курс ЦБ РФ на " & parseXML(CBR_XML_URL, CURRENT_RATE_DATE_XPATH)

        ' Если меняется ячейка с текущим курсом USD, то подгружаем значение с сайта ЦБ РФ
        Set isect = Application.Intersect(Target, .Range(USD_RATE_CELL_NAME))
        If Not (isect Is Nothing) Then: isect.Value2 = Replace(parseXML(CBR_XML_URL, USD_RATE_XPATH), ",", ".")

        ' Если меняется ячейка с текущим курсом EUR, то подгружаем значение с сайта ЦБ РФ
        Set isect = Application.Intersect(Target, .Range(EUR_RATE_CELL_NAME))
        If Not (isect Is Nothing) Then: isect.Value2 = Replace(parseXML(CBR_XML_URL, EUR_RATE_XPATH), ",", ".")

        ' Если меняются ячейки с индексом, то корректируем их содержимое
        Set isect = Application.Intersect(Target, .Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.INDEX_NUMBER))
        If Not (isect Is Nothing) Then
            correctIndexRange indexRange:=isect
            updateIndexDesc
        End If

        'если меняются ячейки в таблице расчёта, то все формулы сбрасываются к значениям по умолчанию
        Set isect = Application.Intersect(Target, .Range(PURCHASE_TABLE_NAME))
'        If Not (isect Is Nothing) Then: resetFormulasInPurchaseTable
        If Not (isect Is Nothing) Then
            For i = isect.Rows(1).Cells(1).column To isect.Rows(1).Cells(isect.columns.Count).column
                resetFormulasInPurchaseTable i
            Next i
        End If
    End With
CleanExit:
    Set Target = Nothing
    Set isect = Nothing
    Set priceRange = Nothing
    'Application.EnableEvents = True
    changeUpdatingState True
    Exit Sub

CutHandler:
    MsgBox "Вырезание ячеек на этом листе запрещено. Используйте копирование, вставку и удаление" & _
           vbCrLf & vbCrLf & "Error " & err.number & ": " & err.Description
    Resume CleanExit

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit

ErrorHandler2:
    MsgBox "Error " & err.number & ": " & err.Description & vbCrLf & vbCrLf & _
           "Не найден лист с расчётом, таблица расчёта или колонка в таблице расчёта. Операция не завершена."
    Resume CleanExit

End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim priceRange As Range
    Dim isect As Range
    Dim cell As Range
    Dim tempPN As Variant
    Dim tempName As Variant
    Dim tempComment As String

    Application.EnableEvents = False

    On Error GoTo ErrorHandler2
    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(PURCHASE_TABLE_NAME)
        Set priceRange = Union(.columns(PurchaseColumns.PRICE_GPL), _
                               .columns(PurchaseColumns.PRICE_PURCHASE))
        On Error GoTo ErrorHandler

        Set isect = Application.Intersect(Target, priceRange)
        If Not (isect Is Nothing) Then
            priceRange.ClearComments
            For Each cell In isect
                tempPN = .Cells(cell.Row - .Cells(1).Row + 1, PurchaseColumns.PN).Value2
                tempName = .Cells(cell.Row - .Cells(1).Row + 1, PurchaseColumns.NAME_AND_DESCRIPTION).Value2
                tempComment = vbNullString
                If Not IsEmpty(tempPN) Then: tempComment = CStr(tempPN) & vbCrLf
                If Not IsEmpty(tempName) Then: tempComment = tempComment & CStr(tempName)
                If tempComment <> vbNullString Then: cell.AddComment tempComment
            Next cell
        End If
    End With

CleanExit:
    Set Target = Nothing
    Set isect = Nothing
    Set cell = Nothing
    Set priceRange = Nothing
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit

ErrorHandler2:
    MsgBox "Error " & err.number & ": " & err.Description & vbCrLf & vbCrLf & _
           "Не найден лист с расчётом, таблица расчёта или колонка в таблице расчёта. Операция не завершена."
    Resume CleanExit
End Sub
