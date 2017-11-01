VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист4"
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
    Dim isect As Range

    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    With Application.Worksheets(AGREEMENT_SHEET_NAME)
        ' Если вставляем или вводим данные в колонки с ценами, то предварительно
        ' пытаемся преобразовать их в число
        Set isect = Application.Intersect(Target, .Range(TENDER_CELL_NAME))
        If Not (isect Is Nothing) Then
            If isect.Value2 = "да" Then
                .Range(TENDER_ARRAY_NAME).Resize(.Range(TENDER_ARRAY_NAME).Rows.Count - 3).offset(2).EntireRow.Hidden = False
            Else
                .Range(TENDER_ARRAY_NAME).Resize(.Range(TENDER_ARRAY_NAME).Rows.Count - 3).offset(2).EntireRow.Hidden = True
            End If
        End If
        
        Set isect = Application.Intersect(Target, .Range(ASSURANCE_CELL_NAME))
        If Not (isect Is Nothing) Then
            If isect.Value2 = "да" Then
                .Range(ASSURANCE_ARRAY_NAME).Resize(.Range(ASSURANCE_ARRAY_NAME).Rows.Count - 3).offset(2).EntireRow.Hidden = False
            Else
                .Range(ASSURANCE_ARRAY_NAME).Resize(.Range(ASSURANCE_ARRAY_NAME).Rows.Count - 3).offset(2).EntireRow.Hidden = True
            End If
        End If
    End With
    
CleanExit:
    Set isect = Nothing
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description
    Resume CleanExit
End Sub

