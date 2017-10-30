VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    sheetsListUpdate

CleanExit:
    Application.EnableEvents = True
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub Workbook_Open()
' -------------------------------------------------------------------------------- '
' Обновление значения ячеек с датой по ЦБ РФ и курсов USD/EUR на эту дату с сайта
' ЦБ РФ во время открытия книги
' -------------------------------------------------------------------------------- '
    Dim oXML As Object
    Set oXML = CreateObject("MSXML2.DOMDocument")

    On Error GoTo ErrorHandler

    initAutoCorrectState
    Application.EnableEvents = False

    Application.UseSystemSeparators = False
    Application.DecimalSeparator = ","
    Application.ThousandsSeparator = " "

    resetFormulasInPurchaseTable
    initializeShapes

    oXML.async = False
    oXML.Load CBR_XML_URL

    If oXML.readystate = 4 Then
        With ThisWorkbook.Sheets(SPEC_SHEET_NAME)
            On Error Resume Next
            .Range(CURRENT_RATE_DATE_CELL_NAME).Value2 = "Курс ЦБ РФ на " & oXML.SelectSingleNode(CURRENT_RATE_DATE_XPATH).Text
            .Range(USD_RATE_CELL_NAME).Value2 = Replace(oXML.SelectSingleNode(USD_RATE_XPATH).Text, ",", ".")
            .Range(EUR_RATE_CELL_NAME).Value2 = Replace(oXML.SelectSingleNode(EUR_RATE_XPATH).Text, ",", ".")
            On Error GoTo ErrorHandler
        End With
    End If

    sheetsListUpdate

CleanExit:
    Application.EnableEvents = True
    Set oXML = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub
