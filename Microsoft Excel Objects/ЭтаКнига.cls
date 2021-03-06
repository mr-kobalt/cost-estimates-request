VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
' -------------------------------------------------------------------------------- '
' ���������� �������� ����� � ����� �� �� �� � ������ USD/EUR �� ��� ���� � �����
' �� �� �� ����� �������� �����
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
        With Application.Worksheets(PURCHASE_SHEET_NAME)
            On Error Resume Next
            .Range(CURRENT_RATE_DATE_CELL_NAME).Value2 = "���� �� �� �� " & oXML.SelectSingleNode(CURRENT_RATE_DATE_XPATH).Text
            .Range(USD_RATE_CELL_NAME).Value2 = Replace(oXML.SelectSingleNode(USD_RATE_XPATH).Text, ",", ".")
            .Range(EUR_RATE_CELL_NAME).Value2 = Replace(oXML.SelectSingleNode(EUR_RATE_XPATH).Text, ",", ".")
            On Error GoTo ErrorHandler
        End With
    End If

CleanExit:
    Application.EnableEvents = True
    Set oXML = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description
    Resume CleanExit
End Sub
