Attribute VB_Name = "data_extraction"
Option Explicit

Public Function parseXML(xmlURL As String, xPath As String) As String
    Dim oXML As Object
    Set oXML = CreateObject("MSXML2.DOMDocument")

    oXML.async = False
    oXML.Load xmlURL

    If oXML.readystate = 4 Then
        On Error Resume Next
        parseXML = oXML.SelectSingleNode(xPath).Text
        On Error GoTo ErrorHandler
    Else
        parseXML = vbNullString
    End If

CleanExit:
    Set oXML = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    parseXML = vbNullString
    Resume CleanExit
End Function
