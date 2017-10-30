Attribute VB_Name = "data_extraction"
Option Explicit

Public Function parseXML(xmlURL As String, xPath As String) As String
    Dim oXML As Object
    Set oXML = CreateObject("MSXML2.DOMDocument")

    oXML.async = False
    oXML.Load xmlURL

    If oXML.readystate = 4 Then
        parseXML = oXML.SelectSingleNode(xPath).Text
    Else
        parseXML = vbNullString
    End If
    
    Set oXML = Nothing
End Function

