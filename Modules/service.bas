Attribute VB_Name = "service"
Option Explicit

Public Sub changeUpdatingState(Optional turnOn As Boolean = True, Optional calcState As XlCalculation = xlCalculationAutomatic)
    With Application
        If turnOn Then     ' включение обновлений экрана и отслеживания событий
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = calcState
        Else                    ' отключение обновлений экрана и отслеживания событий
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        End If
    End With
End Sub

Public Sub initAutoCorrectState()
    With Application.AutoCorrect
        .AutoExpandListRange = False
        .AutoFillFormulasInLists = True
    End With
End Sub
