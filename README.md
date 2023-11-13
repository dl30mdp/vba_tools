# vba
vba scripts

Private Sub DownApps()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False
End Sub

Private Sub UpApps()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    ActiveSheet.DisplayPageBreaks = False
End Sub

Private Sub HideAll()
    ShHidden.Visible = xlSheetVeryHidden
    ShData.Visible = xlSheetVeryHidden
    ShQuery.Visible = xlSheetVeryHidden
End Sub

Private Sub uHideAll()
    ShHidden.Visible = True
    ShData.Visible = True
    ShQuery.Visible = True
End Sub

Private Sub enumWeekDay()
    Monday = 1
    Tuesday = 2 
    Wednesday = 3
    Thursday = 4
    Friday = 5
    Saturday = 6
    Sunday = 7
End Sub

Private Sub enumMonthNumb()
    January = 01
    February = 02
    March = 03
    April = 04
    May = 05
    June = 06
    July = 07
    August = 08
    September = 09
    October = 10
    November = 11
    December = 12
End Sub
