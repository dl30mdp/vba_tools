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

    ShGen.Visible = xlSheetVeryHidden
    ShData.Visible = xlSheetVeryHidden
    
End Sub

Private Sub uHideAll()

    ShGen.Visible = True
    ShData.Visible = True
    
End Sub

Public Sub enumWeekDay()

    Monday = 1
    Tuesday = 2 
    Wednesday = 3
    Thursday = 4
    Friday = 5
    Saturday = 6
    Sunday = 7
    
End Sub

Public Sub enumMonthNumb()

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

Public Sub sample()

    Dim lEnu_WeekDay As enu
    Dim weekDay As String

    If lEnu_weekDay = Monday Then weekDay = "Monday"
    
End Sub

Public Sub openFile(sFile As String)
    
    'Dim sFile As String
    'sFile = "C:\docs\data.xlsx"
    
    ' Use Dir to check if file exists
    If Dir(sFile) = "" Then
        ' if file does not exist display message
        MsgBox "Could not find the file " & sFile
        Exit Sub
    End If
    
    ' Code will only reach here if file exists
    Workbooks.Open sFile
    
End Sub

'Public Sub errorHandler()
'
'    On Error GoTo ErrorHandler
'    
'    Dim x As Long, y As Long
'    
'    x = 6
'    y = 6 / 0
'    x = 7
'    
'Done:
'
'    Exit Sub
'    
'ErrorHandler:
'
'    MsgBox "The following error occurred: " & Err.Description
'    
'End Sub
