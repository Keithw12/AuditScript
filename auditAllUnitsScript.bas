Attribute VB_Name = "auditAllUnitsScript"
'Global Variables
Public debugFlag As Boolean
Public totalNumOfMembersFound As Integer
Public totalNumOfMembersROBFound As Integer

'Description:  Marks an "X" in the current row for the column of the relevant file if it is present
'Params:  result - a number that is a file identifier, row - the current row we are on
Sub logResult(result As Integer, row As Integer)
    'Outputs the result of our file search
    '0 - No Match, 1 - 4433 Match, 2 - 4394, 3 - 2842, 4 - Deriv Classification, 5 - Security Briefing, 6 - 2875S, 7 - 2875N, 8 - ROB
    If result = 1 Then
        Cells(row, 2) = "X"
    ElseIf result = 2 Then
        Cells(row, 3) = "X"
    ElseIf result = 3 Then
        Cells(row, 4) = "X"
    ElseIf result = 4 Then
        Cells(row, 5) = "X"
    ElseIf result = 5 Then
        Cells(row, 6) = "X"
    ElseIf result = 6 Then
        Cells(row, 7) = "X"
    ElseIf result = 7 Then
        Cells(row, 8) = "X"
    ElseIf result = 8 Then
        totalNumOfMembersROBFound = totalNumOfMembersROBFound + 1
        Cells(row, 9) = "X"
    End If
End Sub
'Description:  Checks each file in "memberFolderReference" with performFileCheck() then uses logResult() to output the result on the spreadsheet
'Params:  memberFolderReference - a reference to the current member folder, row - the current spreadsheet row we are logging the result to
Sub iterateFiles(memberFolderReference As Folder, row As Integer)
    Dim fileReference As file
    Dim result As Integer
    
    For Each fileReference In memberFolderReference.Files
        result = performFileCheck(fileReference.name)
        If result = 0 Then GoTo NextIter
        Call logResult(result, row)
NextIter:
        Debug.Print fileReference.name
    Next
End Sub
Function openMemberFolder(folderPath As String, row As Integer) As Folder
    Dim fs As New FileSystemObject
    Dim f3 As Folder
    
    'f3 is an object of the member folder we are accessing
    Set f3 = fs.GetFolder(folderPath)
    Set openMemberFolder = f3
End Function
Function performFileCheck(fileName As String) As Integer
    'Compares file name to RegEx Patterns and returns a result of the check
    'Check Return Values: 0 - No Match, 1 - 4433 Match, 2 - 4394, 3 - 2842, 4 - Deriv Classification, 5 - Security Briefing, 6 - 2875S, 7 - 2875N, 8 - ROB
    Dim matchResult As Integer: matchResult = 0
    Dim regEx As New RegExp
    
    regEx.IgnoreCase = True
    
    'Check for 4433
    regEx.Pattern = "4433"
    If regEx.Test(fileName) Then
        performFileCheck = 1
        Exit Function
    End If
    regEx.Pattern = "4394"
    If regEx.Test(fileName) Then
        performFileCheck = 2
        Exit Function
    End If
    regEx.Pattern = "2842"
    If regEx.Test(fileName) Then
        performFileCheck = 3
        Exit Function
    End If
    regEx.Pattern = "Derivative"
    If regEx.Test(fileName) Then
        performFileCheck = 4
        Exit Function
    End If
    regEx.Pattern = "Security Briefing"
    If regEx.Test(fileName) Then
        performFileCheck = 5
        Exit Function
    End If
    regEx.Pattern = "2875S"
    If regEx.Test(fileName) Then
        performFileCheck = 6
        Exit Function
    Else
        regEx.Pattern = "2875"
        If regEx.Test(fileName) Then
            regEx.Pattern = "SIPR"
            If regEx.Test(fileName) Then
                performFileCheck = 6
                Exit Function
            End If
        End If
    End If
    regEx.Pattern = "2875N"
    If regEx.Test(fileName) Then
        performFileCheck = 7
        Exit Function
    End If
    regEx.Pattern = "Rules of Behavior"
    If regEx.Test(fileName) Then
        performFileCheck = 8
        Exit Function
    End If
End Function
Private Sub clearCells()
    Range("A:I").Clear
End Sub

Private Sub createHeaderRow()
    Cells(1, 1) = "Name"
    Cells(1, 2) = "4433"
    Cells(1, 3) = "4394"
    Cells(1, 4) = "2842"
    Cells(1, 5) = "Derivative Classification"
    Cells(1, 6) = "Security Briefing"
    Cells(1, 7) = "2875S"
    Cells(1, 8) = "2875N"
    Cells(1, 9) = "Rules of Behavior"
End Sub

Private Sub deleteSheets()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        Application.DisplayAlerts = False
        If wb.Worksheets.Count > 1 Then
            ws.Delete
        End If
        Application.DisplayAlerts = True
    Next
End Sub
Private Sub GenerateStatistics()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim ws As Worksheet
    
    'Add Worksheet to beginning
    Set ws = wb.Sheets.Add(Type:=xlWorksheet, Before:=wb.Worksheets(1))
    
    Worksheets(1).Activate
    ActiveSheet.name = "Statistics"
    Cells(1, 1) = "Total Members Found"
    Cells(2, 1) = totalNumOfMembersFound
    Cells(1, 2) = "Total Members with ROB"
    Cells(2, 2) = totalNumOfMembersROBFound
    'Output percentage
    Range("C2:C2").NumberFormat = "0.00\%"
    Range("C2:C2").Value = (totalNumOfMembersROBFound / totalNumOfMembersFound) * 100
End Sub



Sub MainMacro()
    'Init globals
    debugFlag = False
    totalNumOfMembersFound = 0
    totalNumOfMembersROBFound = 0

    'Reference Windows Scfript Host Object Model
    Dim currentRow As Integer: currentRow = 2
    Dim nameCol As Integer: nameCol = 1
    Dim firstChar As String
    Dim unitCount As Integer: unitCount = 1
    Dim fs As New FileSystemObject
    Dim recordsFolderReference As Folder
    Dim unitFolderReference As Folder
    Dim cssFolderReference As Folder
    Dim cssSubfolderReference As Folder
    Dim memberFolderReference As Folder
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet

    Dim userGivenPath As String: userGivenPath = ""
    Dim myFile As String
    Dim FldrPicker As FileDialog
    
    
    

    'Macro optimizations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    'Shows dialog to user to select the CSS folder
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
        With FldrPicker
            .Title = "Select the main records folder for all units"
            .AllowMultiSelect = False
            If .Show <> -1 Then GoTo NextCode
            userGivenPath = .SelectedItems(1) & "\"
        End With
        
NextCode:
    'In Case of Cancel
    If userGivenPath = "" Then GoTo ResetSettings
    
    Call deleteSheets
    
    Set recordsFolderReference = fs.GetFolder(userGivenPath)
    
    'Set the first sheet as active (for all cell operations to run on this sheet)
    For Each unitFolderReference In recordsFolderReference.SubFolders
        'Check if it's a unit folder or a folder with a '_' or '(' prefixed to it.  We ignore '_', '(' prefixed folders.
        firstChar = Left(unitFolderReference.name, 1)
        If firstChar = "_" Then
            GoTo SkipNonUnitFolder
        ElseIf firstChar = "(" Then
            GoTo SkipNonUnitFolder
        End If
        'f2.Path is Type String and is the full path of member subdirectory
        Debug.Print unitFolderReference.Path
        'Just member subdirectory name
        Debug.Print unitFolderReference.name
        'Check to see if we are on the second unit
        If unitCount >= 2 Then
            'Adds a worksheet at the end
            Set ws = wb.Sheets.Add(Type:=xlWorksheet, After:=wb.Worksheets(wb.Worksheets.Count))
            'Make this sheet active
            Worksheets(wb.Worksheets.Count).Activate
            ActiveSheet.name = unitFolderReference.name
            Call createHeaderRow
        Else
            Worksheets(1).Activate
            ActiveSheet.name = unitFolderReference.name
            Call clearCells
            'Create the header
            Call createHeaderRow
        End If
        Set cssFolderReference = fs.GetFolder(unitFolderReference.Path & "\CSS\")
        Debug.Print cssFolderReference.Path
        
        'Iterate through each subfolder (f2.Path is the path of the subfolder)
        For Each cssSubfolderReference In cssFolderReference.SubFolders
            'Check if it's a member folder or a folder with a '_' prefixed to it.  We ignore '_' prefixed folders.
            firstChar = Left(cssSubfolderReference.name, 1)
            If firstChar = "_" Then
                GoTo NextIteration
            End If
            'cssSubfolderReference.Path is Type String and is the full path of member subdirectory
            Debug.Print cssSubfolderReference.Path
            'Just member subdirectory name
            Debug.Print cssSubfolderReference.name
            'Write Member name to first column
            Cells(currentRow, nameCol) = cssSubfolderReference.name
            'Open member folder and iterate through each file
            Set memberFolderReference = openMemberFolder(cssSubfolderReference.Path, currentRow)
            Call iterateFiles(memberFolderReference, currentRow)
            totalNumOfMembersFound = totalNumOfMembersFound + 1
            currentRow = currentRow + 1
NextIteration:
        Next
        'Set row below header row
        currentRow = 2
        unitCount = unitCount + 1
        If debugFlag Then
            If unitCount = 3 Then
                GoTo GenerateStats
            End If
        End If
SkipNonUnitFolder:
    Next
        
GenerateStats:
    'Generate report based on collected data
    Call GenerateStatistics
    
    'Message Box when tasks are completed
    MsgBox "Task Complete!"
    
ResetSettings:
    'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub



