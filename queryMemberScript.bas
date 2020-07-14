Attribute VB_Name = "queryMemberScript"
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

'Description:  Deletes all but 1 Sheet in the Workbook.  The workbook has to have at least 1 sheet.
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
Function formatMemberName(memberName As String) As String()
    Dim name() As String
    'if memberName = "Doe.John" then name(0) = "Doe", name(1) = "John"
    name = Split(memberName, ".")
    formatMemberName = name
End Function

Function isValidMemberFolderName(memberNameArr() As String) As Boolean
    'check if number of elements in array is 2 or more, if not we don't have a first and last name and it's invalid
    If UBound(memberNameArr, 1) <> 1 Then
        isValidMemberFolderName = False
    Else
        isValidMemberFolderName = True
    End If
End Function

Function compareNames(givenMemberNameArr() As String, memberFolderNameArr() As String) As Boolean
    If givenMemberNameArr(0) = memberFolderNameArr(1) Then
            If givenMemberNameArr(1) = memberFolderNameArr(0) Then
                compareNames = True
            Else
                compareNames = False
            End If
    Else
        compareNames = False
    End If
End Function


'Description:  Read each subfolder(member folder) without a '_' prefix in the Unit folder
Private Function findMemberFolder(unitFolderReference As Folder, memberName As String) As Boolean
    Dim cssFolderReference As Folder
    Dim cssSubfolderReference As Folder
    Dim memberFolderReference As Folder
    Dim currentRow As Integer: currentRow = 2
    Dim firstChar As String
    Dim nameCol As Integer: nameCol = 1
    Dim fs As New FileSystemObject
    
    'memberNameArr(0) = "Doe", memberNameArr(1) = "John"
    Dim memberNameArr() As String
    Dim givenMemberNameArr() As String
    
    '(0) - first name, (1) - last name
    givenMemberNameArr = Split(memberName, " ")
    
    Set cssFolderReference = fs.GetFolder(unitFolderReference.Path & "\CSS\")
        
        Debug.Print cssFolderReference.Path
        
        'Iterate through each subfolder (f2.Path is the path of the subfolder)
        For Each cssSubfolderReference In cssFolderReference.SubFolders
            'Check if it's a member folder or a folder with a '_' prefixed to it.  We ignore '_' prefixed folders.
            firstChar = Left(cssSubfolderReference.name, 1)
            If firstChar = "_" Then
                GoTo NextIteration
            End If
            memberNameArr = formatMemberName(cssSubfolderReference.name)
            If Not (isValidMemberFolderName(memberNameArr)) Then
                Debug.Print "Invalid folder name"
                GoTo NextIteration
            End If
            
            If compareNames(givenMemberNameArr, memberNameArr) Then
                Debug.Print "Found Member:"
                Debug.Print memberNameArr(1) & " " & memberNameArr(0)
                findMemberFolder = True
                Call Shell("explorer.exe" & " " & cssSubfolderReference.Path, vbNormalFocus)
                Exit Function
            End If
NextIteration:
        Next
        findMemberFolder = False
End Function
'Description:  Check if it's a unit folder or a folder with a '_' or '(' prefixed to it.  We ignore '_', '(' prefixed folders.
Function isUnitFolder(unitFolder As String) As Boolean
    Dim firstChar As String
        firstChar = Left(unitFolder, 1)
        If firstChar = "_" Then
            isUnitFolder = False
            Exit Function
        ElseIf firstChar = "(" Then
            isUnitFolder = False
            Exit Function
        End If
        isUnitFolder = True
End Function


'Description:  Main function that is called from queryMemberForm
'Params: formInputString - "First Last" name format
Sub MainMacro(formInputString As String)
    'Init globals
    'debugFlag limits the script to run on 3 units to save time on script testing
    debugFlag = False
    totalNumOfMembersFound = 0
    totalNumOfMembersROBFound = 0

    'Reference Windows Scfript Host Object Model
    
    Dim unitCount As Integer: unitCount = 1
    Dim userGivenPath As String: userGivenPath = ""
    
    Dim fs As New FileSystemObject
    Dim recordsFolderReference As Folder
    Dim unitFolderReference As Folder
    Dim FldrPicker As FileDialog
    
    'Assuming that we won't have more than 30 unit folders for now
    Dim unitFolderNamesArray(30)
    
    
    

    'Macro optimizations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    'Shows dialog to user to select the reocrds folder
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
    
    Set recordsFolderReference = fs.GetFolder(userGivenPath)
    
    'Set the first sheet as active (for all cell operations to run on this sheet)
    For Each unitFolderReference In recordsFolderReference.SubFolders
        If Not (isUnitFolder(unitFolderReference.name)) Then
            GoTo SkipNonUnitFolder
        End If

        If findMemberFolder(unitFolderReference, formInputString) Then
            GoTo ResetSettings
        End If
        
        unitCount = unitCount + 1
        If debugFlag Then
            If unitCount = 3 Then
                GoTo ResetSettings
            End If
        End If
SkipNonUnitFolder:
    Next
    
NotFound:
    Debug.Print "Couldn't locate member."
        
ResetSettings:
    'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub




