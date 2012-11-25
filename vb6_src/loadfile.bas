Attribute VB_Name = "loadfile"
Public Function loadSession(aFileName As String)

On Error GoTo errorOnLoad

    Dim fileNum As Long
    Dim fileIsOpen As Boolean
    Dim aForm As New frmSessionEdit
    
    Dim readLine As String
    
    fileNum = FreeFile
    fileIsOpen = False


    Open aFileName For Input As fileNum
    fileIsOpen = True
    
    'get the tester initials
    Dim aString As String
    aString = StrReverse(aFileName)
    aString = Left$(aString, InStr(1, aString, "\") - 1)
    aString = StrReverse(aString)
    Dim firstDash As Long
    Dim secondDash As Long
    firstDash = InStr(1, aString, "-")
    secondDash = InStr(firstDash + 1, aString, "-")
    aForm.testerInitials = Mid$(aString, firstDash + 1, secondDash - firstDash - 1)
    aForm.sessionCount = Asc(Mid$(aString, InStr(1, aString, ".") - 1, 1)) - Asc("A")
    
    If Not EOF(fileNum) Then
        Line Input #fileNum, readLine
    End If
    Do Until EOF(fileNum)

        Select Case readLine
        Case "CHARTER"
            'discard -----...
            Line Input #fileNum, readLine
            readLine = loadCharter(fileNum, aForm)
            
        Case "START"
            'discard -----...
            Line Input #fileNum, readLine
            readLine = loadStart(fileNum, aForm)
            
        Case "TESTER"
            'discard -----...
            Line Input #fileNum, readLine
            readLine = loadTester(fileNum, aForm)
            
        Case "TASK BREAKDOWN"
            'discard -----...
            Line Input #fileNum, readLine
            readLine = loadTaskBreakdown(fileNum, aForm)
            
        Case "DATA FILES"
            'discard -----...
            Line Input #fileNum, readLine
            readLine = loadDataFiles(fileNum, aForm)
            
        Case "TEST NOTES"
            'discard -----...
            Line Input #fileNum, readLine
            readLine = loadTestNotes(fileNum, aForm)
             
        Case "BUGS"
            'discard -----...
            Line Input #fileNum, readLine
            readLine = loadBugs(fileNum, aForm)
            
        Case "ISSUES"
            'discard -----...
            Line Input #fileNum, readLine
            readLine = loadIssues(fileNum, aForm)
            
        Case Else
            'discard the line
            Line Input #fileNum, readLine
        End Select
        
    Loop
    
    
    Close fileNum
    fileIsOpen = False
    
    aForm.Show
    
Exit Function

errorOnLoad:

    If fileIsOpen = True Then
        Close fileNum
    End If
    Set aForm = Nothing
    MsgBox "ERROR: " & Err.Number & " " & Err.Description
End Function


Public Function loadCharter(aFileNum As Long, aSession As frmSessionEdit) As String

    'charter is a block of text followed by
    'blank lines
    'and then #AREAS
    
    'charter processing ends with
    'START
    'which is passed back

    Dim readLine As String
    Dim areaDetails As String
    Dim processing As Integer   '1 for charter, 2 for areas
    
    aSession.charterText = ""
    areaDetails = ""
    processing = 1
    
    Line Input #aFileNum, readLine
    
    Do Until EOF(aFileNum)
    
        Select Case readLine
        Case ""
            If processing = 1 Then
                aSession.charterText = aSession.charterText & vbCrLf
            End If
        Case "#AREAS"
            processing = 2
        Case "START"
            loadCharter = readLine
            Exit Function
        Case Else
            If processing = 1 Then
                aSession.charterText = aSession.charterText & readLine & vbCrLf
            Else
                aSession.loadedSessionAreas.Add readLine
            End If
        End Select
        
        Line Input #aFileNum, readLine
        
    Loop
    
    loadCharter = readLine
End Function

Public Function loadStart(aFileNum As Long, aSession As frmSessionEdit) As String

    'start is a TIMESTAMP
    'blank lines
    
    'charter processing ends with
    'TESTER
    'which is passed back

    Dim readLine As String
    Dim timestamp As String
       
    Line Input #aFileNum, readLine
    
    Do Until EOF(aFileNum)
    
        Select Case readLine
        Case ""
        Case "TESTER"
            loadStart = readLine
            Exit Function
        Case Else
            aSession.startDate = readLine
        End Select
        
        Line Input #aFileNum, readLine
        
    Loop
    
    loadStart = readLine
End Function

Public Function loadTester(aFileNum As Long, aSession As frmSessionEdit) As String

    'tester is a NAME
    'blank lines
    
    'tester processing ends with
    'TASK BREAKDOWN
    'which is passed back

    Dim readLine As String
    
    Line Input #aFileNum, readLine
    
    Do Until EOF(aFileNum)
    
        Select Case readLine
        Case ""
        Case "TASK BREAKDOWN"
            loadTester = readLine
            Exit Function
        Case Else
            aSession.testername = readLine
        End Select
        
        Line Input #aFileNum, readLine
        
    Loop
    
    loadTester = readLine
End Function

Public Function loadTaskBreakdown(aFileNum As Long, aSession As frmSessionEdit) As String

    'taskbreak down is
    '#DURATION
    'text
    'blank lines
    '#TEST DESIGN AND EXECUTION
    'text
    'blank lines
    '#BUG INVESTIGATION AND REPORTING
    'text
    'blank lines
    '#SESSION SETUP
    'text
    'blank lines
    '#CHARTER VS. OPPORTUNITY
    'text
    'blank lines
    
    'processing ends with
    'DATA FILES
    'which is passed back

    Dim readLine As String
    Dim processing As Integer ' 1 duration, 2 test design etc..
    
    Line Input #aFileNum, readLine
    
    processing = 0
    
    Do Until EOF(aFileNum)
    
        Select Case readLine
        Case ""
        Case "DATA FILES"
            loadTaskBreakdown = readLine
            Exit Function
        Case "#DURATION"
            processing = 1
        Case "#TEST DESIGN AND EXECUTION"
            processing = 2
        Case "#BUG INVESTIGATION AND REPORTING"
            processing = 3
        Case "#SESSION SETUP"
            processing = 4
        Case "#CHARTER VS. OPPORTUNITY"
            processing = 5
        Case Else
            Select Case processing
            Case 1 'duration
                aSession.duration = readLine
            Case 2
                aSession.testDesign = readLine
            Case 3
                aSession.bugInvest = readLine
            Case 4
                aSession.sessionSetup = readLine
            Case 5
                aSession.charterVop = readLine
            End Select
            processing = 0
        End Select
        
        Line Input #aFileNum, readLine
    Loop
    
    loadTaskBreakdown = readLine
End Function

Public Function loadDataFiles(aFileNum As Long, aSession As frmSessionEdit) As String

    'is a set of datafiles or #N/A
    'blank lines
    
    'processing ends with
    'TEST NOTES
    'which is passed back

    Dim readLine As String
    
    Line Input #aFileNum, readLine
    
    Do Until EOF(aFileNum)
    
        Select Case readLine
        Case ""
        Case "TEST NOTES"
            loadDataFiles = readLine
            Exit Function
        Case "#N/A"
        Case Else
                aSession.loadedDataFiles.Add readLine
        End Select
        
        Line Input #aFileNum, readLine
        
    Loop
    
    loadDataFiles = readLine
End Function


Public Function loadTestNotes(aFileNum As Long, aSession As frmSessionEdit) As String

    'a block of text or #N/A followed by
    'blank lines
    
    'processing ends with
    'BUGS
    'which is passed back

    Dim readLine As String
    
    aSession.testNotes = ""
        
    Line Input #aFileNum, readLine
    
    Do Until EOF(aFileNum)
    
        Select Case readLine
        Case ""
                aSession.testNotes = aSession.testNotes & vbCrLf
                
        Case "#N/A"

        Case "BUGS"
            loadTestNotes = readLine
            Exit Function
            
        Case Else
                aSession.testNotes = aSession.testNotes & readLine & vbCrLf
        End Select
        
        Line Input #aFileNum, readLine
        
    Loop
    
    loadTestNotes = readLine
End Function

Public Function loadBugs(aFileNum As Long, aSession As frmSessionEdit) As String

    'a block of bugs or #N/A followed by
    ' a bug is
    ' #BUG nnnnnnn
    ' text
    ' blank lines
    
    'processing ends with
    'ISSUES
    'which is passed back

    Dim readLine As String
    Dim processingABug As Boolean
    Dim bugNum As String
    Dim bugText As String
        
    Line Input #aFileNum, readLine
    processingABug = False
    bugNum = ""
    bugText = ""
    
    Do Until EOF(aFileNum)
    
        Select Case readLine
        Case ""
            If processingABug Then
                bugText = bugText & vbCrLf
            End If
                
        Case "#N/A"

        Case "ISSUES"
            If processingABug Then
                bugNum = bugNum & "|" & bugText
                aSession.loadedBugs.Add bugNum
                bugNum = ""
                bugText = ""
            End If
            loadBugs = readLine
            Exit Function
            
        Case Else
            If Left$(readLine, 4) = "#BUG" Then
                If processingABug Then
                    bugNum = bugNum & "|" & bugText
                    aSession.loadedBugs.Add bugNum
                    bugNum = ""
                    bugText = ""
                End If
                bugNum = Trim$(Mid$(readLine, 5))
                processingABug = True
            Else
                If processingABug Then
                    bugText = bugText & readLine & vbCrLf
                End If
            End If
        End Select
        
        Line Input #aFileNum, readLine
        
    Loop
    
    loadBugs = readLine
End Function

Public Function loadIssues(aFileNum As Long, aSession As frmSessionEdit) As String

    'a block of issues or #N/A followed by
    ' a issue is
    ' #ISSUE nnnnnnn
    ' text
    ' blank lines
    
    'processing ends with
    'EOF
    'which is passed back

    Dim readLine As String
    Dim processingAnIssue As Boolean
    Dim issueNum As String
    Dim issueText As String
        
    Line Input #aFileNum, readLine
    processingAnIssue = False
    issueNum = ""
    issueText = ""
    
    Do Until EOF(aFileNum)
    
        Select Case readLine
        Case ""
            If processingAnIssue Then
                issueText = issueText & vbCrLf
            End If
                
        Case "#N/A"
            
        Case Else
            If Left$(readLine, 6) = "#ISSUE" Then
                If processingAnIssue Then
                    issueNum = issueNum & "|" & issueText
                    aSession.loadedIssues.Add issueNum
                    issueNum = ""
                    issueText = ""
                End If
                issueNum = Trim$(Mid$(readLine, 7))
                processingAnIssue = True
            Else
                If processingAnIssue Then
                    issueText = issueText & readLine & vbCrLf
                End If
            End If
        End Select
        
        Line Input #aFileNum, readLine
        
    Loop
    
    If processingAnIssue Then
        issueNum = issueNum & "|" & issueText
        aSession.loadedIssues.Add issueNum
        issueNum = ""
        issueText = ""
    End If

    
    loadIssues = readLine
End Function

