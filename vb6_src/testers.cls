VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "testers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public testers As Collection


Private Sub Class_Initialize()

    Set testers = New Collection
    
End Sub

Public Function loadFromFile(aFilename As String) As Boolean

    Dim fileID As Long
    Dim fileline As String
    Dim fileChunks() As String
    Dim aTesterName As String
    Dim aTesterInitials As String
    Dim aTester As tester
    Dim loopIter As Long
    Dim isFileOpen As Boolean
    
    isFileOpen = False
    
    On Error GoTo errorOnLoad
    
    If Dir$(aFilename) <> "" Then
        'file exists
        fileID = FreeFile
        Open aFilename For Input As fileID
        isFileOpen = True
        
            Do Until EOF(fileID)
                Line Input #fileID, fileline
                fileChunks = Split(fileline, vbTab)
                aTesterName = ""
                aTesterInitials = ""
                For loopIter = 0 To UBound(fileChunks)
                    If Len(Trim(fileChunks(loopIter))) <> 0 Then
                        If aTesterName = "" Then
                            aTesterName = Trim(fileChunks(loopIter))
                        Else
                            aTesterInitials = Trim(fileChunks(loopIter))
                        End If
                    End If
                Next
                
                If aTesterName <> "" Then
                    'add it to the collection
                        Me.createTester aTesterName, aTesterInitials
                End If
                
            Loop
        
        Close fileID
        isFileOpen = False
    End If

    loadFromFile = True
    
    Exit Function
errorOnLoad:
    MsgBox "Error on load testers from " & aFilename & vbCrLf & Err.Number & vbCrLf & Err.Description
    If isFileOpen Then
        Close fileID
    End If
    loadFromFile = False

End Function

Public Function saveToFile(aFilename As String) As Boolean

    Dim fileID As Long
    Dim fileline As String
    Dim fileChunks() As String
    Dim aTesterName As String
    Dim aTesterInitials As String
    Dim aTester As tester
    Dim loopIter As Long
    Dim isFileOpen As Boolean
    
    isFileOpen = False
    
    On Error GoTo errorOnLoad
    
        fileID = FreeFile
        Open aFilename For Output As fileID
        isFileOpen = True
        
        For Each aTester In testers
            Print #fileID, aTester.name & vbTab & aTester.initials
        Next
        
        Close fileID
        isFileOpen = False
    

    saveToFile = True
    
    Exit Function
errorOnLoad:
    MsgBox "Error on save testers to " & aFilename & vbCrLf & Err.Number & vbCrLf & Err.Description
    If isFileOpen Then
        Close fileID
    End If
    saveToFile = False

End Function
Public Sub createTester(aTesterName As String, Optional aTesterInitials As String = "")
    Dim aTester As tester
    
    On Error GoTo noCreateTester
    
    Set aTester = New tester
    
    aTester.name = aTesterName
    If aTesterInitials <> "" Then
        aTester.initials = aTesterInitials
    Else
        aTester.initials = Me.calculateInitials(aTesterName)
    End If
    
    testers.Add aTester
    
    Exit Sub
    
noCreateTester:
    aTester = Nothing
    MsgBox "error creating tester " & aTesterName & " " & Err.Number & " " & Err.Description
End Sub

Public Function calculateInitials(from As String) As String

    Dim names() As String
    Dim initials As String
    Dim nextInitial
    Dim aCount As Integer
    
    If from = "" Then Exit Function
    
    names = Split(from, " ")
    aCount = 0
    Do While aCount <= UBound(names)
            nextInitial = UCase$(Left$(names(aCount), 1))
            If nextInitial <> "" Then
                initials = initials & nextInitial
            End If
        aCount = aCount + 1
        
        If Len(initials) > 2 Then Exit Do
    Loop

    aCount = Len(initials)
    Do While aCount <= 2
        initials = initials & "X"
        aCount = aCount + 1
    Loop


    calculateInitials = initials
End Function
