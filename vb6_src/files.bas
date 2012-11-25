Attribute VB_Name = "files"
Option Explicit
Public testerNamesFileName As String
Public localSessionsFileName As String
Public localdatafilesFileName As String
Public networksessionsFileName As String
Public networkdatafilesFileName As String
Public syntaxcheckerFileName As String
Public coverageiniFileName As String
Public pathsFileName As String


Public Function setDefaultPaths()

    testerNamesFileName = App.Path & "\testers.txt"
    localSessionsFileName = App.Path
    localdatafilesFileName = App.Path
    networksessionsFileName = "C:\sessions\submitted\"
    networkdatafilesFileName = "C:\sessions\datafiles\"
    syntaxcheckerFileName = "C:\sessions\scan-submitted-only.bat"
    coverageiniFileName = "C:\sessions\coverage.ini"

End Function

Public Function loadFilePaths()

    Dim fileNum As Long
    Dim readLine As String
    Dim readPathName As String
    Dim readPath As String
    Dim fileIsOpen As Boolean
    
On Error GoTo errorOnLoad

    fileNum = FreeFile
    fileIsOpen = False
    
    pathsFileName = App.Path & "\paths.txt"

    Open pathsFileName For Input As fileNum
    fileIsOpen = True
    
    Do Until EOF(fileNum)
        Line Input #fileNum, readLine
        If Left$(readLine, 1) = "#" Then
            readPathName = Right$(readLine, Len(readLine) - 1)
        ElseIf readPathName <> "" Then
            readPath = Trim$(readLine)
        End If
        
        If readPathName <> "" And readPath <> "" Then
            Select Case readPathName
                Case "coverageini"
                    coverageiniFileName = readPath
                Case "localsessions"
                    If Right$(readPath, 1) <> "\" Then
                        readPath = readPath & "\"
                    End If
                    localSessionsFileName = readPath
                Case "localdatafiles"
                    If Right$(readPath, 1) <> "\" Then
                        readPath = readPath & "\"
                    End If
                    localdatafilesFileName = readPath
                Case "networksessions"
                    If Right$(readPath, 1) <> "\" Then
                        readPath = readPath & "\"
                    End If
                    networksessionsFileName = readPath
                Case "networkdatafiles"
                    If Right$(readPath, 1) <> "\" Then
                        readPath = readPath & "\"
                    End If
                    networkdatafilesFileName = readPath
                Case "syntaxchecker"
                    syntaxcheckerFileName = readPath
                Case "testerNamesFile"
                    testerNamesFileName = readPath
            End Select
            
            readPathName = ""
            readPath = ""
            
        End If
        
    Loop
    
    Close fileNum
    fileIsOpen = False
    
    Exit Function
errorOnLoad:
    If fileIsOpen Then
        Close fileNum
        fileIsOpen = False
    End If
    
End Function

