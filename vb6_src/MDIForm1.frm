VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Session Manager"
   ClientHeight    =   6960
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6735
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCreateFromCoverageSession 
         Caption         =   "New Session"
      End
      Begin VB.Menu mnuOpenSession 
         Caption         =   "&Open Session"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()

    Dim aForm As frmSessionEdit
    
    On Error Resume Next
    
    Set theTesters = New testers
    setDefaultPaths
    loadFilePaths
    theTesters.loadFromFile testerNamesFileName
    
    
End Sub

Private Sub mnuCreateFromCoverageSession_Click()
    Dim aForm As Form
    
On Error GoTo handleCreateError

    Set aForm = New frmSessionEdit
    aForm.sessionID = getNextSessionID
    aForm.Height = theMDIForm.ScaleHeight * 0.75
    aForm.Width = theMDIForm.ScaleWidth * 0.75
    aForm.Show
    
    Exit Sub
    
handleCreateError:
    If Err.Number <> 32755 Then
        MsgBox "ERROR: " & Err.Number & " " & Err.Description
    End If
End Sub

Private Sub mnuOpenSession_Click()

On Error GoTo handleLoadError
    'find session
    
    
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNFileMustExist
        CommonDialog1.FileName = localSessionsFileName & "*.ses"
        CommonDialog1.Filter = "Session Files (*.ses)|*.ses|All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.ShowOpen
  
        If CommonDialog1.FileName <> "" Then 'load session
            loadSession CommonDialog1.FileName
        End If

    Exit Sub
    
handleLoadError:
    If Err.Number <> 32755 Then
        MsgBox "ERROR: " & Err.Number & " " & Err.Description
    End If
    
    'load session
    'show form

End Sub


