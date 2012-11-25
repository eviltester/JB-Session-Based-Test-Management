VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSessionEdit 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   6165
   Begin VB.Frame frameDetails 
      Caption         =   "Frame1"
      Height          =   2175
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Tag             =   "2175"
      Top             =   1920
      Width           =   5175
      Begin VB.Frame frameStart 
         Caption         =   "Start"
         Height          =   735
         Left            =   0
         TabIndex        =   41
         Top             =   240
         Width           =   2535
         Begin MSComCtl2.DTPicker DTPickerStart 
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MM/dd/yy hh:mm"
            Format          =   22806531
            CurrentDate     =   37273.0840277778
         End
      End
      Begin VB.Frame frameDuration 
         Caption         =   "Duration"
         Height          =   495
         Left            =   0
         TabIndex        =   39
         Top             =   1560
         Width           =   2535
         Begin VB.TextBox txtMultiplier 
            Height          =   285
            Left            =   840
            TabIndex        =   63
            Top             =   120
            Width           =   255
         End
         Begin VB.ComboBox cmbDuration 
            Height          =   315
            ItemData        =   "Form1.frx":0000
            Left            =   1200
            List            =   "Form1.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "*"
            Height          =   255
            Left            =   1080
            TabIndex        =   62
            Top             =   120
            Width           =   135
         End
      End
      Begin VB.Frame frameCharterVOp 
         Caption         =   "Charter Vs. Opportunity"
         Height          =   615
         Left            =   0
         TabIndex        =   35
         Top             =   960
         Width           =   2535
         Begin VB.TextBox txtCharterVOp 
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "100/0"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   255
            Left            =   1576
            TabIndex        =   37
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "cmbDuration"
            BuddyDispid     =   196613
            OrigLeft        =   2280
            OrigTop         =   120
            OrigRight       =   2520
            OrigBottom      =   615
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
         End
      End
      Begin VB.Frame frameBreakdown 
         Height          =   1815
         Left            =   2520
         TabIndex        =   22
         Top             =   240
         Width           =   2415
         Begin VB.TextBox txt0to100 
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   25
            Text            =   "Text3"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt0to100 
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   24
            Text            =   "Text3"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txt0to100 
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   23
            Text            =   "Text3"
            Top             =   1440
            Width           =   495
         End
         Begin MSComctlLib.Slider slider0to100 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
         End
         Begin MSComctlLib.Slider slider0to100 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
         End
         Begin MSComctlLib.Slider slider0to100 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   29
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "cmbDuration"
            BuddyDispid     =   196613
            OrigLeft        =   2280
            OrigTop         =   120
            OrigRight       =   2520
            OrigBottom      =   615
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   30
            Top             =   960
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "cmbDuration"
            BuddyDispid     =   196613
            OrigLeft        =   2280
            OrigTop         =   120
            OrigRight       =   2520
            OrigBottom      =   615
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   31
            Top             =   1440
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "cmbDuration"
            BuddyDispid     =   196613
            OrigLeft        =   2280
            OrigTop         =   120
            OrigRight       =   2520
            OrigBottom      =   615
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin VB.Label Label5 
            Caption         =   "Design && Exec"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Bug Investigate && Rprt"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Session Setup"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   1695
         End
      End
      Begin VB.CommandButton btnShowHide 
         Height          =   255
         Index           =   3
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblHiddenFrameName 
         Caption         =   "Label4"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   56
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Frame1"
      Height          =   855
      Index           =   7
      Left            =   0
      TabIndex        =   49
      Top             =   4800
      Width           =   2655
      Begin VB.CommandButton btnShowHide 
         Height          =   255
         Index           =   7
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   0
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid msflexIssues 
         Height          =   975
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1720
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.Label lblHiddenFrameName 
         Caption         =   "Label4"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   60
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Frame1"
      Height          =   615
      Index           =   6
      Left            =   0
      TabIndex        =   47
      Top             =   4560
      Width           =   2775
      Begin VB.CommandButton btnShowHide 
         Height          =   255
         Index           =   6
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid msflexBugs 
         Height          =   495
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.Label lblHiddenFrameName 
         Caption         =   "Label4"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   59
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Frame1"
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   43
      Top             =   4320
      Width           =   4455
      Begin VB.TextBox txtTestNotes 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   46
         Text            =   "Form1.frx":0004
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnShowHide 
         Height          =   255
         Index           =   5
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblHiddenFrameName 
         Caption         =   "Label4"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   58
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Frame1"
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   19
      Top             =   3960
      Width           =   4335
      Begin VB.TextBox invisibleFileTextBox 
         Height          =   285
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   61
         Text            =   "Form1.frx":000A
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton btnShowHide 
         Height          =   255
         Index           =   4
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid msflexDataFiles 
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1085
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.Label lblHiddenFrameName 
         Caption         =   "Label4"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   57
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Frame1"
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   17
      Top             =   1440
      Width           =   4575
      Begin VB.TextBox txtAreas 
         Height          =   855
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Form1.frx":0010
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton btnShowHide 
         Height          =   255
         Index           =   2
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin MSComctlLib.TreeView treeAreas 
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1508
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label lblHiddenFrameName 
         Caption         =   "Label4"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   55
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Frame1"
      Height          =   855
      Index           =   1
      Left            =   -120
      TabIndex        =   13
      Top             =   600
      Width           =   5175
      Begin VB.TextBox txtCharter 
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Text            =   "Form1.frx":0016
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton btnShowHide 
         Height          =   255
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblHiddenFrameName 
         Caption         =   "Label4"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   54
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Frame1"
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Tag             =   "735"
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtSessionID 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtInitials 
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnShowHide 
         Height          =   255
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblHiddenFrameName 
         Caption         =   "Label4"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   53
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Session"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Initials"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Tester"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   5880
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6195
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image tbShow 
      Height          =   225
      Left            =   5280
      Picture         =   "Form1.frx":001C
      Top             =   360
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image tbHide 
      Height          =   240
      Left            =   5280
      Picture         =   "Form1.frx":032E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmSessionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nextIssueID As Long
Public testerInitials As String
Public testername As String
Public testerNames As Collection
Public sessionCount As Integer
Public charterText As String
Public testNotes As String
Public loadedSessionAreas As Collection
Public loadedDataFiles As Collection
Public loadedBugs As Collection
Public loadedIssues As Collection
Public startDate As String

Public duration As String
Public testDesign As String
Public bugInvest As String
Public sessionSetup As String
Public charterVop As String

Public sessionID As Long 'a unique id of this form for
                         ' this instance of the system

Private frameNames(0 To 10) As String

Const mustAddUpTo As Integer = 100
Const maxCellHeight As Integer = 5

Public Function getNextIssueID() As Long
    getNextIssueID = nextIssueID
    nextIssueID = nextIssueID + 1
End Function

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()

    On Error GoTo handleSaveError
    
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNPathMustExist & cdlOFNOverwritePrompt
        CommonDialog1.FileName = localSessionsFileName & Me.Caption & ".ses"
        CommonDialog1.Filter = "Session Files (*.ses)|*.ses"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.ShowSave
  
        If CommonDialog1.FileName <> "" Then 'save a new file
            writeSessionTo CommonDialog1.FileName
        End If

    Exit Sub
    
handleSaveError:
    If Err.Number <> 32755 Then
        MsgBox "ERROR: " & Err.Number & " " & Err.Description
    End If
End Sub

Public Function writeSessionTo(fileNamePath As String)

    Dim aFileNum As Integer
    Dim aTreeNode As Node
    Dim anIter As Long
    Dim aCount As Long
    
    On Error GoTo writeError
    
    aFileNum = FreeFile
    Open fileNamePath For Output As aFileNum
    
        Print #aFileNum, "CHARTER"
        Print #aFileNum, "-----------------------------------------------"
        Print #aFileNum, txtCharter.Text
        Print #aFileNum, ""
        
        Print #aFileNum, "#AREAS"
        For Each aTreeNode In treeAreas.Nodes
            If aTreeNode.Checked = True Then
                If aTreeNode.Key <> "R" Then
                    Print #aFileNum, Right$(aTreeNode.Key, Len(aTreeNode.Key) - 2)
                End If
            End If
        Next
        Print #aFileNum, ""
    
        Print #aFileNum, "START"
        Print #aFileNum, "-----------------------------------------------"
        Print #aFileNum, Format$(DTPickerStart.Value, "mm/dd/yy Hh:Nnam/pm")
        Print #aFileNum, ""
        
        Print #aFileNum, "TESTER"
        Print #aFileNum, "-----------------------------------------------"
        Print #aFileNum, Combo1.Text
        Print #aFileNum, ""
        
        Print #aFileNum, "TASK BREAKDOWN"
        Print #aFileNum, "-----------------------------------------------"
        
        Print #aFileNum, ""
        Print #aFileNum, "#DURATION"
        anIter = InStr(1, cmbDuration.Text, " ")
        If txtMultiplier.Text <> "" Then
            Print #aFileNum, Trim$(Left$(cmbDuration.Text, anIter)) & " * " & txtMultiplier.Text
        Else
            Print #aFileNum, Trim$(Left$(cmbDuration.Text, anIter))
        End If
        
        Print #aFileNum, ""
        Print #aFileNum, "#TEST DESIGN AND EXECUTION"
        Print #aFileNum, txt0to100(0).Text
        
        Print #aFileNum, ""
        Print #aFileNum, "#BUG INVESTIGATION AND REPORTING"
        Print #aFileNum, txt0to100(1).Text
        
        Print #aFileNum, ""
        Print #aFileNum, "#SESSION SETUP"
        Print #aFileNum, txt0to100(2).Text
        
        Print #aFileNum, ""
        Print #aFileNum, "#CHARTER VS. OPPORTUNITY"
        Print #aFileNum, txtCharterVOp.Text
        Print #aFileNum, ""
        
        Print #aFileNum, "DATA FILES"
        Print #aFileNum, "-----------------------------------------------"
        
        aCount = 0
        For anIter = 1 To msflexDataFiles.Rows - 1
            If msflexDataFiles.TextMatrix(anIter, 0) <> "" And msflexDataFiles.TextMatrix(anIter, 0) <> createNewThangSymbol Then
                aCount = aCount + 1
                Print #aFileNum, msflexDataFiles.TextMatrix(anIter, 0)
            End If
        Next
        If aCount = 0 Then
            Print #aFileNum, "#N/A"
        End If
        Print #aFileNum, ""
        
        Print #aFileNum, "TEST NOTES"
        Print #aFileNum, "-----------------------------------------------"
        If txtTestNotes.Text = "" Then
            Print #aFileNum, "#N/A"
        Else
            Print #aFileNum, txtTestNotes.Text
        End If
        Print #aFileNum, ""
        
        Print #aFileNum, "BUGS"
        Print #aFileNum, "-----------------------------------------------"
        
        aCount = 0
        For anIter = 1 To msflexBugs.Rows - 1
            If msflexBugs.TextMatrix(anIter, 1) <> "" And msflexBugs.TextMatrix(anIter, 0) <> createNewThangSymbol Then
                aCount = aCount + 1
                Print #aFileNum, "#BUG " & msflexBugs.TextMatrix(anIter, 0)
                Print #aFileNum, msflexBugs.TextMatrix(anIter, 1)
                Print #aFileNum, ""
            End If
        Next
        If aCount = 0 Then
            Print #aFileNum, "#N/A"
        End If
        Print #aFileNum, ""
        
        Print #aFileNum, "ISSUES"
        Print #aFileNum, "-----------------------------------------------"
        
        aCount = 0
        For anIter = 1 To msflexIssues.Rows - 1
            If msflexIssues.TextMatrix(anIter, 1) <> "" And msflexIssues.TextMatrix(anIter, 0) <> createNewThangSymbol Then
                aCount = aCount + 1
                Print #aFileNum, "#ISSUE " & aCount 'msflexIssues.TextMatrix(anIter, 0)
                Print #aFileNum, msflexIssues.TextMatrix(anIter, 1)
                Print #aFileNum, ""
            End If
        Next
        If aCount = 0 Then
            Print #aFileNum, "#N/A"
        End If
        Print #aFileNum, ""
        
    Close aFileNum

    writeSessionTo = True
    Exit Function

writeError:

    Close aFileNum
    MsgBox "Error: The file may not have been written fully" & vbCrLf & Err.Number & " " & Err.Description
    writeSessionTo = False
End Function

Private Sub btnShowHide_Click(Index As Integer)

    Dim theBtnShowHide As CommandButton
    Dim theFrameDetails As Frame
    Dim theLabelFrameName As Label
    Dim aTreeNode As Node
    Dim aCount As Long
    Dim anIter As Long
    
    On Error Resume Next
    
    Set theBtnShowHide = btnShowHide(Index)
    Set theFrameDetails = frameDetails(Index)
    Set theLabelFrameName = lblHiddenFrameName(Index)
    
    If theBtnShowHide.Picture = tbShow Then
        theBtnShowHide.Picture = tbHide
        theFrameDetails.BorderStyle = 1
        fillFreeSpace
        'theLabelFrameName.Visible = False
    Else
        theFrameDetails.Height = theBtnShowHide.Height - 10
        theFrameDetails.BorderStyle = 0
        theBtnShowHide.Picture = tbShow
        fillFreeSpace
        showLabelFrameName theLabelFrameName, theBtnShowHide
        
        'now fix the label details
        theLabelFrameName.Caption = frameNames(Index) & ": "
        Select Case Index
        Case 0  ' session details
            theLabelFrameName.Caption = theLabelFrameName.Caption & " " & Combo1.Text & " (" & txtInitials.Text & ") [" & txtSessionID.Text & "]"
        Case 1  ' charter - show the first line
            theLabelFrameName.Caption = theLabelFrameName.Caption & Replace(Left$(LTrim$(txtCharter.Text), 50), vbCrLf, "")
        Case 2  ' areas - count selected
            aCount = 0
            For Each aTreeNode In treeAreas.Nodes
                If aTreeNode.Checked = True Then
                    aCount = aCount + 1
                End If
            Next
            theLabelFrameName.Caption = theLabelFrameName.Caption & aCount & " Areas Selected"
        Case 3
            theLabelFrameName.Caption = theLabelFrameName.Caption & cmbDuration.Text & " [" & DTPickerStart.Value & "] CvO (" & txtCharterVOp & ") T/B/S: " & txt0to100(0) & "/" & txt0to100(1) & "/" & txt0to100(2)
        Case 4
            aCount = 0
            For anIter = 1 To msflexDataFiles.Rows - 1
                If msflexDataFiles.TextMatrix(anIter, 0) <> createNewThangSymbol Then
                    aCount = aCount + 1
                End If
            Next
            theLabelFrameName.Caption = theLabelFrameName.Caption & aCount & " Data Files"
        Case 5
            theLabelFrameName.Caption = theLabelFrameName.Caption & Replace(Left$(LTrim$(txtTestNotes.Text), 50), vbCrLf, "")
        End Select
        
    End If
    
End Sub


Public Sub showLabelFrameName(theLabelFrameName As Label, theBtnShowHide As CommandButton)

        theLabelFrameName.Top = theBtnShowHide.Top
        theLabelFrameName.Left = theBtnShowHide.Left + theBtnShowHide.Width + 20
        theLabelFrameName.Height = 230
        theLabelFrameName.BackColor = vbActiveTitleBar
        theLabelFrameName.ForeColor = vbTitleBarText
        theLabelFrameName.Visible = True
        theLabelFrameName.ZOrder

End Sub


Private Sub Combo1_Change()
    On Error Resume Next
    txtInitials.Text = theTesters.calculateInitials(Combo1.Text)
    setFormCaption
End Sub

Private Sub Combo1_LostFocus()
    On Error Resume Next
    txtInitials.Text = theTesters.calculateInitials(Combo1.Text)
    setFormCaption
End Sub

Private Sub Form_Initialize()

    nextIssueID = 1
    sessionCount = 0
    Set loadedSessionAreas = New Collection
    Set loadedDataFiles = New Collection
    Set loadedBugs = New Collection
    Set loadedIssues = New Collection
End Sub

Private Sub Form_Load()

    Dim controlIter As Integer
    Dim aTester As tester
    Dim aDelimiterAt As Long
    
    On Error Resume Next
    
    For controlIter = 0 To btnShowHide.UBound
        btnShowHide(controlIter).Picture = tbShow
    Next
    For controlIter = 0 To frameDetails.UBound
        frameDetails(controlIter).Height = btnShowHide(controlIter).Height
        frameDetails(controlIter).Left = 0
        frameDetails(controlIter).Width = Me.ScaleWidth - frameDetails(controlIter).Left
        frameDetails(controlIter).BorderStyle = 0
    Next
    
    frameNames(0) = "Session Details"
    frameNames(1) = "Charter"
    frameNames(2) = "Areas"
    frameNames(3) = "Task Breakdown"
    frameNames(4) = "Data Files"
    frameNames(5) = "Test Notes"
    frameNames(6) = "Bugs"
    frameNames(7) = "Issues"
    
    For controlIter = 0 To frameDetails.UBound
        'frameDetails(controlIter).Caption = "   " & frameNames(controlIter)
        frameDetails(controlIter).Caption = ""
        lblHiddenFrameName(controlIter).Caption = frameNames(controlIter)
        showLabelFrameName lblHiddenFrameName(controlIter), btnShowHide(controlIter)
    Next

    With msflexDataFiles
        .Rows = 0
        .Cols = 1
        .FormatString = "<Filename"
        .AddItem ("FileName")
        .AddItem (createNewThangSymbol)
        .FixedRows = 1
        .FixedCols = 0
    End With
    
    For controlIter = 1 To loadedDataFiles.Count
        msflexDataFiles.AddItem loadedDataFiles(controlIter)
    Next
    
    If loadedDataFiles.Count <> 0 Then
        msflexDataFiles.AddItem createNewThangSymbol
    End If
    
    AutosizeGridColumns msflexDataFiles, msflexDataFiles.Rows, msflexDataFiles.Width, Me
    
    With msflexBugs
        .Rows = 0
        .Cols = 2
        .AddItem ("ID" & vbTab & "Details")
        .FormatString = "<ID|<Details"
        .AddItem (createNewThangSymbol)
        .FixedRows = 1
        .FixedCols = 1
    End With
    For controlIter = 1 To loadedBugs.Count
        aDelimiterAt = InStr(1, loadedBugs(controlIter), "|")
        msflexBugs.AddItem Left$(loadedBugs(controlIter), aDelimiterAt - 1) & vbTab & Mid$(loadedBugs(controlIter), aDelimiterAt + 1)
        msflexBugs.RowData(msflexBugs.Rows - 1) = getNextInternalBugID
    Next
    
    If loadedBugs.Count <> 0 Then
        msflexBugs.AddItem createNewThangSymbol
    End If
    
    
    AutosizeGridColumns msflexBugs, msflexBugs.Rows, msflexBugs.Width, Me
    
    With msflexIssues
        .Rows = 0
        .Cols = 2
        .AddItem ("ID" & vbTab & "Details")
        .AddItem (createNewThangSymbol)
        .FormatString = "<ID|<Details"
        .FixedRows = 1
        .FixedCols = 1
    End With
    For controlIter = 1 To loadedIssues.Count
        aDelimiterAt = InStr(1, loadedIssues(controlIter), "|")
        msflexIssues.AddItem Left$(loadedIssues(controlIter), aDelimiterAt - 1) & vbTab & Mid$(loadedIssues(controlIter), aDelimiterAt + 1)
        msflexIssues.RowData(msflexIssues.Rows - 1) = getNextinternalIssueID
    Next
    
    If loadedIssues.Count <> 0 Then
        msflexIssues.AddItem createNewThangSymbol
    End If
    AutosizeGridColumns msflexIssues, msflexIssues.Rows, msflexIssues.Width, Me
    
    'load the existing tester names into the combo
    For Each aTester In theTesters.testers
        Combo1.AddItem aTester.name
    Next
    
    Combo1.Text = testername
    txtInitials = testerInitials
    txtSessionID.Text = Chr$(Asc("A") + sessionCount)
    txtCharter.Text = charterText
    txtTestNotes.Text = testNotes

    txtAreas.Text = ""
    
    Slider1.Max = 100
    If charterVop <> "" Then
        Slider1.Value = Left$(charterVop, InStr(1, charterVop, "/") - 1)
    Else
        Slider1.Value = 100
    End If
    
    cmbDuration.AddItem ("short (60)")
    cmbDuration.AddItem ("normal (90)")
    cmbDuration.AddItem ("long (120)")
    If duration = "" Then
        cmbDuration.ListIndex = 0
    Else
        'parse the duration
        If InStr(1, duration, " * ") <> 0 Then
            txtMultiplier = Mid$(duration, InStr(1, duration, " * ") + 3)
        End If
        Select Case Left$(duration, 1)
        Case "s"
            cmbDuration.ListIndex = 0
        Case "n"
            cmbDuration.ListIndex = 1
        Case "l"
            cmbDuration.ListIndex = 2
        End Select
    End If
    
    slider0to100(0).Max = mustAddUpTo
    slider0to100(1).Max = mustAddUpTo
    slider0to100(2).Max = mustAddUpTo
    
    If testDesign <> "" Then
        slider0to100(0).Value = testDesign
    End If
    If bugInvest <> "" Then
        slider0to100(1).Value = bugInvest
    End If
    If sessionSetup <> "" Then
        slider0to100(2).Value = sessionSetup
    End If
    
    keepSlidersInStep (0)
    
    If startDate = "" Then
        DTPickerStart.Value = Format(Now, "MM/dd/yy h:m")
    Else
        DTPickerStart.Value = startDate
    End If
    
    loadAreas (coverageiniFileName)
    buildTreeOfAreas
    fillFreeSpace
    setFormCaption
End Sub

Public Sub keepSlidersInStep(mainSlider As Integer)
    'all the sliders must add up to 100
    'keep the mainslider steady
    
    Dim currTotal As Integer
    Dim adjustBy As Integer
    Dim adjustIterator As Integer
    Dim adjustThis As Integer
    Dim numberOfSliders As Integer
    Dim newval As Integer
    
    On Error Resume Next
    
    currTotal = slider0to100(0).Value + slider0to100(1).Value + slider0to100(2).Value
    adjustBy = mustAddUpTo - currTotal
    
    'only ever adjust the one under it, unless that is at 0 or 100 at which point use the next one
    
    numberOfSliders = slider0to100.Count
    For adjustIterator = 1 To numberOfSliders - 1
        adjustThis = (mainSlider + adjustIterator) Mod numberOfSliders
        newval = slider0to100(adjustThis).Value + adjustBy
        'this will lose any excess
        slider0to100(adjustThis).Value = slider0to100(adjustThis).Value + adjustBy
        
        If newval < 0 Then
            adjustBy = newval
        ElseIf newval > 100 Then
            adjustBy = newval - 100
        Else
            adjustBy = 0
        End If
    Next
    
    For adjustIterator = 0 To slider0to100.UBound
        txt0to100(adjustIterator).Text = slider0to100(adjustIterator).Value
    Next
    
    

End Sub

Public Function buildTreeOfAreas()

    Dim areaCount As Long
    Dim loadedAreaCount As Long
    Dim lastArea As Long
    Dim splitArea() As String
    Dim splitCount As Integer
    Dim Key As String
    Dim tempnode As Node
    Dim selectedArea As Boolean
    
    On Error Resume Next
    lastArea = areas.Count

    treeAreas.Nodes.Clear
    treeAreas.Checkboxes = True
    Set tempnode = treeAreas.Nodes.Add(, , "R", "Areas")
    If (tempnode Is Nothing) = False Then tempnode.Tag = "Areas"
    
    For areaCount = 1 To lastArea
    
        selectedArea = False
        For loadedAreaCount = 1 To loadedSessionAreas.Count
            If loadedSessionAreas(loadedAreaCount) = areas(areaCount) Then
                selectedArea = True
                Exit For
            End If
        Next
        
        splitArea = Split(areas(areaCount), "|")
        Key = "R"
        For splitCount = 0 To UBound(splitArea)
            Set tempnode = treeAreas.Nodes.Add(Key, tvwChild, Key & "|" & Trim$(splitArea(splitCount)), Trim$(splitArea(splitCount)))
            If splitCount = UBound(splitArea) Then
                If (tempnode Is Nothing) = False Then
                    tempnode.Tag = "Branch"
                    If selectedArea Then
                        tempnode.Checked = True
                    End If
                End If
            Else
                If (tempnode Is Nothing) = False Then tempnode.Tag = "Area"
            End If
            Key = Key & "|" & Trim$(splitArea(splitCount))
        Next
    Next
    

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'check for child bug or issue forms
    
End Sub

Public Sub addEditBug(theUniqueID As Long, theID As String, theDetails As String)
    Dim writeID As String
    Dim atRow As Long
    Dim anIter As Long
    Dim foundExisting As Boolean
    
    On Error Resume Next
    
    If theUniqueID = -1 Then
        'it is a new one
                '-1 is one before the last <...>
        atRow = msflexBugs.Rows - 1
        
    Else
        'an amendment
        'find the bug
        foundExisting = False
        For anIter = 1 To msflexBugs.Rows - 1
            If msflexBugs.RowData(anIter) = theUniqueID Then
                foundExisting = True
                Exit For
            End If
        Next
        
        If foundExisting Then
            atRow = anIter
            'remove the existing one
            msflexBugs.RemoveItem atRow
        Else
            If MsgBox("the bug has been deleted, do you wish to recreate it?", vbYesNo) = vbYes Then
                atRow = msflexBugs.Rows - 1
            Else
                Exit Sub
            End If
        End If
    End If
    
        If theID = "" Then
            writeID = "<not_entered>"
        Else
            writeID = theID
        End If

        msflexBugs.AddItem writeID & vbTab & theDetails, atRow
        If theUniqueID = -1 Then
            msflexBugs.RowData(atRow) = getNextInternalBugID
        Else
            msflexBugs.RowData(atRow) = theUniqueID
        End If
        AutosizeGridColumns msflexBugs, msflexBugs.Rows - 1, msflexBugs.Width, Me
        invisibleFileTextBox.Width = msflexBugs.ColWidth(1)
        ReSizeCellHeight msflexBugs, atRow, 1, Me, invisibleFileTextBox, maxCellHeight
    
    
    If msflexBugs.TextMatrix(1, 0) <> createNewThangSymbol Then
        msflexBugs.AddItem createNewThangSymbol, 1
    End If
End Sub

Public Sub addEditIssue(theUniqueID As Long, theID As String, theDetails As String)
    Dim writeID As String
    Dim atRow As Long
    Dim anIter As Long
    Dim foundExisting As Boolean
    
    On Error Resume Next
    
    If theUniqueID = -1 Then
        'it is a new one
                '-1 is one before the last <...>
        atRow = msflexIssues.Rows - 1
        
    Else
        'an amendment
        'find the bug
        foundExisting = False
        For anIter = 1 To msflexIssues.Rows - 1
            If msflexIssues.RowData(anIter) = theUniqueID Then
                foundExisting = True
                Exit For
            End If
        Next
        
        If foundExisting Then
            atRow = anIter
            'remove the existing one
            msflexIssues.RemoveItem atRow
        Else
            If MsgBox("the issue has been deleted, do you wish to recreate it?", vbYesNo) = vbYes Then
                atRow = msflexIssues.Rows - 1
            Else
                Exit Sub
            End If
        End If
    End If
    
        If theID = "[Auto]" Then
            writeID = getNextIssueID
        Else
            writeID = theID
        End If
        
        msflexIssues.AddItem writeID & vbTab & theDetails, atRow
        If theUniqueID = -1 Then
            msflexIssues.RowData(atRow) = getNextinternalIssueID
        Else
            msflexIssues.RowData(atRow) = theUniqueID
        End If
        AutosizeGridColumns msflexIssues, msflexIssues.Rows - 1, msflexIssues.Width, Me
        invisibleFileTextBox.Width = msflexIssues.ColWidth(1)
        ReSizeCellHeight msflexIssues, atRow, 1, Me, invisibleFileTextBox, maxCellHeight
    
    
    If msflexIssues.TextMatrix(1, 0) <> createNewThangSymbol Then
        msflexIssues.AddItem createNewThangSymbol, 1
    End If
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    
    Dim anInter As Long
    
    btnSave.Top = Me.ScaleHeight - StatusBar1.Height - btnSave.Height - 50
    btnCancel.Top = btnSave.Top
    
    fillFreeSpace
    
    
    
    txtAreas.Visible = False
    
    'fix their contents
    txtCharter.Width = frameDetails(1).Width - (txtCharter.Left * 2)
    treeAreas.Width = frameDetails(2).Width - (treeAreas.Left * 2)
    msflexDataFiles.Width = frameDetails(3).Width - (msflexDataFiles.Left * 2)
    invisibleFileTextBox.Width = msflexDataFiles.Width
    txtTestNotes.Width = frameDetails(4).Width - (txtTestNotes.Left * 2)
    msflexBugs.Width = frameDetails(5).Width - (msflexBugs.Left * 2)
    msflexIssues.Width = frameDetails(6).Width - (msflexIssues.Left * 2)
    
    For anInter = 0 To lblHiddenFrameName.UBound
        lblHiddenFrameName(anInter).Width = frameDetails(anInter).Width - btnShowHide(anInter).Width
    Next
    
    AutosizeGridColumns msflexBugs, msflexBugs.Rows, msflexBugs.Width, Me
    AutosizeGridColumns msflexIssues, msflexIssues.Rows, msflexIssues.Width, Me
    msflexDataFiles.ColWidth(0) = msflexIssues.Width
    'AutosizeGridColumns msflexDataFiles, msflexDataFiles.Rows, msflexDataFiles.Width, Me
    
    For anInter = 1 To msflexDataFiles.Rows - 1
        ReSizeCellHeight msflexDataFiles, anInter, 0, Me, invisibleFileTextBox, maxCellHeight
    Next
End Sub

Public Sub setFormCaption()
On Error Resume Next
    Me.Caption = "ET-" & txtInitials.Text & "-" & Format(Date, "YYMMDD") & "-" & txtSessionID.Text
End Sub

Private Sub fillFreeSpace()

    On Error Resume Next
    
    Dim frameControlIter As Long
    Dim freespace As Long
    Dim countitems As Long
    
    'open any which are fixed and open to their fixed and open sizes
    For frameControlIter = 0 To frameDetails.UBound
        If frameDetails(frameControlIter).Tag <> "" Then
            If IsNumeric(frameDetails(frameControlIter).Tag) And btnShowHide(frameControlIter).Picture = tbHide Then
                frameDetails(frameControlIter).Height = frameDetails(frameControlIter).Tag
            End If
        End If
        
        'as a sideeffect, widen all the frames
        frameDetails(frameControlIter).Width = Me.ScaleWidth - frameDetails(frameControlIter).Left
        frameDetails(frameControlIter).Left = 0
    Next
    
    'put then all in a row
    frameDetails(0).Top = 0
    For frameControlIter = 1 To frameDetails.UBound
        frameDetails(frameControlIter).Top = frameDetails(frameControlIter - 1).Height + frameDetails(frameControlIter - 1).Top
    Next
    
    ' the rest share the space
    ' is there any?
    freespace = btnSave.Top - (frameDetails(frameDetails.UBound).Top + frameDetails(frameDetails.UBound).Height)
        
    'find out how much space the others are using, divide it by 3
    ' and apportion it evenly
    countitems = 0
    For frameControlIter = 0 To frameDetails.UBound
        If frameDetails(frameControlIter).Tag = "" Then
            If btnShowHide(frameControlIter).Picture = tbHide Then
                freespace = freespace + frameDetails(frameControlIter).Height
                countitems = countitems + 1
            End If
        End If
    Next
        
    If countitems = 0 Then Exit Sub
    
    freespace = freespace / countitems
    
    For frameControlIter = 0 To frameDetails.UBound
        If frameDetails(frameControlIter).Tag = "" Then
            If btnShowHide(frameControlIter).Picture = tbHide Then
                frameDetails(frameControlIter).Height = freespace
            End If
        End If
        
    Next
    
    
    'put them all in a row
    frameDetails(0).Top = 0
    For frameControlIter = 1 To frameDetails.UBound
        frameDetails(frameControlIter).Top = frameDetails(frameControlIter - 1).Height + frameDetails(frameControlIter - 1).Top
    Next
    
    ' Fix the contents height
    txtCharter.Height = frameDetails(1).Height - btnShowHide(1).Height - 50
    treeAreas.Height = frameDetails(2).Height - btnShowHide(2).Height - 50
    msflexDataFiles.Height = frameDetails(4).Height - btnShowHide(4).Height - 50
    txtTestNotes.Height = frameDetails(5).Height - btnShowHide(5).Height - 50
    msflexBugs.Height = frameDetails(6).Height - btnShowHide(6).Height - 50
    msflexIssues.Height = frameDetails(7).Height - btnShowHide(7).Height - 50

    
End Sub

Private Sub Form_Terminate()
    Dim aForm As Form
    
    For Each aForm In Forms
        If TypeName(aForm) = "frmTextDetails" Then
            If aForm.localSessionID = sessionID Then
                Unload aForm
            End If
        End If
    Next
    
End Sub

'Private Sub fillFreeSpace2()
'
'    On Error Resume Next
'
'    Dim freespace As Long
'    Dim countitems As Integer
'
'    'the first is fixed
'    If btnCloseSessionDetails.Visible = True Then
'        frameSession.Height = Combo1.Top + Combo1.Height + 50
'    Else
'        frameSession.Height = btnCloseSessionDetails.Height
'    End If
'
'    frameCharter.Top = frameSession.Height + frameSession.Top
'    frameAreas.Top = frameCharter.Height + frameCharter.Top
'    frameMisc.Top = frameAreas.Top + frameAreas.Height
'
'    ' the rest share the space
'    ' is there any?
'    freespace = btnSave.Top - (frameMisc.Top + frameMisc.Height)
'    'If freespace < 0 Then
'        'freespace = 0
'    'End If
'
'    'find out how much space the others are using, divide it by 3
'    ' and apportion it evenly
'    countitems = 0
'    If btnHideCharterDetails.Visible = True Then
'        freespace = freespace + frameCharter.Height
'        countitems = countitems + 1
'    End If
'    If btnHideAreas.Visible = True Then
'        freespace = freespace + frameAreas.Height
'        countitems = countitems + 1
'    End If
'    If btnHideOthers.Visible = True Then
'        freespace = freespace + frameMisc.Height
'        countitems = countitems + 1
'    End If
'
'    If countitems = 0 Then Exit Sub
'
'    freespace = freespace / countitems
'
'    If btnHideCharterDetails.Visible = True Then
'        frameCharter.Height = freespace
'    End If
'    If btnHideOthers.Visible = True Then
'        frameMisc.Height = freespace
'    End If
'    If btnHideAreas.Visible = True Then
'        frameAreas.Height = freespace
'    End If
'
'    frameCharter.Top = frameSession.Height + frameSession.Top
'    frameAreas.Top = frameCharter.Height + frameCharter.Top
'    frameMisc.Top = frameAreas.Top + frameAreas.Height
'
'
'    txtCharter.Height = frameCharter.Height - btnShowCharterDetails.Height - 50
'    treeAreas.Height = frameAreas.Height - btnShowAreas.Height - 50
'    TabStrip1.Height = frameMisc.Height - TabStrip1.Top - 50
'
'End Sub


Private Sub msflexBugs_DblClick()
    Dim currRow As Long
    Dim aForm As Form
    
On Error Resume Next

    currRow = msflexBugs.Row
    If msflexBugs.TextMatrix(currRow, 0) = createNewThangSymbol Then
        ' edit new bug
        'does the form exist?
        For Each aForm In Forms
            If TypeName(aForm) = "frmTextDetails" Then
                If aForm.localSessionID = sessionID And Left$(aForm.Caption, 7) = "New Bug" Then
                    aForm.SetFocus
                    Exit For
                End If
            End If
        Next
        
        If aForm Is Nothing Then
            Set aForm = New frmTextDetails
            aForm.Caption = "New Bug: " & Me.Caption
            aForm.txtDetails.Text = ""
            aForm.txtID.Text = ""
            aForm.txtID.Locked = False
            aForm.txtID.Enabled = True
            aForm.localSessionID = Me.sessionID
            aForm.thisEntity = aBug
            aForm.uniqueID = -1
            aForm.Show
        End If
        
        
    Else
    
        'does the form exist?
        For Each aForm In Forms
            If TypeName(aForm) = "frmTextDetails" Then
                If aForm.localSessionID = sessionID Then
                    If aForm.thisEntity = aBug And aForm.uniqueID = msflexBugs.RowData(msflexBugs.Row) Then
                        aForm.SetFocus
                        Exit For
                    End If
                End If
            End If
        Next

        If aForm Is Nothing Then
            'edit an existing bug
            Set aForm = New frmTextDetails
            aForm.Caption = "Amend Bug: " & Me.Caption
            aForm.txtDetails.Text = msflexBugs.TextMatrix(msflexBugs.Row, 1)
            If msflexBugs.TextMatrix(msflexBugs.Row, 0) = "<not_entered>" Then
                aForm.txtID.Text = ""
            Else
                aForm.txtID.Text = msflexBugs.TextMatrix(msflexBugs.Row, 0)
            End If
            aForm.txtID.Locked = False
            aForm.txtID.Enabled = True
            aForm.localSessionID = Me.sessionID
            aForm.thisEntity = aBug
            aForm.uniqueID = msflexBugs.RowData(msflexBugs.Row)
            aForm.Show
        End If
        
    End If
    
End Sub

Private Sub msflexBugs_KeyUp(KeyCode As Integer, Shift As Integer)

On Error Resume Next
    If KeyCode = 46 Or KeyCode = 8 Then  '46 is backspace 8 is delete
        
        If msflexBugs.TextMatrix(msflexBugs.Row, 0) <> createNewThangSymbol Then
            'delete the item
            If MsgBox("Are you sure you want to delete this Bug?", vbYesNo) = vbYes Then
                msflexBugs.RemoveItem msflexBugs.Row
            End If
        End If
    ElseIf KeyCode = 13 Then    'enter
        msflexBugs_DblClick
    End If

End Sub

Private Sub msflexDataFiles_DblClick()

    Dim currRow As Long
    
    On Error GoTo handleerror
    currRow = msflexDataFiles.Row
    If msflexDataFiles.TextMatrix(currRow, 0) = createNewThangSymbol Then
        ' get new filename
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNFileMustExist
        CommonDialog1.Filter = "All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.ShowOpen
  
        If CommonDialog1.FileName <> "" Then 'add a new file
            If currRow = 1 Then
                msflexDataFiles.AddItem CommonDialog1.FileName, 2
                ReSizeCellHeight msflexDataFiles, 2, 0, Me, invisibleFileTextBox, maxCellHeight
            Else
                msflexDataFiles.AddItem CommonDialog1.FileName, currRow
                ReSizeCellHeight msflexDataFiles, currRow, 0, Me, invisibleFileTextBox, maxCellHeight
            End If
            
            If currRow = 1 And msflexDataFiles.Rows = 3 Then
                msflexDataFiles.AddItem createNewThangSymbol
            End If
        End If
    Else
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNFileMustExist
        CommonDialog1.Filter = "All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.FileName = msflexDataFiles.TextMatrix(currRow, 0)
        CommonDialog1.ShowOpen
  
        If CommonDialog1.FileName <> "" Then 'add rename file
            msflexDataFiles.TextMatrix(currRow, 0) = CommonDialog1.FileName
            ReSizeCellHeight msflexDataFiles, currRow, 0, Me, invisibleFileTextBox, maxCellHeight
        End If
    End If
handleerror:
    'should really only handle error 32755 (cancel)
End Sub

Private Sub msflexDataFiles_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 46 Or KeyCode = 8 Then  '46 is backspace 8 is delete
        
        If msflexDataFiles.TextMatrix(msflexDataFiles.Row, 0) <> createNewThangSymbol Then
            'delete the item
            If MsgBox("Are you sure you want to delete this datafile?", vbYesNo) = vbYes Then
                msflexDataFiles.RemoveItem msflexDataFiles.Row
            End If
        End If
    ElseIf KeyCode = 13 Then    'enter
        msflexDataFiles_DblClick
    End If
End Sub


Private Sub msflexIssues_DblClick()
    Dim currRow As Long
    Dim aForm As Form
    
    On Error Resume Next

    currRow = msflexIssues.Row
    If msflexIssues.TextMatrix(currRow, 0) = createNewThangSymbol Then
        ' edit new bug
        
                'does the form exist?
        For Each aForm In Forms
            If TypeName(aForm) = "frmTextDetails" Then
                If aForm.localSessionID = sessionID And Left$(aForm.Caption, 9) = "New Issue" Then
                    aForm.SetFocus
                    Exit For
                End If
            End If
        Next
        
        If aForm Is Nothing Then

            Set aForm = New frmTextDetails
            aForm.Caption = "New Issue: " & Me.Caption
            aForm.txtDetails.Text = ""
            aForm.txtID.Text = "[Auto]"
            aForm.txtID.Locked = True
            aForm.txtID.Enabled = True
            aForm.localSessionID = Me.sessionID
            aForm.thisEntity = anIssue
            aForm.uniqueID = -1
            aForm.Show
        End If
        

    Else
    
        'does the form exist?
        For Each aForm In Forms
            If TypeName(aForm) = "frmTextDetails" Then
                If aForm.localSessionID = sessionID Then
                    If aForm.thisEntity = anIssue And aForm.uniqueID = msflexBugs.RowData(msflexBugs.Row) Then
                        aForm.SetFocus
                        Exit For
                    End If
                End If
            End If
        Next

        If aForm Is Nothing Then
    
            'edit an existing bug
            Set aForm = New frmTextDetails
            aForm.Caption = "Amend Issue: " & Me.Caption
            aForm.txtDetails.Text = msflexIssues.TextMatrix(msflexIssues.Row, 1)
            aForm.txtID.Text = msflexIssues.TextMatrix(msflexIssues.Row, 0)
            aForm.txtID.Locked = True
            aForm.txtID.Enabled = True
            aForm.localSessionID = Me.sessionID
            aForm.thisEntity = anIssue
            aForm.uniqueID = msflexIssues.RowData(msflexIssues.Row)
            aForm.Show
        End If

    End If
End Sub

Private Sub slider0to100_Change(Index As Integer)
On Error Resume Next
    keepSlidersInStep (Index)
    UpDown2(Index).Value = slider0to100(Index).Value
End Sub

Private Sub Slider1_Change()
On Error Resume Next
    txtCharterVOp = Slider1.Value & "/" & (100 - Slider1.Value)
    UpDown1.Value = Slider1.Value
End Sub

Private Sub txt0to100_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim Cancel As Boolean
    On Error Resume Next
    If KeyAscii = 13 Then   'enter
        txt0to100_Validate Index, Cancel
        If Cancel = False Then
            slider0to100(Index).Value = txt0to100(Index).Text
            keepSlidersInStep (Index)
        End If
    End If
End Sub

Private Sub txt0to100_LostFocus(Index As Integer)
On Error Resume Next
    slider0to100(Index).Value = txt0to100(Index).Text
    keepSlidersInStep (Index)
End Sub

Private Sub txt0to100_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
    If IsNumeric(txt0to100(Index)) Then
        If txt0to100(Index) <= mustAddUpTo And txt0to100(Index) >= 0 Then
        Else
            Cancel = True
        End If
    Else
        Cancel = True
    End If
    
    If Cancel = True Then
        MsgBox "Value must be numeric between 0 and " & mustAddUpTo
    End If
    
End Sub

Private Sub txtMultiplier_Validate(Cancel As Boolean)
On Error Resume Next
    If txtMultiplier.Text <> "" Then
        If Not IsNumeric(txtMultiplier.Text) Then
            MsgBox "The multiplier must by numeric"
            Cancel = True
        End If
    End If
End Sub

Private Sub txtSessionID_Change()
On Error Resume Next
    setFormCaption
End Sub

Private Sub UpDown1_Change()
On Error Resume Next
    Slider1.Value = UpDown1.Value
    
End Sub

Private Sub UpDown2_Change(Index As Integer)
On Error Resume Next
    slider0to100(Index).Value = UpDown2(Index).Value

End Sub
