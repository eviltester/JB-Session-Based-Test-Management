VERSION 5.00
Begin VB.Form frmTextDetails 
   Caption         =   "LinesOfText"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Apply"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtDetails 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmTextDetails.frx":0000
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblID 
      Caption         =   "ID"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmTextDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public localSessionID As Long
Public uniqueID As Long

Public Enum entityType
    anIssue
    aBug
End Enum

Public thisEntity As entityType


Private Sub btnCancel_Click()

    Unload Me
    
End Sub

Private Sub btnOK_Click()

    Dim parentSession As Form
On Error Resume Next
    If txtDetails.Text = "" Then
        MsgBox "The description cannot be blank"
        Exit Sub
    End If
    
    For Each parentSession In Forms
        If TypeName(parentSession) = "frmSessionEdit" Then
            If parentSession.sessionID = localSessionID Then
                Exit For
            End If
        End If
    Next
    
    If thisEntity = aBug Then
        parentSession.addEditBug uniqueID, txtID.Text, txtDetails.Text
    ElseIf thisEntity = anIssue Then
        parentSession.addEditIssue uniqueID, txtID.Text, txtDetails.Text
    End If

    
        
    If uniqueID = -1 Then
        If thisEntity <> anIssue Then
            txtID.Text = ""
        End If
        txtDetails.Text = ""
'    Else
'        Unload Me
    End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
    btnOK.Top = Me.ScaleHeight - btnOK.Height - 50
    btnCancel.Top = btnOK.Top
    txtDetails.Height = btnOK.Top - 50 - txtDetails.Top
    txtDetails.Width = Me.ScaleWidth - txtDetails.Left
End Sub
