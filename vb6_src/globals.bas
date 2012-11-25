Attribute VB_Name = "globals"
Option Explicit
Public theTesters As testers

Public areasPath As String
Public datafilesPath As String
Public todosPath As String
Public sessionsPath As String  'submitted

Public areas As Collection

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const EM_GETLINECOUNT = &HBA

Public nextSessionID As Long
Public nextInternalBugID As Long
Public nextInternalIssueID As Long
'Public nextIssueID As Long
Public theMDIForm As Form

Public Const createNewThangSymbol As String = "<...>"

Public Sub main()

    On Error Resume Next
    
    nextSessionID = 1
    nextInternalBugID = 1
    nextInternalIssueID = 1
    
'    nextIssueID = 1
    Set theMDIForm = New MDIForm1
    theMDIForm.Show
    
    
End Sub

Public Function getNextSessionID() As Long
    getNextSessionID = nextSessionID
    nextSessionID = nextSessionID + 1
End Function
Public Function getNextInternalBugID() As Long
    getNextInternalBugID = nextInternalBugID
    nextInternalBugID = nextInternalBugID + 1
End Function
Public Function getNextinternalIssueID() As Long
    getNextinternalIssueID = nextInternalIssueID
    nextInternalIssueID = nextInternalIssueID + 1
End Function
'Public Function getNextIssueID() As Long
'    getNextIssueID = nextIssueID
'    nextIssueID = nextIssueID + 1
'End Function




Public Function loadAreas(FileName As String)

    Dim fileNum As Integer
    Dim fileline As String
    
    'load all the areas into a collection
    On Error GoTo noLoad
    
    fileNum = FreeFile
    Open FileName For Input As fileNum
    
    Set areas = Nothing
    Set areas = New Collection
    
    Do Until EOF(fileNum)
        Line Input #fileNum, fileline
        If fileline <> "" Then
            If Left$(fileline, 1) <> "#" Then
                areas.Add (fileline)
                'Debug.Print fileline
            End If
        End If
    Loop
    
    Close fileNum
    
    loadAreas = True
    Exit Function
    
noLoad:
    Close fileNum
    loadAreas = False
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'**************************************
' Name: Auto resize flexgrid column widt
'     hs
' Description:Automatically resize the c
'     olumns in any flex grid to give a nice,
'     professional appearance.
'Public Sub automatically resizes MS Flex Grid columns to match the width of the text, no matter the size of the grid or the number of columns.
'    Reads first n number of rows of data, and adjusts column size To match the widest cell of text. Will even expand columns proportionately If they aren't wide enough to fill out the entire width of the grid. Configurable constraints allow you to designate
'    1) Any flex grid To resize
'    2) Maximum column width
'    3) the maximum number of rows In depth To look For the widest cell of text.
' By: Jonathan W. Lartigue
'
' Inputs:msFG (MSFlexGrid) = The name of
'     the flex grid to resize .... MaxRowsToPa
'     rse (integer) = The maximum number of ro
'     ws (depth) of the table to scan for cell
'     width (e.g. 50) .... MaxColWidth (Intege
'     r) = The maximum width of any given cell
'     in twips (e.g. 5000)
'
' Assumes:Simply drop this public sub in
'     to your form or module and access it fro
'     m anywhere in your program to automatica
'     lly resize any flex grid.
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.8547/lngWId.1/qx/vb/scripts/ShowCode.
'     htm'for details.'**************************************
Public Sub AutosizeGridColumns(ByRef msFG As MSFlexGrid, ByVal MaxRowsToParse As Integer, ByVal MaxColWidth As Integer, aForm As Form)
    Dim I, J As Integer
    Dim txtString As String
    Dim intTempWidth, intBiggestWidth As Integer
    Dim intRows As Integer
    Const intPadding = 150
    
    On Error Resume Next
    With msFG
        For I = 0 To .Cols - 1
            ' Loops through every column
            .Col = I
            ' Set the active colunm
            intRows = .Rows
            ' Set the number of rows
            If intRows > MaxRowsToParse Then intRows = MaxRowsToParse
            ' If there are more rows of data, reset
            ' intRows to the MaxRowsToParse constant
            '
           
            intBiggestWidth = 0
            ' Reset some values to 0
            For J = 0 To intRows - 1
                ' check up to MaxRowsToParse # of rows a
                '     nd obtain
                ' the greatest width of the cell content
                '     s
               
                .Row = J
               
                txtString = .Text
                intTempWidth = aForm.TextWidth(txtString) + intPadding
                ' The intPadding constant compensates fo
                '     r text insets
                ' You can adjust this value above as des
                '     ired.
               
                If intTempWidth > intBiggestWidth Then intBiggestWidth = intTempWidth
                ' Reset intBiggestWidth to the intMaxCol
                '     Width value if necessary
            Next J
            .ColWidth(I) = intBiggestWidth
        Next I
        
        
        Dim onlydolast As Boolean
        onlydolast = True
        
        If onlydolast = True Then
            'add up all cols except the last one
            intTempWidth = 0
            For I = 0 To .Cols - 2
                intTempWidth = intTempWidth + .ColWidth(I)
                ' Add up the width of all the columns
            Next I
            .ColWidth(.Cols - 1) = msFG.Width - intTempWidth - 350
        Else
             ' Now check to see if the columns aren't
             '     as wide as the grid itself.
             ' If not, determine the difference and e
             '     xpand each column proportionately
             ' to fill the grid
             intTempWidth = 0
            
             For I = 0 To .Cols - 1
                 intTempWidth = intTempWidth + .ColWidth(I)
                 ' Add up the width of all the columns
             Next I
            
             If intTempWidth < msFG.Width Then
                 ' Compate the width of the columns to th
                 '     e width of the grid control
                 ' and if necessary expand the columns.
                 intTempWidth = Fix((msFG.Width - intTempWidth) / .Cols)
                 ' Determine the amount od width expansio
                 '     n needed by each column
                 For I = 0 To .Cols - 1
                     .ColWidth(I) = .ColWidth(I) + intTempWidth
                     ' add the necessary width to each column
                     '
                    
                 Next I
             End If
        End If
    End With
End Sub

'HOWTO: Adjust RowHeight of MSFlexGrid to Accommodate WordWrap (Q178127)
Public Sub ReSizeCellHeight(MSFlexGrid1 As MSFlexGrid, MyRow As Long, MyCol As Long, aForm As Form, text1 As TextBox, Optional maxRows As Integer = -1)

         Dim LinesOfText As Long
         Dim HeightOfLine As Long

         'Set MSFlexGrid to appropriate Cell
         MSFlexGrid1.Row = MyRow
         MSFlexGrid1.Col = MyCol

         'Set textbox width to match current width of selected cell
         text1.Width = MSFlexGrid1.ColWidth(MyCol)

         'Set font info of textbox to match FlexGrid control
         text1.Font.name = MSFlexGrid1.Font.name
         text1.Font.Size = MSFlexGrid1.Font.Size
         text1.Font.Bold = MSFlexGrid1.Font.Bold
         text1.Font.Italic = MSFlexGrid1.Font.Italic
         text1.Font.Strikethrough = MSFlexGrid1.Font.Strikethrough
         text1.Font.Underline = MSFlexGrid1.Font.Underline

         'Set font info of form to match FlexGrid control
         aForm.Font.name = MSFlexGrid1.Font.name
         aForm.Font.Size = MSFlexGrid1.Font.Size
         aForm.Font.Bold = MSFlexGrid1.Font.Bold
         aForm.Font.Italic = MSFlexGrid1.Font.Italic
         aForm.Font.Strikethrough = MSFlexGrid1.Font.Strikethrough
         aForm.Font.Underline = MSFlexGrid1.Font.Underline

         'Put the text from the selected cell into the textbox
         text1.Text = "A" 'MSFlexGrid1.Text

         'Get the height of the text in the textbox
         HeightOfLine = aForm.TextHeight(text1.Text)
         'HeightOfLine = HeightOfLine + 40   ' a little border

         text1.Text = MSFlexGrid1.Text
         'Call API to determine how many lines of text are in text box
         LinesOfText = SendMessage(text1.hwnd, EM_GETLINECOUNT, 0&, 0&)

         'Check to see if row is not tall enough
         'If MSFlexGrid1.RowHeight(MyRow) < (LinesOfText * HeightOfLine) Then
            'Adjust the RowHeight based on the number of lines in textbox

        If maxRows = -1 Then
            MSFlexGrid1.RowHeight(MyRow) = (LinesOfText * HeightOfLine) + 40
        Else
            If LinesOfText > maxRows Then
                MSFlexGrid1.RowHeight(MyRow) = HeightOfLine * maxRows  '((maxRows * HeightOfLine) / 2) + 40
            Else
                MSFlexGrid1.RowHeight(MyRow) = (LinesOfText * HeightOfLine) + 40
            End If
        End If

         'End If

      End Sub

