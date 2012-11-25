VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: Skip
'*****************************************************************************************
'* Class       : CIniFile
'* Description : Class  for working with initialization (*.ini) file.
'* Notes       : Use this class to give your application the capability to store and
'*               retrieve its settings to and from an initialization file.
'*               For advanced use, please see the Source+ Library online documentation
'*               in the Source+ Explorer (the "Documentation" tab).
'*****************************************************************************************

Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 17680 + vbObjectError
Public Enum EErrIniFile
    eErrIniFile_InvalidFileName = E_ERR_BASE + 1
    eErrIniFile_FileNotFound
    eErrIniFile_NotInitialized
    eErrIniFile_InvalidSection
    eErrIniFile_InvalidKey
    eErrIniFile_ComponentFailure
End Enum
Private Const S_ERR_InvalidFileName = "Invalid file name"
Private Const S_ERR_FileNotFound = "File not found"
Private Const S_ERR_NotInitialized = "Object not properly initialized"
Private Const S_ERR_InvalidSection = "Invalid section"
Private Const S_ERR_InvalidKey = "Invalid key"
Private Const S_ERR_ComponentFailure = "CIniFile component failure"

' Private class constants
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_LENGTH_EX = 512

' Private class type definitions
Private Type FILETIME
    dwLowDateTime    As Long
    dwHighDateTime   As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_LENGTH_EX
    cAlternate       As String * 14
End Type

' Private class API function declarations
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

' Private variables for internal class use
Private m_sDefault   As String

' Private variables to hold property values
Private m_sFile      As String


'*****************************************************************************************
'* Sub         : Init
'* Notes       : Use this routine for basic object initialization.
'*****************************************************************************************
Public Sub Init(FileName As String)
    On Error GoTo hComponentFailure
    
    If Len(FileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidFileName, App.EXEName & ".CIniFile", S_ERR_InvalidFileName
    End If
    
    m_sFile = FileName

    Exit Sub

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Property    : FileName
'* Notes       : Returns or sets the name of the initialization (*.ini) file.
'*****************************************************************************************
Public Property Let FileName(Value As String)
    On Error GoTo hComponentFailure
    
    If Len(Value) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidFileName, App.EXEName & ".CIniFile", S_ERR_InvalidFileName
    End If
    
    m_sFile = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Property

Public Property Get FileName() As String
    On Error GoTo hComponentFailure
    
    FileName = m_sFile

    Exit Property

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Function    : DeleteKey
'* Notes       : Deletes a key setting from a specified section of the initialization file.
'*               Returns true if the key was deleted and false otherwise.
'*****************************************************************************************
Public Function DeleteKey(ByVal Section As String, ByRef Key As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim iRet As Long
    
    If Len(m_sFile) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_NotInitialized, App.EXEName & ".CIniFile", S_ERR_NotInitialized
    End If
    
    If Not FileExists(m_sFile) Then
        On Error GoTo 0
        Err.Raise eErrIniFile_FileNotFound, App.EXEName & ".CIniFile", S_ERR_FileNotFound
    End If
    
    If Len(Section) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidSection, App.EXEName & ".CIniFile", S_ERR_InvalidSection
    End If
    
    If Len(Key) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidKey, App.EXEName & ".CIniFile", S_ERR_InvalidKey
    End If
    
    iRet = WritePrivateProfileString(Section, Key, 0&, m_sFile)
    WritePrivateProfileString vbNullString, vbNullString, vbNullString, m_sFile
    DeleteKey = (iRet <> 0)

    Exit Function

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : DeleteSection
'* Notes       : Deletes a section of the initialization file.
'*               Returns true if the section was deleted and false otherwise.
'*****************************************************************************************
Public Function DeleteSection(ByVal Section As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim iRet As Long
    
    If Len(m_sFile) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_NotInitialized, App.EXEName & ".CIniFile", S_ERR_NotInitialized
    End If
    
    If Not FileExists(m_sFile) Then
        On Error GoTo 0
        Err.Raise eErrIniFile_FileNotFound, App.EXEName & ".CIniFile", S_ERR_FileNotFound
    End If
    
    If Len(Section) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidSection, App.EXEName & ".CIniFile", S_ERR_InvalidSection
    End If
    
    iRet = WritePrivateProfileString(Section, 0&, 0&, m_sFile)
    DeleteSection = (iRet <> 0)
    WritePrivateProfileString vbNullString, vbNullString, vbNullString, m_sFile

    Exit Function

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : EnumerateAllKeys
'* Notes       : Retrieves all the keys (in a an array of strings) belonging to a section
'*               of an initialization file.
'*               Returns the number of retrieved keys.
'*****************************************************************************************
Public Function EnumerateAllKeys(Section As String, ByRef Key() As String) As Long
    On Error GoTo hComponentFailure
    
    Dim lApi_Ret     As Long
    Dim lCount       As Long
    Dim lNextPos     As Long
    Dim lPos         As Long
    Dim lSize        As Long
    Dim buffer       As String
    Dim sCur         As String
    Dim sSections    As String
    
    If Len(m_sFile) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_NotInitialized, App.EXEName & ".CIniFile", S_ERR_NotInitialized
    End If
    
    If Not FileExists(m_sFile) Then
        On Error GoTo 0
        Err.Raise eErrIniFile_FileNotFound, App.EXEName & ".CIniFile", S_ERR_FileNotFound
    End If
    
    If Len(Section) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidSection, App.EXEName & ".CIniFile", S_ERR_InvalidSection
    End If
    
    lCount = 0
    Erase Key
    
    If (Len(Section) > 0) Then
        
        buffer = Space$(8192)
        lSize = Len(buffer)
        lApi_Ret = GetPrivateProfileString(Section, 0&, m_sDefault, buffer, lSize, m_sFile)
        If (lSize > 0) Then
            sSections = Left$(buffer, lApi_Ret)
        Else
            sSections = ""
        End If

        lPos = 1
        lNextPos = InStr(lPos, sSections, Chr$(0))
        
        Do While lNextPos <> 0
            sCur = Mid$(sSections, lPos, (lNextPos - lPos))
            
            If (sCur <> Chr$(0)) Then
                lCount = lCount + 1
                
                ReDim Preserve Key(1 To lCount) As String
                Key(lCount) = Mid$(sSections, lPos, (lNextPos - lPos))
                
                lPos = lNextPos + 1
                lNextPos = InStr(lPos, sSections, Chr$(0))
            End If
        
        Loop
    
    End If
    
    EnumerateAllKeys = lCount

    Exit Function

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Sub         : EnumerateAllSections
'* Notes       : Retrieves all the sections belonging to an initialization file.
'*               Returns the number of retrieved sections.
'*****************************************************************************************
Public Function EnumerateAllSections(ByRef Sections() As String) As Long
    On Error GoTo hComponentFailure
    
    Dim lApi_Ret    As Long
    Dim lCount      As Long
    Dim lNextPos    As Long
    Dim lPos        As Long
    Dim lSize       As Long
    Dim sCur        As String
    Dim buffer      As String
    Dim sSections   As String
    
    If Len(m_sFile) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_NotInitialized, App.EXEName & ".CIniFile", S_ERR_NotInitialized
    End If
    
    If Not FileExists(m_sFile) Then
        On Error GoTo 0
        Err.Raise eErrIniFile_FileNotFound, App.EXEName & ".CIniFile", S_ERR_FileNotFound
    End If
    
    lCount = 0
    Erase Sections
    
    If (Len(m_sFile) > 0) Then
        
        buffer = Space$(8192)
        lSize = Len(buffer)
        lApi_Ret = GetPrivateProfileString(0&, 0&, m_sDefault, buffer, lSize, m_sFile)
        If (lSize > 0) Then
            sSections = Left$(buffer, lApi_Ret)
        Else
            sSections = ""
        End If

        lPos = 1
        lNextPos = InStr(lPos, sSections, Chr$(0))
        
        Do While lNextPos <> 0
            
            If (lNextPos <> lPos) Then
                sCur = Mid$(sSections, lPos, (lNextPos - lPos))
                lCount = lCount + 1
                
                ReDim Preserve Sections(1 To lCount) As String
                Sections(lCount) = sCur
            End If
            
            lPos = lNextPos + 1
            lNextPos = InStr(lPos, sSections, Chr$(0))
        
        Loop
    
    End If
    
    EnumerateAllSections = lCount

    Exit Function

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : GetNumericValue
'* Notes       : Returns an integer value containing the key setting stored in the
'*               specified section of an initialization file.
'*****************************************************************************************
Public Function GetNumericValue(Section As String, Key As String) As Integer
    On Error GoTo hComponentFailure
    
    If Len(m_sFile) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_NotInitialized, App.EXEName & ".CIniFile", S_ERR_NotInitialized
    End If
    
    If Not FileExists(m_sFile) Then
        On Error GoTo 0
        Err.Raise eErrIniFile_FileNotFound, App.EXEName & ".CIniFile", S_ERR_FileNotFound
    End If
    
    GetNumericValue = GetPrivateProfileInt(Section, Key, -1, m_sFile)

    Exit Function

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : GetStringValue
'* Notes       : Returns a string value containing the key setting stored in the specified
'*               section of an initialization file.
'*****************************************************************************************
Public Function GetStringValue(Section As String, Key As String) As String
    On Error GoTo hComponentFailure
    
    Dim lResult  As Long
    Dim sText    As String
    
    If Len(m_sFile) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_NotInitialized, App.EXEName & ".CIniFile", S_ERR_NotInitialized
    End If
    
    If Not FileExists(m_sFile) Then
        On Error GoTo 0
        Err.Raise eErrIniFile_FileNotFound, App.EXEName & ".CIniFile", S_ERR_FileNotFound
    End If
    
    sText = String$(255, 0)
    lResult = GetPrivateProfileString(Section, Key, "", sText, Len(sText), m_sFile)
    
    If lResult = 0 Then
        GetStringValue = ""
    Else
        GetStringValue = Left$(sText, InStr(sText, Chr$(0)) - 1)
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Sub         : SaveNumericValue
'* Notes       : Saves a key with a numeric value in the specified section of an
'*               initialization file.
'*****************************************************************************************
Public Sub SaveNumericValue(Section As String, Key As String, iData As Integer)
    On Error GoTo hComponentFailure
    
    Dim sData As String
    
    If Len(m_sFile) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_NotInitialized, App.EXEName & ".CIniFile", S_ERR_NotInitialized
    End If
    
    If Len(Section) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidSection, App.EXEName & ".CIniFile", S_ERR_InvalidSection
    End If
    
    If Len(Key) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidKey, App.EXEName & ".CIniFile", S_ERR_InvalidKey
    End If

    sData = iData
    WritePrivateProfileString Section, Key, sData, m_sFile
    WritePrivateProfileString vbNullString, vbNullString, vbNullString, m_sFile
    
    Exit Sub

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : SaveStringValue
'* Notes       : Saves a key containing a string value in the specified section of an
'*               initialization file.
'*****************************************************************************************
Public Sub SaveStringValue(Section As String, Key As String, sText As String)
    On Error GoTo hComponentFailure
    
    If Len(m_sFile) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_NotInitialized, App.EXEName & ".CIniFile", S_ERR_NotInitialized
    End If
    
    If Len(Section) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidSection, App.EXEName & ".CIniFile", S_ERR_InvalidSection
    End If
    
    If Len(Key) = 0 Then
        On Error GoTo 0
        Err.Raise eErrIniFile_InvalidKey, App.EXEName & ".CIniFile", S_ERR_InvalidKey
    End If
    
    WritePrivateProfileString Section, Key, sText, m_sFile
    WritePrivateProfileString vbNullString, vbNullString, vbNullString, m_sFile

    Exit Sub

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Function    : FileExists
'* Notes       : API-based routine to check for a file's existence.
'*               Returns True if the specified file exists, False otherwise.
'*****************************************************************************************
Private Function FileExists(FileName As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim wfd      As WIN32_FIND_DATA
    Dim hFile    As Long
    
    FileExists = False
    
    hFile = FindFirstFile(FileName, wfd)
    If hFile <> INVALID_HANDLE_VALUE Then
        FileExists = True
        
        If FindClose(hFile) = 0 Then
            On Error GoTo 0
            Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
        End If
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Sub         : Class_Initialize
'* Notes       : Class data space initialization.
'*****************************************************************************************
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure
    
    m_sFile = App.Path & "\MyFileSettings.ini"

    Exit Sub

hComponentFailure:
    Err.Raise eErrIniFile_ComponentFailure, App.EXEName & ".CIniFile", S_ERR_ComponentFailure
End Sub
