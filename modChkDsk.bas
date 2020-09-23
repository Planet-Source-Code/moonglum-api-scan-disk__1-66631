Attribute VB_Name = "modChkDsk"
Option Explicit

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const INVALID_HANDLE_VALUE = -1

Public Enum CALLBACKCOMMAND
    Progress
    DONEWITHSTRUCTURE
    UNKNOWN2
    UNKNOWN3
    UNKNOWN4
    UNKNOWN5
    INSUFFICIENTRIGHTS
    UNKNOWN7
    UNKNOWN8
    UNKNOWN9
    UNKNOWNA
    Done
    UNKNOWNC
    UNKNOWND
    Output
    STRUCTUREPROGRESS
End Enum

Private Declare Sub Chkdsk Lib "fmifs.dll" (ByVal DriveRoot As String, _
                                            ByVal Format As String, _
                                            ByVal CorrectErrors As Long, _
                                            ByVal Verbose As Long, _
                                            ByVal CheckOnlyIfDirty As Long, _
                                            ByVal ScanDrive As Long, _
                                            ByVal Unused2 As Long, _
                                            ByVal Unused3 As Long, _
                                            ByVal Callback As Long)
            
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
      
Private m_cSink As IChkDskSink
Public Sub CheckDisk(ByVal sDrive As String, _
                     ByVal bCorrectErrors As Boolean, _
                     ByVal bVerbose As Boolean, _
                     ByVal bCheckOnlyIfDirty As Boolean, _
                     ByVal bScanDrive As Boolean, _
                     ByRef cSink As IChkDskSink)    'Our API Checkdisk
                               
    Dim sFileSystem As String
    Dim hVolume As Long
    Dim hLib As Long
    Dim hProcess As Long
    
    If Not (m_cSink Is Nothing) Then Exit Sub                                   '>>>>>
    
    hLib = LoadLibrary("fmifs.dll")                                             'See if the library is available
    If Not (hLib = 0) Then
        hProcess = GetProcAddress(hLib, "Chkdsk")                               'Get the address of the exported function
        If Not (hProcess = 0) Then                                              'All a go ?
            Set m_cSink = cSink                                                 'Instantiate our implements class
            
            sFileSystem = Space$(32)                                            'Initialize a buffer
            If GetVolumeInformation(sDrive, _
                                    vbNullString, _
                                    0&, _
                                    0&, _
                                    0&, _
                                    0&, _
                                    sFileSystem, _
                                    Len(sFileSystem)) Then                      'Get the file system for the drive, the chkdsk function requires this
                                      
                sFileSystem = TrimNull(sFileSystem)                             'Truncate our buffer
                If bCorrectErrors Then                                          'Are we correcting errors ?
                    hVolume = CreateFile("\\.\" & Left$(sDrive, 1) & ":", _
                                         GENERIC_WRITE, _
                                         0, _
                                         ByVal 0&, _
                                         OPEN_EXISTING, _
                                         0, _
                                         0)                                     'Try to open drive exclusively
                                         
                    If (hVolume = INVALID_HANDLE_VALUE) Then                    'If we weren't able to,
                        bCorrectErrors = False                                  'skip correcting errors
                    Else    'If (hVolume = INVALID_HANDLE_VALUE) Then
                        CloseHandle hVolume                                     'Clean up our handle
                    End If  'If (hVolume = INVALID_HANDLE_VALUE) Then
                End If  'If bCorrectErrors Then
                
                sDrive = StrConv(sDrive, vbUnicode)                             'Get the unicode for the drive
                sFileSystem = StrConv(sFileSystem, vbUnicode)                   'Get the unicode for the file system
                
                Call Chkdsk(sDrive, _
                            sFileSystem, _
                            bCorrectErrors, _
                            bVerbose, _
                            bCheckOnlyIfDirty, _
                            bScanDrive, _
                            0&, _
                            0&, _
                            AddressOf CheckDiskCallback)                        'Here we go !
                            
            End If  'If GetVolumeInformation(sDrive,...
        Else    'If Not (hProcess = 0) Then
        
        End If  'If Not (hProcess = 0) Then
        
        FreeLibrary hLib                                                        'Free up our handle to the library
    Else    'If Not (hLib = 0) Then
        
    End If  'If Not (hLib = 0) Then
    
    Set m_cSink = Nothing                                                       'Clean up our class object
End Sub
Public Function CheckDiskCallback(ByVal Command As CALLBACKCOMMAND, _
                                  ByVal Modifier As Long, _
                                  ByVal Argument As Long) As Long               'Call back for our checkdisk

    Dim percent As Long
    Dim Status As Long
    Dim OutStrPtr As Long, OutLines As Long
    Dim OutLen As Long, Output As String
    Dim Progress As Long
    Dim bCancel As Boolean
    
    DoEvents                                                                    'Free up some CPU
    
    Select Case Command
        Case CALLBACKCOMMAND.Progress                                           'PDWORD
            CopyMemory ByVal VarPtr(Progress), ByVal Argument, 4
            m_cSink.Progress Progress, bCancel                                  'Raise the progress and cancel state
            
        Case CALLBACKCOMMAND.Done
            CopyMemory ByVal VarPtr(Status), ByVal Argument, 1                  'PBOOLEAN
            m_cSink.Done (Status = 0)                                           'Raise done and error state
            
        Case CALLBACKCOMMAND.Output                                             'PTEXTOUTPUT
            If IsBadReadPtr(ByVal Argument, 8) = 0 Then                         'Do we have read access to the memory range ?
                CopyMemory ByVal VarPtr(OutLines), ByVal Argument, 4
                CopyMemory ByVal VarPtr(OutStrPtr), ByVal Argument + 4, 4
                If IsBadStringPtr(ByVal OutStrPtr, 255) = 0 Then                'Do we have read access to a string pointer memory range ?
                    OutLen = lstrlen(ByVal OutStrPtr)
                    Output = Space$(OutLen * 2)
                    CopyMemory ByVal StrPtr(Output), ByVal OutStrPtr, OutLen * 2
                    Output = StrConv(Output, vbUnicode)
                    OemToChar Output, Output                                    'translate any OEM-defined character set
                    Output = Mid$(Output, 1, InStr(Output, vbNullChar) - 1)
                    
                    m_cSink.Status Output, bCancel                              'Pass on the text
                End If  'If IsBadStringPtr(ByVal OutStrPtr...
            End If  'If IsBadReadPtr(ByVal Argument...
            
        Case Else
            Debug.Print "Unknown"
            
    End Select  'Select Case Command

    If bCancel Then                                                             'Return if cancelled
        CheckDiskCallback = 0
    Else
        CheckDiskCallback = 1
    End If
End Function
Private Function TrimNull(startstr As String) As String
   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))                       'Trim nulls off a string
End Function

