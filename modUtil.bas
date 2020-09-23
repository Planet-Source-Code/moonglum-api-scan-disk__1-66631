Attribute VB_Name = "modUtil"
Option Explicit

'Most of these file system routines are based on Randy Birch's stuff
'http://vbnet.mvps.org/

Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5  'can be a CD or a DVD
Private Const DRIVE_RAMDISK = 6

Public Enum eDriveType
    eDriveUnknown = 0
    eDriveRemoveable = DRIVE_REMOVABLE
    eDriveFixed = DRIVE_FIXED
    eDriveRemote = DRIVE_REMOTE
    eDriveCDROM = DRIVE_CDROM
    eDriveRAMDisk = DRIVE_RAMDISK
End Enum

Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Function EnumAllDrives(ByRef sDrives() As String, _
                              ByRef lCount As Long)                             'Enumerate all available drives
    Dim sBuff As String                                                         'Buffer
    Dim sTemp As String
    
    sBuff = Space$((26 * 4) + 1)                                                'Initialize buffer for all letters
    
    If GetLogicalDriveStrings(Len(sBuff), sBuff) Then
        sTemp = Replace$(sBuff, Chr$(0), Chr$(32))
        sDrives = Split(Trim$(sTemp), Chr$(32))                                 'Split up returned string,
        lCount = UBound(sDrives)                                                'into an array
        EnumAllDrives = True
    End If
End Function
Public Function GetDriveType(ByVal sDriveLetter As String) As eDriveType
    If Not (Right$(sDriveLetter, 2) = ":\") Then                                'API function requires the :\
        sDriveLetter = sDriveLetter & ":\"
    End If
    
    Select Case GetDriveTypeA(sDriveLetter)
        Case 0, 1:                      GetDriveType = eDriveUnknown
        Case DRIVE_REMOVABLE:           GetDriveType = eDriveRemoveable
        Case DRIVE_FIXED:               GetDriveType = eDriveFixed
        Case DRIVE_REMOTE:              GetDriveType = eDriveRemote
        Case DRIVE_CDROM:               GetDriveType = eDriveCDROM
        Case DRIVE_RAMDISK:             GetDriveType = eDriveRAMDisk
    End Select
End Function
Public Function DriveIsReady(ByVal sDrive As String) As Boolean
    Dim sVolumeName As String
    Dim dwVolumeSize  As Long

    sVolumeName = Space$(32)
    dwVolumeSize = Len(sVolumeName)
    
    DriveIsReady = GetVolumeInformation(sDrive, _
                                         sVolumeName, _
                                         dwVolumeSize, _
                                         0&, 0&, 0&, _
                                         vbNullString, _
                                         0&)

End Function

