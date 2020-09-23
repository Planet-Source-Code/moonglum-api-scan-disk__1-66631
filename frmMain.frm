VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ScanDisk"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   555
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   3660
      Width           =   945
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Scan Drive"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   8
      Top             =   3810
      Width           =   2175
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Check Only If Dirty"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   3525
      Width           =   2175
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Verbose"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   3225
      Width           =   1455
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Correct Errors"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   2940
      Width           =   1455
   End
   Begin VB.ComboBox cboDrives 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4140
      Width           =   945
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   2835
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   90
      Width           =   4485
   End
   Begin MSComctlLib.ProgressBar pbrScan 
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   4500
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Check Disk"
      Enabled         =   0   'False
      Height          =   555
      Index           =   0
      Left            =   3600
      TabIndex        =   0
      Top             =   3060
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Drive :"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   4200
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API version of check disk.
'Note :  the chkdsk is limited in it's ability to fix
'errors on a drive that it cannot open exclusively.

'Based upon Mark Russinovich's C++ code,
'which is included in this project

'Mark Russinovich
'Systems Internals
'http://www.sysinternals.com

'I use an implements due to the use of the addressof operator by the Chkdsk API
'Private Declare Sub Chkdsk Lib "fmifs.dll" (ByVal DriveRoot As String, _                   'The drive root to be scanned
'                                            ByVal Format As String, _                      'The format of the drive (NTFS/FAT32/FAT)
'                                            ByVal CorrectErrors As Long, _                 'Fix the errors on disk ?
'                                            ByVal Verbose As Long, _                       On FAT/FAT32: Displays the full path and name of every file on the disk.  On NTFS: Displays cleanup messages if any.
'                                            ByVal CheckOnlyIfDirty As Long, _              Check only if the diry bit has been set
'                                            ByVal ScanDrive As Long, _
'                                            ByVal Unused2 As Long, _
'                                            ByVal Unused3 As Long, _
'                                            ByVal Callback As Long)

'Chkdsk does not set the "Dirty Bit" on an in-use volume in order to check the volume at the next boot.
'Instead, it sets a registry entry to tell Autochk to run against that volume.
'The "Dirty Bit" is set by the file system itself only if it detects a problem.
                                            
                                            
                                            
Private Enum eButtons                                   'Button array
    eCheck
    eCancel
End Enum

Private Enum eOptions                                   'Checkbox array
    eCorrectErrors
    eVerbose
    eDirty
    eScanDrive
End Enum

Implements IChkDskSink                                  'Receive addressof events
Private m_bCancel As Boolean                            'Cancel scan
Private Sub Form_Load()
    Dim sDrives() As String                             'Array for holding available drives
    Dim lCount As Long                                  'Count of drives
    Dim idx As Integer                                  'Loop variable
    
    If EnumAllDrives(sDrives(), lCount) Then            'Enumerate all drives
        For idx = LBound(sDrives) To UBound(sDrives)    'loop thru all drives
            Select Case GetDriveType(sDrives(idx))      'Check the drive type
                Case eDriveFixed, eDriveRemoveable      'Only display fixed and removeable drives
                    If DriveIsReady(sDrives(idx)) Then  'Is the drive ready ?
                        cboDrives.AddItem sDrives(idx)  'Add it to the combobox
                    End If  'If DriveIsReady(sDrives(idx)) Then
                    
                Case Else   'Select Case GetDriveType(sDrives(idx))
            End Select  'Select Case GetDriveType(sDrives(idx))
        Next idx
    End If  'If EnumAllDrives(sDrives(), lCount) Then
End Sub
Private Sub cboDrives_Click()
    'Enable scan button once a drive has been selected
    cmdButton(eButtons.eCheck).Enabled = Not (cboDrives.ListIndex = -1)
End Sub
Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case eButtons.eCheck       'Check Disk
            m_bCancel = False                           'Set/Reset cancel variable
            Screen.MousePointer = vbHourglass           'Show that we are busy
            txtOutput.Text = vbNullString               'Reset the output window
            pbrScan.Value = 0                           'Reset the progress bar
            cmdButton(eButtons.eCancel).Enabled = True  'Enable the cancel button
            cmdButton(eButtons.eCheck).Enabled = False  'Disable the scan button
            Call CheckDisk(cboDrives.Text, _
                            chkOptions(eOptions.eCorrectErrors).Value, _
                            chkOptions(eOptions.eVerbose).Value, _
                            chkOptions(eOptions.eDirty).Value, _
                            chkOptions(eOptions.eScanDrive).Value, _
                            Me)                         'Call our checkdisk function
            
        Case eButtons.eCancel       'Cancel
            m_bCancel = True                            'Set our module level cancel variable
    
    End Select  'Select Case Index
End Sub
Private Sub IChkDskSink_Progress(ByVal lNewValue As Long, _
                                 bCancel As Boolean)
    pbrScan.Value = lNewValue                           'Update the progress bar
    bCancel = m_bCancel                                 'Set our module level cancel variable
End Sub
Private Sub IChkDskSink_Status(ByVal sValue As String, bCancel As Boolean)
    txtOutput.Text = txtOutput.Text & sValue            'Append our status to the output window
    txtOutput.SelStart = Len(txtOutput.Text)            'Move the focus so text scrolls
    DoEvents                                            'Give the output enough time
    bCancel = m_bCancel                                 'Set our module level cancel variable
End Sub
Private Sub IChkDskSink_Done(bError As Boolean)
    Screen.MousePointer = vbDefault                     'Reset the busy state
    cmdButton(eButtons.eCancel).Enabled = False         'Disable our cancel button
    cmdButton(eButtons.eCheck).Enabled = True           'Enable check disk button
    
    pbrScan.Value = 0                                   'Reset the progress bar
    If m_bCancel Then
        MsgBox "Cancelled !"
    Else
        MsgBox "Done !"
    End If
End Sub

