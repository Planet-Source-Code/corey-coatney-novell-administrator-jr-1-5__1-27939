VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServerStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Statistics"
   ClientHeight    =   6165
   ClientLeft      =   2910
   ClientTop       =   2430
   ClientWidth     =   9555
   Icon            =   "frmServerStatistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9555
   Begin VB.CommandButton BtnClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Frame Frame5 
      Caption         =   "Administration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5880
      TabIndex        =   12
      Top             =   3120
      Width           =   3495
      Begin VB.CommandButton BtnServerAdmin 
         Caption         =   "&Server Admin"
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Send Broadcast messages, drop log in connections, or view who is logged into the server."
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Your Connection Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   5415
      Begin MSComctlLib.ListView lstUser 
         Height          =   2415
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3705
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "General Server Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   5415
      Begin MSComctlLib.ListView lstServer 
         Height          =   2415
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   4939
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Critical Server Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2895
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton BtnUtilization 
         Caption         =   "Server Utilization..."
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label LblServerName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "CPU:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LblCPU 
         Caption         =   "N/A"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "frmServerStatistics.frx":08CA
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         Caption         =   "Uptime:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label LblUptime 
         Caption         =   "N/A"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Other:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label LblOther 
         Caption         =   "N/A"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   1800
         Width           =   2295
      End
   End
   Begin VB.Timer tmrAnimate1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7200
      Top             =   3360
   End
   Begin VB.Timer tmrAnimate2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7680
      Top             =   3360
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8160
      Top             =   3960
   End
   Begin VB.Timer tmrDisconnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8640
      Top             =   3960
   End
   Begin MSComctlLib.ProgressBar prStatus 
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   3120
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   7
      Left            =   2880
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   6
      Left            =   2880
      Top             =   5400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   5
      Left            =   2880
      Top             =   5640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   4
      Left            =   2880
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   5
      Left            =   2400
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   4
      Left            =   2400
      Top             =   5400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   3
      Left            =   2400
      Top             =   5640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   7
      Left            =   2640
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   6
      Left            =   2640
      Top             =   5400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   5
      Left            =   2640
      Top             =   5640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   4
      Left            =   2640
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2160
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgConnected 
      Height          =   225
      Left            =   8760
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   0
      Left            =   8520
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   1
      Left            =   8520
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   2
      Left            =   8520
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   3
      Left            =   8520
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   0
      Left            =   8040
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   1
      Left            =   8040
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   2
      Left            =   8040
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   0
      Left            =   8400
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   1
      Left            =   8280
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   2
      Left            =   8280
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   3
      Left            =   8280
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEmpty 
      Height          =   255
      Left            =   7800
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmServerStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NW_MAX_TREE_NAME_LEN = 33
Const NW_MAX_SERVER_NAME_LEN = 49

Private Type VERSION_INFO
            ServerName(47) As Byte
            fileServiceVersion As Byte
            fileServiceSubVersion As Byte
            maximumServiceConnections As Integer
            connectionsInUse As Integer
            maxNumberVolumes As Integer
            revision As Byte
            SFTLevel As Byte
            TTSLevel As Byte
            maxConnectionsEverUsed As Integer
            accountVersion As Byte
            VAPVersion As Byte
            queueVersion As Byte
            printVersion As Byte
            virtualConsoleVersion As Byte
            restrictionLevel As Byte
            internetBridge As Byte
            reserved(59) As Byte
End Type

Private Type SERVER_AND_VCONSOLE_INFO
            currentServerTime As Long
            vconsoleVersion As Byte
            vconsoleRevision As Byte
End Type

Private Type FSE_SERVER_INFO
            replyCanceledCount As Long
            writeHeldOffCount As Long
            writeHeldOffWithDupRequest As Long
            invalidRequestTypeCount As Long
            beingAbortedCount As Long
            alreadyDoingReallocCount As Long
            deAllocInvalidSlotCount As Long
            deAllocBeingProcessedCount As Long
            deAllocForgedPacketCount As Long
            deAllocStillTransmittingCount As Long
            startStationErrorCount As Long
            invalidSlotCount As Long
            beingProcessedCount As Long
            forgedPacketCount As Long
            stillTransmittingCount As Long
            reExecuteRequestCount As Long
            invalidSequenceNumCount As Long
            duplicateIsBeingSentAlreadyCnt As Long
            sentPositiveAcknowledgeCount As Long
            sentDuplicateReplyCount As Long
            noMemForStationCtrlCount As Long
            noAvailableConnsCount As Long
            reallocSlotCount As Long
            reallocSlotCameTooSoonCount As Long
End Type

Private Type FILE_SERVER_COUNTERS
            tooManyHops As Integer
            unknownNetwork As Integer
            noSpaceForService As Integer
            noReceiveBuffers As Integer
            notMyNetwork As Integer
            netBIOSProgatedCount As Long
            totalPacketsServiced As Long
            totalPacketsRouted As Long
End Type


Private Type NWFSE_FILE_SERVER_INFO
            serverTimeAndVConsoleInfo As SERVER_AND_VCONSOLE_INFO
            reserved As Long
            NCPStationsInUseCount As Long
            NCPPeakStationsInUseCount As Long
            numOfNCPRequests As Long
            serverUtilization As Long
            ServerInfo As FSE_SERVER_INFO
            fileServerCounters As FILE_SERVER_COUNTERS
End Type

Private Type tagNWCCTranAddr
            type1 As Long
            len1 As Long
            buffer As Long
End Type

Private Type tagNWCCVersion
            major As Long
            minor As Long
            revision As Long
End Type

Private Type tagNWCCConnInfo
            authenticationState As Long
            broadcastState As Long
            connRef As Long
            TreeName(NW_MAX_TREE_NAME_LEN - 1) As Byte
            connNum As Long
            userID As Long
            ServerName(NW_MAX_SERVER_NAME_LEN - 1) As Byte
            NDSState As Long
            maxPacketSize As Long
            licenseState As Long
            distance As Long
            serverVersion As tagNWCCVersion
            tranAddr As tagNWCCTranAddr
End Type

Const NWCC_NAME_FORMAT_BIND = &H2
Const NWCC_OPEN_UNLICENSED = &H2

Private Declare Function NWCallsInit Lib "calwin32" (reserved1 As Byte, reserved2 As Byte) As Long
Private Declare Function NWCCCloseConn Lib "clxwin32" (ByVal connHandle As Long) As Long
Private Declare Function NWCCOpenConnByName Lib "clxwin32" (ByVal startConnHandle As Long, ByVal name1 As String, ByVal nameFormat As Long, ByVal openState As Long, ByVal tranType As Long, pConnHandle As Long) As Long
Private Declare Function NWGetFileServerInformation Lib "calwin32" (ByVal conn As Long, ByVal ServerName As String, majorVer As Byte, minVer As Byte, rev As Byte, maxConns As Integer, maxConnsUsed As Integer, ConnsInUse As Integer, numVolumes As Integer, SFTLevel As Byte, TTSLevel As Byte) As Long
Private Declare Function NWGetFileServerVersionInfo Lib "calwin32" (ByVal conn As Long, versBuffer As VERSION_INFO) As Long
Private Declare Function NWGetVolumeName Lib "calwin32" (ByVal conn As Long, ByVal VolNum As Integer, ByVal volName As String) As Long
Private Declare Function NWGetConnectionNumber Lib "calwin32" (ByVal connHandle As Long, connNumber As Integer) As Long
Private Declare Function NWGetConnectionInformation Lib "calwin32" (ByVal connHandle As Long, ByVal connNumber As Integer, ByVal pObjName As String, pObjType As Integer, pObjID As Long, pLoginTime As Byte) As Long
Private Declare Function NWCheckConsolePrivileges Lib "calwin32" (ByVal conn As Long) As Long
Private Declare Function NWGetNetworkSerialNumber Lib "calwin32" (ByVal conn As Long, serialNum As Long, appNum As Integer) As Long
Private Declare Function NWLongSwap Lib "calwin32" (ByVal swapLong As Long) As Long
Private Declare Function NWGetFileServerInfo Lib "calwin32" (ByVal conn As Long, fseFileServerInfo As NWFSE_FILE_SERVER_INFO) As Long
Private Declare Function NWCCGetAllConnInfo Lib "clxwin32" (ByVal connHandle As Long, ByVal connInfoVersion As Long, connInfoBuffer As tagNWCCConnInfo) As Long

Public ccode As Long
Public connHandle
Public gServerName As String
Private Sub BtnClose_Click()
'close all connections

'close the form
Unload Me
End Sub

Private Sub BtnServerAdmin_Click()
Dim msg, y

On Error Resume Next   ' Defer error handling.

frmServerAdmin.Show

frmServerAdmin.Caption = "Server Administration for Server: " & ServerName

If Err.Number <> 0 Then
   msg = "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & Err.Description
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
End If

End Sub



Private Sub BtnUtilization_Click()
Dim frmX As frmUtilization
Set frmX = New frmUtilization
ServerName = LblServerName.Caption
frmX.Caption = LblServerName.Caption & " Utilization Statistics"
frmX.Show
    
End Sub

Private Sub Form_Load()
Dim msg

    On Error Resume Next   ' Defer error handling.

    
    Dim i As Integer
    Dim y
        
    
    ccode = NWCallsInit(0, 0)
    If ccode <> 0 Then
        MsgBox "There was an error connecting to the Netware DLL files. Netware Error: " & Chr(13) & ccode, vbCritical
    End If
    
    'gServerName = UCase(InputBox("Enter Server Name"))
    gServerName = ServerName
   
    
    ccode = NWCCOpenConnByName(0, gServerName, NWCC_NAME_FORMAT_BIND, NWCC_OPEN_UNLICENSED, 0, connHandle)
    If (ccode <> 0) Then
         MsgBox ("WARNING: Could connect to the server [ " & gServerName & " ]. Please check the server name and try again. Netware error: " & Hex(ccode))
         BtnServerAdmin.Enabled = False
         BtnUtilization.Enabled = False
         Exit Sub
    End If
    
     '---------------Get critical info -------------------
    frmCriticalStats.Show
    frmCriticalStats.Text1.Text = ServerName
    frmCriticalStats.Timer1.Enabled = True
    
    
    Call getBasicInfo
'----------------------------------------------------
    LblCPU.Caption = statCPU
    LblUptime.Caption = statUptime
    LblOther.Caption = statOther
'-----------------------------------------------------
        
   If Err.Number <> 0 Then
   msg = "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & Err.Description
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
   End If
        
End Sub

Private Sub getBasicInfo()
    Dim volName As String * 50
    Dim numVol As Integer
    Dim numMaxEver As Integer
    Dim res As Integer
    Dim majorVer As Byte
    Dim minVer As Byte
    Dim rev As Byte
    Dim maxConns As Integer
    Dim ConnsInUse As Integer
    Dim SFTLevel As Byte
    Dim TTSLevel As Byte
    Dim i As Integer
    Dim connNumber As Integer
    Dim objname As String * 50
    Dim serialNum As Long
    Dim appNum As Integer
    Dim versBuffer As VERSION_INFO
    Dim connInfoBuffer As tagNWCCConnInfo
    Dim TreeName As String
    Dim itmX As ListItem
    Dim l As Integer
    Dim tmpServerName As String * 100

    Set itmX = lstServer.ListItems.Add(, , "Server Name")
    itmX.SubItems(1) = Me.gServerName
    res = 0
    
    'Insert server name into the critical label field.
    LblServerName.Caption = gServerName
    
    
    ccode = NWGetFileServerInformation(connHandle, tmpServerName, majorVer, minVer, rev, maxConns, numMaxEver, ConnsInUse, numVol, SFTLevel, TTSLevel)
    If (ccode <> 0) Then
           MsgBox ("NWGetFileServerInformation returned:  " & Hex(ccode))
           Exit Sub
    End If

    ccode = NWGetFileServerVersionInfo(Me.connHandle, versBuffer)
    If (ccode <> 0) Then
         MsgBox ("NWGetFileServerVersionInfo returned:  " & Hex(ccode))
         Exit Sub
    End If

    Set itmX = lstServer.ListItems.Add(, , "Max Connections Used")
    itmX.SubItems(1) = versBuffer.maxConnectionsEverUsed

    i = 0
    Do
        volName = "^^^"
        ccode = NWGetVolumeName(Me.connHandle, i, volName)
        If InStr(1, volName, "^^") > 0 Then
            Set itmX = lstServer.ListItems.Add(, , "Number Of Volumes")
            itmX.SubItems(1) = i
            Exit Do
        End If
        i = i + 1
        If (ccode <> 0) Then
            MsgBox ("NWGetVolumeName returned:  " & Hex(ccode))
        End If
    Loop Until (i = 255)
    
    Set itmX = lstServer.ListItems.Add(, , "Server Version")
    itmX.SubItems(1) = majorVer & "." & minVer
    
    Set itmX = lstServer.ListItems.Add(, , "Max Connections")
    itmX.SubItems(1) = maxConns
    
    Set itmX = lstServer.ListItems.Add(, , "Connections in Use")
    itmX.SubItems(1) = ConnsInUse

    Set itmX = lstServer.ListItems.Add(, , "SFT Level")
    itmX.SubItems(1) = SFTLevel
 
    Set itmX = lstServer.ListItems.Add(, , "TTS Level")
    itmX.SubItems(1) = TTSLevel

    ccode = NWGetConnectionNumber(Me.connHandle, connNumber)
    If (ccode <> 0) Then
        MsgBox ("NWGetConnectionNumber returned:  " & Hex(ccode))
    End If

    ccode = NWGetConnectionInformation(Me.connHandle, connNumber, objname, 0, 0, 0)
    If ccode = 0 Then
        Set itmX = lstUser.ListItems.Add(, , "Login Id")
        itmX.SubItems(1) = objname
        
        Set itmX = lstUser.ListItems.Add(, , "Connection Number")
        itmX.SubItems(1) = connNumber
    
        Set itmX = lstUser.ListItems.Add(, , "Console Operator")
        ccode = NWCheckConsolePrivileges(Me.connHandle)
    End If
    
    serialNum = 0
    ccode = NWGetNetworkSerialNumber(Me.connHandle, serialNum, appNum)
    If (ccode = 0) Then
        serialNum = NWLongSwap(serialNum)
        Set itmX = lstServer.ListItems.Add(, , "Serial Number")
        itmX.SubItems(1) = Hex(serialNum) & ":" & Hex(appNum)
    End If
    
    ccode = NWCCGetAllConnInfo(Me.connHandle, 1, connInfoBuffer)
    If ccode <> 0 Then
        MsgBox "NWCCGetAllConnInfo Returned: " & Hex(ccode), vbExclamation
    End If

    Set itmX = lstUser.ListItems.Add(, , "Authentication State")
    Select Case connInfoBuffer.authenticationState
    Case 0
        itmX.SubItems(1) = "None"
    Case 1
        itmX.SubItems(1) = "Bindery"
    Case 2
        itmX.SubItems(1) = "NDS"
    End Select

    Set itmX = lstUser.ListItems.Add(, , "Broadcast State")
    Select Case connInfoBuffer.broadcastState
    Case 0
        itmX.SubItems(1) = "Permit All"
    Case 1
        itmX.SubItems(1) = "Permit System"
    Case 2
        itmX.SubItems(1) = "Permit None"
    Case 3
        itmX.SubItems(1) = "Permit Poll"
    End Select

    Set itmX = lstUser.ListItems.Add(, , "Conn Reference")
    itmX.SubItems(1) = connInfoBuffer.connRef

    If majorVer <> 3 And majorVer <> 2 Then
        Set itmX = lstServer.ListItems.Add(, , "Tree Name")
        While connInfoBuffer.TreeName(l) <> 95
            TreeName = TreeName & Chr(connInfoBuffer.TreeName(l))
            l = l + 1
        Wend
        itmX.SubItems(1) = TreeName
    End If

    Set itmX = lstUser.ListItems.Add(, , "NDS State")
    Select Case connInfoBuffer.NDSState
    Case 0
        itmX.SubItems(1) = "NDS_NOT_CAPABLE"
    Case 1
        itmX.SubItems(1) = "NDS_CAPABLE"
    End Select

    Set itmX = lstUser.ListItems.Add(, , "Max Packet Size")
    itmX.SubItems(1) = connInfoBuffer.maxPacketSize

    Set itmX = lstUser.ListItems.Add(, , "License State")
    Select Case connInfoBuffer.licenseState
    Case 0
        itmX.SubItems(1) = "NOT_LICENSED"
    Case 1
        itmX.SubItems(1) = "CONNECTION_LICENSED"
    Case 2
        itmX.SubItems(1) = "HANDLE_LICENSED"
    End Select

    Set itmX = lstUser.ListItems.Add(, , "Distance")
    itmX.SubItems(1) = TickToStr(connInfoBuffer.distance)



End Sub

Private Sub c(Cancel As Integer)
    ccode = NWCCCloseConn(connHandle)
    If (ccode <> 0) Then
           MsgBox ("NWCCCloseConn returned:  " & Hex(ccode))
    End If
    
End Sub

Public Function TickToStr(Ticks As Long) As String
    Dim TickPs, TickPm, TickPh, TickPd 'real
    Dim Days As Single
    Dim hours As Single
    Dim minutes As Double
    Dim seconds As Double

    TickPs = 18.205         '18.205
    TickPm = 18.205 * 60    '1092.3
    TickPh = 18.205 * 60 * 60 '65538
    TickPd = 18.205 * 60 * 60 * 24 '1572912
    Days = Ticks / TickPd
    Days = Fix(Days)
      Ticks = Ticks - (Days * TickPd)
    hours = Ticks / TickPh
    hours = Fix(hours)
      Ticks = Ticks - (hours * TickPh)
    minutes = Ticks / TickPm
    minutes = Fix(minutes)
      Ticks = Ticks - (minutes * TickPm)
    seconds = Ticks / TickPs
    seconds = Fix(seconds)
      Ticks = Ticks - (seconds * TickPs)
    If Ticks < 0 Then
     Ticks = 0
    End If
    If Days < 0 Then
     Days = 0
    End If
    If hours < 0 Then
     hours = 0
    End If
    If minutes < 0 Then
     minutes = 0
    End If
    If seconds < 0 Then
     seconds = 0
    End If
    If Ticks > 17 Then
      seconds = seconds + 1
      Ticks = 0
    End If
    If seconds > 59 Then
      minutes = minutes + 1
      seconds = 0
    End If
    If minutes > 59 Then
      hours = hours + 1
      minutes = 0
    End If
    If hours > 23 Then
      Days = Days + 1
      hours = 0
    End If
    TickToStr = Days & " Days " & hours & " Hrs " & minutes & " Min " & seconds & " Sec " & Ticks & " Ticks"
End Function





