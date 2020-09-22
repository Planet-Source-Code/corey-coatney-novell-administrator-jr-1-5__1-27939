VERSION 5.00
Begin VB.Form frmUtilization 
   Caption         =   "Server Utilization"
   ClientHeight    =   3195
   ClientLeft      =   4365
   ClientTop       =   4380
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUtilization.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   8115
   Begin VB.Frame Frame1 
      Caption         =   "Server XXX Utilization Statistics"
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   8055
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000012&
         FillColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1575
         Left            =   1080
         Picture         =   "frmUtilization.frx":030A
         ScaleHeight     =   1515
         ScaleWidth      =   6795
         TabIndex        =   4
         Top             =   360
         Width           =   6855
      End
      Begin VB.Label LblUtilization 
         Alignment       =   2  'Center
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Utilization: "
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
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label LblLow 
         Caption         =   "Low:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label LblHigh 
         Caption         =   "High:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Timer TimerTime 
      Interval        =   1000
      Left            =   3120
      Top             =   2760
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Pause"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2400
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3720
      Top             =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Current Interval is 1/2 second. This will be adjustable in upcoming versions."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   5295
   End
   Begin VB.Label LblTimeElapsed 
      Caption         =   "Time Elapsed: "
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmUtilization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
'Copyright ® 1999 Novell, Inc.  All Rights Reserved.
'
'  With respect to this file, Novell hereby grants to Developer a
'  royalty-free, non-exclusive license to include this sample code
'  and derivative binaries in its product. Novell grants to Developer
'  worldwide distribution rights to market, distribute or sell this
'  sample code file and derivative binaries as a component of
'  Developer 's product(s).  Novell shall have no obligations to
'  Developer or Developer's customers with respect to this code.
'
'DISCLAIMER:
'
'  Novell disclaims and excludes any and all express, implied, and
'  statutory warranties, including, without limitation, warranties
'  of good title, warranties against infringement, and the implied
'  warranties of merchantibility and fitness for a particular purpose.
'  Novell does not warrant that the software will satisfy customer's
'  requirements or that the licensed works are without defect or error
'  or that the operation of the software will be uninterrupted.
'  Novell makes no warranties respecting any technical services or
'  support tools provided under the agreement, and disclaims all other
'  warranties, including the implied warranties of merchantability and
'  fitness for a particular purpose.
'
'================================================================
'
' Project: vbutiliz.vbp
'
'    Desc: Sample code which demonstrates how to use DLL function
'          calls in VB when retrieving server's utilization.
'          This code uses NWGetFileServerInfo() function call which is
'          available on 4.x and 5.x servers only !
'          User should be already authenticated on a given server
'          before running this program.
'          Used DLL calls: NWCallsInit(), NWCCOpenConnByName(),
'          NWGetFileServerInfo(), NWCCCloseConn()
'          If you are using VB5, you can encounter byte alignment problems
'          in user defined structure NWFSE_FILE_SERVER_INFO. In such case
'          you have to add one dummy LONG field before serverUtilization.
'
' Programmers:
'
'   Ini       Who                 Firm
'   ------------------------------------------------------------------
'   RLE       Rostislav Letos     Novell DeveloperNet Labs
'
' History:
'
'   When           Who     What
'   ------------------------------------------------------------------
'   99 August      RLE     Initial code
'
'=====================================================================
'
' It`s a good idea to have the following 'Option explicit' switched ON
'   to avoid unexpectable results from implicit VB Variant types
'   when calling DLL functions.
Option Explicit

Private Const NWCC_NAME_FORMAT_BIND = 2
Private Const NWCC_NAME_FORMAT_NDS_TREE = 8
Private Const NWCC_OPEN_LICENSED = 1
Private Const NWCC_OPEN_UNLICENSED = 2

Private Const NWCC_TRAN_TYPE_IPX = 1
Private Const NWCC_TRAN_TYPE_UDP = 2
Private Const NWCC_TRAN_TYPE_DDP = 3
Private Const NWCC_TRAN_TYPE_ASP = 4
Private Const NWCC_TRAN_TYPE_WILD = &H8000

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
    reserved As Integer
    NCPStationsInUseCount As Long
    NCPPeakStationsInUseCount As Long
    numOfNCPRequests As Long
    serverUtilization As Long
    ServerInfo As FSE_SERVER_INFO
    fileServerCounters As FILE_SERVER_COUNTERS
End Type

Private Declare Function NWCallsInit Lib "calwin32" _
    (reserved1 As Byte, reserved2 As Byte) As Long
    
Private Declare Function NWCCOpenConnByName Lib "clxwin32" _
    (ByVal startConnHandle As Long, ByVal Name As String, _
     ByVal nameFormat As Long, ByVal openState As Long, _
     ByVal tranType As Long, pConnHandle As Long) As Long

Private Declare Function NWGetFileServerInfo Lib "calwin32" _
    (ByVal conn As Long, fseFileServerInfo As NWFSE_FILE_SERVER_INFO) As Long
    
Private Declare Function NWCCCloseConn Lib "clxwin32" _
    (ByVal connHandle As Long) As Long


Dim connHandle As Long, retCode As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim Start_Pos
   Dim End_Pos
   Dim X
   Dim A, B
   Dim TodayDate
   Dim Low, High, TempValue
   
   
Private Sub Command1_Click()
Unload Me                   'Unloads the Form
End Sub

Private Sub Command2_Click()
Select Case Timer1.Enabled
    Case True
        Timer1.Enabled = False
        Command2.Caption = "&Resume"
    Case False
        Timer1.Enabled = True
        Command2.Caption = "&Pause"
End Select



End Sub

Private Sub Form_Load()

Dim retCode As Long
    retCode = NWCallsInit(0, 0)
    If retCode <> 0 Then
        Err.Raise retCode, "NWCallsInit", "NWCallsInit() - Cannot initialize !"
        Unload Me
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Start_Pos = 0            'Sets the start position of the cursor
End_Pos = 0                'Sets the end position of the cursor
                           
Me.Height = 3600
Me.ScaleHeight = 2805
Me.ScaleWidth = 7350
Me.Width = 8235
TodayDate = Now

Low = 100
High = 0

Randomize
Call BeginUtilization
Frame1.Caption = "Server Utilization History"

End Sub

Private Sub BeginUtilization()
If ServerName <> "" Then
    

        retCode = NWCCOpenConnByName(0, ServerName, NWCC_NAME_FORMAT_BIND, _
                    NWCC_OPEN_LICENSED, NWCC_TRAN_TYPE_WILD, connHandle)
        If retCode <> 0 Then
            Err.Raise retCode, "", "NWCCOpenConnByName() failed !"
            Exit Sub
        End If
        Timer1.Interval = 500
        Timer1.Enabled = True
    Else
        Exit Sub
        Timer1.Enabled = False
        retCode = NWCCCloseConn(connHandle)
    End If
End Sub




Private Sub Timer1_Timer()
X = X + 1                   'Counts the variable up
If X = 150 Then              'Checking
Picture1.Cls                'If the line has reached the end of the Picture_
                            'control then clear the picture control
X = 0                       'Reinitialize the variable for a new start
Start_Pos = 0               'Reinitialize the variable for a new start
End_Pos = 0                 'Reinitialize the variable for a new start
B = 100                       'Reinitialize the variable for a new start
A = 100                       'Reinitialize the variable for a new start
Else
Dim info As NWFSE_FILE_SERVER_INFO
    retCode = NWGetFileServerInfo(connHandle, info)
    If retCode = 0 Then
        'Label1.Caption = "Current Utilization: " & Str(info.serverUtilization) & "%" 'Sets the Label1 Text
        A = 100 - Val(Str(info.serverUtilization))
        LblUtilization.Caption = 100 - A & "%"
   'Set lows and highs
    TempValue = 100 - A
    
    If TempValue > High Then
        High = TempValue
        LblHigh.Caption = "High: " & High
    End If
    
    If TempValue < Low Then
        Low = TempValue
        LblLow.Caption = "Low: " & Low
    End If
    
    Else
        Label1.Caption = "?"
    End If

   Picture1.ScaleMode = 3     'Sets the ScaleMode of Picture1 to 3
   Picture1.Line (Start_Pos, B)-(End_Pos, A) 'Draws the Line
    B = A                   'Give B the content of A
    Start_Pos = End_Pos     'Give Start_Pos the content of End_Pos
    End_Pos = End_Pos + 3   'Let´s the line grow up for 3 points
End If
End Sub

Private Sub Timer2_Timer()
A = Int((100 * Rnd) + 1)     'Random Values for the plot
X = X + 1                   'Counts the variable up
If X = 150 Then              'Checking
Picture1.Cls                'If the line has reached the end of the Picture_
                            'control then clear the picture control
X = 0                       'Reinitialize the variable for a new start
Start_Pos = 0               'Reinitialize the variable for a new start
End_Pos = 0                 'Reinitialize the variable for a new start
B = 100                       'Reinitialize the variable for a new start
A = 100                       'Reinitialize the variable for a new start
Else

   Picture1.ScaleMode = 3     'Sets the ScaleMode of Picture1 to 3
   Picture1.Line (Start_Pos, B)-(End_Pos, A) 'Draws the Line
    B = A                   'Give B the content of A
    Start_Pos = End_Pos     'Give Start_Pos the content of End_Pos
    End_Pos = End_Pos + 3   'Let´s the line grow up for 3 points
End If

'Label1.Caption = "Current Utilization: " & A & "%" 'Sets the Label1 Text

End Sub

Private Sub TimerTime_Timer()
Dim Elapsed
'Elapsed = DateDiff("n:s", Now, TodayDate)
'LblTimeElapsed.Caption = "Time Elapsed: " & Elapsed


End Sub

