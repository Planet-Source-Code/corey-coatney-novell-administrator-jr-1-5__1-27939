VERSION 5.00
Begin VB.Form frmCriticalStats 
   Caption         =   "Loading..."
   ClientHeight    =   2400
   ClientLeft      =   6165
   ClientTop       =   4635
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   2985
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtUptime 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtCPU 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1200
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1200
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Info"
      Height          =   372
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Text            =   "cinfs0001"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label LblUptime 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label LblOther 
      Caption         =   "Other"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label LblCPU 
      Caption         =   "CPU"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmCriticalStats"
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
' Project: vbuptime.vbp
'
'    Desc: Sample code which demonstrates how to use DLL function
'          calls in VB when getting server`s up time.
'          This code uses NWGetCPUInfo() function call which is available
'          on 4.x and 5.x servers only !
'          User should be already authenticated on given server
'          before running this program.
'          Used DLL calls: NWCallsInit(), NWCCOpenConnByName(), NWGetCPUInfo(),
'          NWCCCloseConn()
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
'   99 January     RLE     Initial code
'
'=====================================================================
'
' It`s a good idea to have the following 'Option explicit' switched ON
'   to avoid unexpectable results from VB Variant types
'   whenever we need to call DLL functions.
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

Private Const FSE_CPU_STR_MAX = 16
Private Const FSE_COPROCESSOR_STR_MAX = 48
Private Const FSE_BUS_STR_MAX = 32

Private Type SERVER_AND_VCONSOLE_INFO
            currentServerTime As Long
            vconsoleVersion As Byte
            vconsoleRevision As Byte
End Type

Private Type CPU_INFO
            pageTableOwnerFlag As Long
            CPUTypeFlag As Long
            coProcessorFlag As Long
            busTypeFlag As Long
            IOEngineFlag As Long
            FSEngineFlag As Long
            nonDedicatedFlag As Long
End Type

Private Type NWFSE_CPU_INFO
            serverTimeAndVConsoleInfo As SERVER_AND_VCONSOLE_INFO
            reserved As Integer
            numOfCPUs As Long
            CPUInfo As CPU_INFO
End Type

Private Declare Function NWCallsInit Lib "calwin32" _
    (reserved1 As Byte, reserved2 As Byte) As Long
    
Private Declare Function NWCCOpenConnByName Lib "clxwin32" _
    (ByVal startConnHandle As Long, ByVal Name As String, _
     ByVal nameFormat As Long, ByVal openState As Long, _
     ByVal tranType As Long, pConnHandle As Long) As Long

Private Declare Function NWGetCPUInfo Lib "calwin32" _
    (ByVal conn As Long, ByVal CPUNum As Long, _
    ByVal CPUName As String, ByVal numCoprocessor As String, _
    ByVal bus As String, fseCPUInfo As NWFSE_CPU_INFO) As Long
    
Private Declare Function NWCCCloseConn Lib "clxwin32" _
    (ByVal connHandle As Long) As Long

Private Sub Command1_Click()
Dim msg

On Error Resume Next   ' Defer error handling.

Dim connHandle As Long, retCode As Long
Dim CPUInfo As NWFSE_CPU_INFO
Dim CPUNum As Long, upTime As Long
Dim upTimeDays As Long, upTimeHours As Long
Dim upTimeMinutes As Long, upTimeSeconds As Long
Dim CPUName As String * FSE_CPU_STR_MAX
Dim COPRName As String * FSE_COPROCESSOR_STR_MAX
Dim BUSName As String * FSE_BUS_STR_MAX

    If Text1.Text = "" Then
        Exit Sub
    End If
    
    retCode = NWCCOpenConnByName(0, Text1.Text, NWCC_NAME_FORMAT_BIND, _
                    NWCC_OPEN_LICENSED, NWCC_TRAN_TYPE_WILD, connHandle)
    If retCode <> 0 Then
        Err.Raise retCode, "", "NWCCOpenConnByName() failed !"
        Exit Sub
    End If

   ' List1.Clear
    retCode = NWGetCPUInfo(connHandle, CPUNum, CPUName, COPRName, BUSName, CPUInfo)
    If retCode <> 0 Then
       ' List1.AddItem "NWGetCPUInfo returns E=" + Str(retCode)
    Else
       ' List1.AddItem "CPU:" + CPUName
        '------------------------------------->CPU name
        txtCPU.Text = CPUName
        
       ' List1.AddItem COPRName
        '---------------------------------->coprocessor name
        txtOther.Text = COPRName
' We have got uptime in ticks (1sec ~ 18.21 ticks),
'   so we have to do some recalculations...
        upTime = CPUInfo.serverTimeAndVConsoleInfo.currentServerTime / 18.2065
        upTimeDays = upTime \ 86400
        upTimeHours = upTime \ 3600 - (upTimeDays * 24)
        upTimeMinutes = upTime \ 60 - (upTimeHours + upTimeDays * 24) * 60
        upTimeSeconds = upTime Mod 60
   '     List1.AddItem "Uptime in seconds:" + Str(upTime)
   '     List1.AddItem "Uptime: " + Str(upTimeDays) + "days, " + Str(upTimeHours) + ":" + Str(upTimeMinutes) + ":" + Str(upTimeSeconds)
        '------------------------>Uptime to form
        txtUptime.Text = Str(upTimeDays) + " days " + Str(upTimeHours) + " hours" + Str(upTimeMinutes) + " min" + Str(upTimeSeconds) + " sec"
    End If
    retCode = NWCCCloseConn(connHandle)
    
 If Err.Number <> 0 Then
   msg = "Error # " & Str(Err.Number) & " was generated by [Critical Stats] " _
         & Err.Source & Chr(13) & Err.Description
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
End If
    
End Sub

Private Sub Form_Load()
Dim msg

On Error Resume Next   ' Defer error handling.

Dim retCode As Long
    retCode = NWCallsInit(0, 0)
    If retCode <> 0 Then
        Err.Raise retCode, "NWCallsInit", "NWCallsInit() - Cannot initialize !"
        Unload Me
    End If
    
If Err.Number <> 0 Then
   msg = "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & Err.Description
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
End If




End Sub


Private Sub Timer1_Timer()
Dim y

Command1_Click
Timer1.Enabled = False

statCPU = txtCPU.Text
statOther = txtOther.Text
statUptime = txtUptime.Text
'y = MsgBox("Here are the values for CPU: " & statCPU & " Other: " & statOther & " CPU: " & statCPU, vbOKOnly, "Critical Stats")

Unload Me
End Sub
