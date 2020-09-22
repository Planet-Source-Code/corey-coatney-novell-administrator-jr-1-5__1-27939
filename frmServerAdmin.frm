VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServerAdmin 
   Caption         =   "Server Administration"
   ClientHeight    =   6420
   ClientLeft      =   2925
   ClientTop       =   2445
   ClientWidth     =   9735
   Icon            =   "frmServerAdmin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   9735
   Begin MSComctlLib.ListView lstConnDetail 
      Height          =   5895
      Left            =   4800
      TabIndex        =   9
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10398
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
         Text            =   "Item Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstConn 
      Height          =   5895
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   615
      Left            =   8640
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Drop Conn"
      Height          =   615
      Left            =   8640
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send Message"
      Height          =   615
      Left            =   8640
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   8640
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Disable Logon"
      Height          =   615
      Left            =   8640
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Enable Logon"
      Height          =   615
      Left            =   8640
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.OptionButton optName 
      Caption         =   "Sort by Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optAddr 
      Caption         =   "Sort By Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmServerAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NWCC_OPEN_UNLICENSED = &H2
Const OT_FILE_SERVER = &H400
Const NWCC_NAME_FORMAT_BIND = &H2

Private Type tNWINET_ADDR
            networkAddr(3) As Byte
            netNodeAddr(5) As Byte
            socket As Integer
            connType As Integer
End Type

Private Type MYLOGINTIME 'Custom code I added to get around VB limitation of no pointers!!!
            logintime(5) As Byte
End Type

Private Type SERVER_AND_VCONSOLE_INFO
            currentServerTime As Long
            vconsoleVersion As Byte
            vconsoleRevision As Byte
End Type

Private Type USER_INFO
            connNum As Long
            useCount As Long
            connServiceType As Byte
            logintime(6) As Byte
            status As Long
            expirationTime As Long
            objType As Long
            transactionFlag As Byte
            logicalLockThreshold As Byte
            recordLockThreshold As Byte
            fileWriteFlags As Byte
            fileWriteState As Byte
            filler As Byte
            fileLockCount As Integer
            recordLockCount As Integer
            totalBytesRead(5) As Byte
            totalBytesWritten(5) As Byte
            totalRequests As Long
            heldRequests As Long
            heldBytesRead(5) As Byte
            heldBytesWritten(5) As Byte
End Type


Private Type NWFSE_USER_INFO
            serverTimeAndVConsoleInfo As SERVER_AND_VCONSOLE_INFO
            reserved As Long
            userInfo As USER_INFO
End Type

Private Declare Function NWCallsInit Lib "calwin32" (reserved1 As Byte, reserved2 As Byte) As Long
Private Declare Function NWCCOpenConnByName Lib "clxwin32" (ByVal startConnHandle As Long, ByVal name1 As String, ByVal nameFormat As Long, ByVal openState As Long, ByVal tranType As Long, pConnHandle As Long) As Long
Private Declare Function NWGetFileServerInformation Lib "calwin32" (ByVal conn As Long, ByVal ServerName As String, majorVer As Byte, minVer As Byte, rev As Byte, maxConns As Integer, maxConnsUsed As Integer, ConnsInUse As Integer, numVolumes As Integer, SFTLevel As Byte, TTSLevel As Byte) As Long
Private Declare Function NWGetConnectionInformation Lib "calwin32" (ByVal connHandle As Long, ByVal connNumber As Integer, ByVal pObjName As String, pObjType As Integer, pObjID As Long, logintime As MYLOGINTIME) As Long
Private Declare Function NWGetInetAddr Lib "calwin32" (ByVal connHandle As Long, ByVal connNum As Integer, pInetAddr As tNWINET_ADDR) As Long
Private Declare Function NWGetUserInfo Lib "calwin32" (ByVal conn As Long, ByVal connNum As Long, ByVal userName As String, fseUserinfo As NWFSE_USER_INFO) As Long
Private Declare Function NWClearConnectionNumber Lib "calwin32" (ByVal connHandle As Long, ByVal connNumber As Integer) As Long
Private Declare Function NWSendBroadcastMessage Lib "calwin32" (ByVal conn As Long, ByVal message As String, ByVal connCount As Integer, connList As Integer, resultList As Byte) As Long
Private Declare Function NWDisableFileServerLogin Lib "calwin32" (ByVal conn As Long) As Long
Private Declare Function NWEnableFileServerLogin Lib "calwin32" (ByVal conn As Long) As Long

Dim ccode As Long
Public connHandle
Public gServerName As String


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim connNum As Integer
    Dim itmX As ListItem

    If MsgBox("If this a active client 32 connection the client will automatically reconnect to the server." & Chr$(13) & "This may cause a server-client conflict loop" & Chr$(13) & "Do you want to drop this connection?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    Set itmX = lstConn.SelectedItem
    connNum = itmX.SubItems(2)
   
    ccode = NWClearConnectionNumber(connHandle, connNum)
    If (ccode <> 0) Then
        If Hex(ccode) = "89C6" Then
            MsgBox "You need console privledges to perform this operation"
            Exit Sub
        End If
        MsgBox ("NWClearConnectionNumber Returned: " & Hex(ccode))
    End If
    GetConnList
End Sub

Private Sub Command3_Click()
    Dim msg As String
    Dim connNum As Integer
    Dim itmX As ListItem

    msg = InputBox("Enter a Message to Send", "Message")
    Set itmX = lstConn.SelectedItem
    connNum = itmX.SubItems(2)
 
    ccode = NWSendBroadcastMessage(connHandle, msg + Chr$(0), 1, connNum, 0)
    If ccode <> 0 Then
        MsgBox "NWSendBroadcast Message Returned: " & Hex(ccode)
    End If
End Sub

Private Sub Command4_Click()
    Call GetConnList
End Sub

Private Sub Command5_Click()
    If MsgBox("This will cause the NetWare server to refuse all login requests," & Chr$(13) & "Do you wish to Disable all Login requests?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    ccode = NWDisableFileServerLogin(connHandle)
    If ccode <> 0 Then
        If Hex(ccode) = "89C6" Then
            MsgBox "You need console privledges to perform this operation"
            Exit Sub
        End If
        MsgBox "NWDisableFileServerLogin Returned: " & Hex(ccode)
    Else
        MsgBox "Logins have been disabled"
    End If
End Sub

Private Sub Command6_Click()
    ccode = NWEnableFileServerLogin(connHandle)
    If ccode <> 0 Then
        If Hex(ccode) = "89C6" Then
            MsgBox "You need console privledges to perform this operation"
            Exit Sub
        End If
        MsgBox "NWEnableFileServerLogin Returned: " & Hex(ccode)
    Else
        MsgBox "Logins have been enabled"
    End If
End Sub

Private Sub Form_Load()
Dim msg

 On Error Resume Next   ' Defer error handling.
    
    Dim i As Integer
    
    ccode = NWCallsInit(0, 0)
    If ccode <> 0 Then
        MsgBox "NWCallsInit failed with: " & Chr(13) & ccode, vbCritical
    End If
    
    gServerName = ServerName
    
    ccode = NWCCOpenConnByName(0, gServerName + Chr(0), NWCC_NAME_FORMAT_BIND, NWCC_OPEN_UNLICENSED, 0, connHandle)
    If (ccode <> 0) Then
         MsgBox ("WARNING: NWCCOpenConnByName returned:  " & Hex(ccode))
         Exit Sub
    End If
    
    Call GetConnList
    
    
    If Err.Number <> 0 Then
   msg = "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & Err.Description
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
    End If
        
End Sub

Private Sub GetConnList()
    Dim i As Integer
    Dim itmX As ListItem
    Dim loginid As String * 50
    Dim maxConns As Integer
    Dim pObjID As Long
    Dim LoginType As Integer
    Dim nLoginTime As MYLOGINTIME
    Dim ConnsInUse As Integer
    Dim addr As tNWINET_ADDR
    Dim nodeaddr As String
    Dim loginstr As String
    Dim fseUserinfo As NWFSE_USER_INFO
    Dim strServerName As String * 50

    lstConn.ListItems.Clear
    lstConn.ColumnHeaders.Clear
    ccode = NWGetFileServerInformation(connHandle, strServerName, 0, 0, 0, maxConns, 0, ConnsInUse, 0, 0, 0)
    If (ccode <> 0) Then
           MsgBox ("NWGetFileServerInformation returned:  " & Hex(ccode))
           Exit Sub
    End If
    
    Caption = strServerName & " Connections"
    
    If optName.Value = True Then
        lstConn.ColumnHeaders.Add , , "Login Id", 2800
        lstConn.ColumnHeaders.Add , , "Node Addr"
        lstConn.ColumnHeaders.Add , , "Conn #"
      '  lstConn.ColumnHeaders.Add , , "Type"
    Else
        lstConn.ColumnHeaders.Add , , "Node Addr"
        lstConn.ColumnHeaders.Add , , "Login Id", 2800
        lstConn.ColumnHeaders.Add , , "Conn #"
      '  lstConn.ColumnHeaders.Add , , "Type"
    End If

    For i = 1 To maxConns
        ccode = NWGetConnectionInformation(connHandle, i, loginid, LoginType, pObjID, nLoginTime)
        If (ccode = 0) Then
            ccode = NWGetInetAddr(connHandle, i, addr)
            If (ccode <> 0) Then
                MsgBox ("NWGetInetAddr Returned: " & Hex(ccode))
                Exit Sub
            End If
    
            nodeaddr = Hex(addr.netNodeAddr(0)) & Hex(addr.netNodeAddr(1)) & Hex(addr.netNodeAddr(2)) & Hex(addr.netNodeAddr(3)) & Hex(addr.netNodeAddr(4)) & Hex(addr.netNodeAddr(5))
            loginstr = loginid
    
            If optName.Value = True Then
                Set itmX = lstConn.ListItems.Add(, , loginstr)
                itmX.SubItems(1) = nodeaddr
                itmX.SubItems(2) = i
     '          itmX.SubItems(3) = addr.connType
            Else
                Set itmX = lstConn.ListItems.Add(, , nodeaddr)
                itmX.SubItems(1) = loginstr
                itmX.SubItems(2) = i
     '           itmX.SubItems(3) = LoginType
            End If
            ccode = NWGetUserInfo(connHandle, i, loginid, fseUserinfo)
        Else
            If Hex(ccode) = "89FB" Then
                Unload Me
                Exit Sub
            End If
'            MsgBox "NWGetConnectionInformation Returned: " & Hex(ccode)
'            Exit Sub
        End If
    Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub lstConn_Click()
    Dim itmX As ListItem
    Dim itmXd As ListItem
    Dim connNum As Integer
    Dim tmpsocket As String
    Dim objType As Integer
    Dim strConnType As String * 6
    Dim MYLOGINTIME As MYLOGINTIME
    Dim objId As Long
    Dim tmpname As String * 50
    Dim addr As tNWINET_ADDR
    Dim fseUserinfo As NWFSE_USER_INFO
        
    lstConnDetail.ListItems.Clear
    Set itmX = lstConn.SelectedItem
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Login Id") 'Populate Login ID
    If optName.Value = True Then
        itmXd.SubItems(1) = itmX.Text
    Else
        itmXd.SubItems(1) = itmX.SubItems(1)
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Node Address")
    If optName.Value = True Then
        itmXd.SubItems(1) = itmX.SubItems(1)
    Else
        itmXd.SubItems(1) = itmX.Text
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Connection")
    connNum = itmX.SubItems(2)
    itmXd.SubItems(1) = connNum
    
    ccode = NWGetConnectionInformation(connHandle, Val(itmX.SubItems(2)), vbNullString, objType, objId, MYLOGINTIME)
    If ccode <> 0 Then
        MsgBox "NWGetconnectionInformation Returned: " & Hex(ccode)
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Object Type")
    If objType = 256 Then
        itmXd.SubItems(1) = "OT_USER"
    Else
        itmXd.SubItems(1) = objType
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Object Id")
    itmXd.SubItems(1) = objId
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Login Time")
    itmXd.SubItems(1) = MYLOGINTIME.logintime(1) & "\" & MYLOGINTIME.logintime(2) & "\" & MYLOGINTIME.logintime(0) & " " & MYLOGINTIME.logintime(3) & ":" & MYLOGINTIME.logintime(4) & ":" & MYLOGINTIME.logintime(5)
    
    ccode = NWGetInetAddr(connHandle, connNum, addr)
    If (ccode <> 0) Then
        MsgBox ("NWGetInetAddr Returned: " & Hex(ccode))
        Exit Sub
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Connection Type")
    Select Case addr.connType
    Case 2
        strConnType = "NCP"
    Case 3
        strConnType = "NLM"
    Case 4
        strConnType = "AFP"
    Case 5
        strConnType = "FTAM"
    Case 6
        strConnType = "ANCP"
    Case Else
        strConnType = addr.connType
    End Select
    
    itmXd.SubItems(1) = strConnType
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Network Address")
    itmXd.SubItems(1) = Hex(addr.networkAddr(0)) & Hex(addr.networkAddr(1)) & Hex(addr.networkAddr(2)) & Hex(addr.networkAddr(3))
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Socket Number")
    tmpsocket = Hex(addr.socket)
    itmXd.SubItems(1) = Right(tmpsocket, 2) & Left(tmpsocket, 2)
    
    ccode = NWGetUserInfo(connHandle, connNum, tmpname, fseUserinfo)
    If ccode <> 0 Then
        If Hex(ccode) = "897D" Then
            Set itmXd = lstConnDetail.ListItems.Add(, , "NOT LOGGED IN")
            itmXd.SubItems(1) = True
        End If
        Exit Sub
        'MsgBox "NWGetuserInfo Returned: " & Hex(ccode)
    Else
        Set itmXd = lstConnDetail.ListItems.Add(, , "NOT LOGGED IN")
        itmXd.SubItems(1) = False
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Connection in use")
    If fseUserinfo.userInfo.useCount = 1 Then
        itmXd.SubItems(1) = True
    Else
        itmXd.SubItems(1) = False
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Connection Status")
    itmXd.SubItems(1) = fseUserinfo.userInfo.status
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Expiration Time")
    itmXd.SubItems(1) = fseUserinfo.userInfo.expirationTime 'need to convert this to time
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Transaction Flag")
    itmXd.SubItems(1) = fseUserinfo.userInfo.transactionFlag
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Logical Lock Threshold")
    itmXd.SubItems(1) = fseUserinfo.userInfo.logicalLockThreshold
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Record Lock Threshold")
    itmXd.SubItems(1) = fseUserinfo.userInfo.recordLockThreshold
  
    Set itmXd = lstConnDetail.ListItems.Add(, , "File Write Flags")
    If fseUserinfo.userInfo.fileWriteFlags = 1 Then
        itmXd.SubItems(1) = "FSE_WRITE"
    ElseIf fseUserinfo.userInfo.fileWriteFlags = 2 Then
        itmXd.SubItems(1) = "FSE_WRITE_ABORTED"
    Else
        itmXd.SubItems(1) = fseUserinfo.userInfo.fileWriteFlags
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "File Write State")
    If fseUserinfo.userInfo.fileWriteState = 0 Then
        itmXd.SubItems(1) = "FSE_NOT_WRITING"
    ElseIf fseUserinfo.userInfo.fileWriteState = 1 Then
        itmXd.SubItems(1) = "FSE_WRITE_IN_PROGRESS"
    ElseIf fseUserinfo.userInfo.fileWriteState = 2 Then
        itmXd.SubItems(1) = "FSE_WRITE_BEING_STOPPED"
    Else
        itmXd.SubItems(1) = fseUserinfo.userInfo.fileWriteState
    End If
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "File Lock Count")
    itmXd.SubItems(1) = fseUserinfo.userInfo.fileLockCount
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Record Lock Count")
    itmXd.SubItems(1) = fseUserinfo.userInfo.recordLockCount
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Total Bytes Read")
    itmXd.SubItems(1) = Format(fseUserinfo.userInfo.totalBytesRead(0) + (fseUserinfo.userInfo.totalBytesRead(1) * 256#) + (fseUserinfo.userInfo.totalBytesRead(2) * 256# * 256#) + (fseUserinfo.userInfo.totalBytesRead(3) * 256# * 256 * 256) + (fseUserinfo.userInfo.totalBytesRead(4) * 256 * 256 * 256 * 256#) + (fseUserinfo.userInfo.totalBytesRead(5) * 256 * 256 * 256 * 256 * 256#), "###,###,##0")
          
    Set itmXd = lstConnDetail.ListItems.Add(, , "Total Bytes Written")
    itmXd.SubItems(1) = Format(fseUserinfo.userInfo.totalBytesWritten(0) + (fseUserinfo.userInfo.totalBytesWritten(1) * 256#) + (fseUserinfo.userInfo.totalBytesWritten(2) * 256# * 256#) + (fseUserinfo.userInfo.totalBytesWritten(3) * 256# * 256 * 256) + (fseUserinfo.userInfo.totalBytesWritten(4) * 256 * 256 * 256 * 256#) + (fseUserinfo.userInfo.totalBytesWritten(5) * 256 * 256 * 256 * 256 * 256#), "###,###,##0")
           
    Set itmXd = lstConnDetail.ListItems.Add(, , "Total Number Requests")
    itmXd.SubItems(1) = Format(fseUserinfo.userInfo.totalRequests, "###,###,###,##0")
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Held Requests")
    itmXd.SubItems(1) = fseUserinfo.userInfo.heldRequests
    
    Set itmXd = lstConnDetail.ListItems.Add(, , "Held Bytes Read")
    itmXd.SubItems(1) = Format(fseUserinfo.userInfo.heldBytesRead(0) + (fseUserinfo.userInfo.heldBytesRead(1) * 256#) + (fseUserinfo.userInfo.heldBytesRead(2) * 256# * 256#) + (fseUserinfo.userInfo.heldBytesRead(3) * 256# * 256 * 256) + (fseUserinfo.userInfo.heldBytesRead(4) * 256 * 256 * 256 * 256#) + (fseUserinfo.userInfo.heldBytesRead(5) * 256 * 256 * 256 * 256 * 256#), "###,###,##0")
          
    Set itmXd = lstConnDetail.ListItems.Add(, , "Held Bytes Written")
    itmXd.SubItems(1) = Format(fseUserinfo.userInfo.heldBytesWritten(0) + (fseUserinfo.userInfo.heldBytesWritten(1) * 256#) + (fseUserinfo.userInfo.heldBytesWritten(2) * 256# * 256#) + (fseUserinfo.userInfo.heldBytesWritten(3) * 256# * 256 * 256) + (fseUserinfo.userInfo.heldBytesWritten(4) * 256 * 256 * 256 * 256#) + (fseUserinfo.userInfo.heldBytesWritten(5) * 256 * 256 * 256 * 256 * 256#), "###,###,##0")
    
End Sub

Private Sub optAddr_Click()
    Call GetConnList
End Sub

Private Sub optName_Click()
    Call GetConnList
End Sub
