VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmObjects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Novell Servers"
   ClientHeight    =   4320
   ClientLeft      =   4785
   ClientTop       =   3540
   ClientWidth     =   4335
   Icon            =   "frmObjects.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4335
   Begin VB.CommandButton BtnRemoveServer 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Timer tmrDisconnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   3240
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   3240
   End
   Begin VB.Timer tmrAnimate2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5160
      Top             =   3240
   End
   Begin VB.Timer tmrAnimate1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4680
      Top             =   3240
   End
   Begin VB.CommandButton BtnUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton BtnAddServer 
      Caption         =   "Add "
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjects.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjects.frx":0BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjects.frx":339A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjects.frx":36B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjects.frx":39D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjects.frx":3CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjects.frx":400A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5953
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Server"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   1553
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Uptime"
         Object.Width           =   3723
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prStatus 
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar staStatus 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   3960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3880
            MinWidth        =   3880
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3441
            MinWidth        =   3441
         EndProperty
      EndProperty
   End
   Begin VB.Image imgEmpty 
      Height          =   255
      Left            =   5280
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   3
      Left            =   5760
      Picture         =   "frmObjects.frx":528E
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   2
      Left            =   5760
      Picture         =   "frmObjects.frx":5618
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   1
      Left            =   5760
      Picture         =   "frmObjects.frx":59A2
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   0
      Left            =   5760
      Picture         =   "frmObjects.frx":5D2C
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   2
      Left            =   5520
      Picture         =   "frmObjects.frx":60B6
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   1
      Left            =   5520
      Picture         =   "frmObjects.frx":6440
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   0
      Left            =   5520
      Picture         =   "frmObjects.frx":67CA
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   3
      Left            =   6000
      Picture         =   "frmObjects.frx":6B54
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   2
      Left            =   6000
      Picture         =   "frmObjects.frx":6EDE
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   1
      Left            =   6000
      Picture         =   "frmObjects.frx":7268
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   0
      Left            =   6000
      Picture         =   "frmObjects.frx":75F2
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgConnected 
      Height          =   240
      Left            =   6240
      Picture         =   "frmObjects.frx":797C
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Declare Variables

Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String


Private Sub BtnAddServer_Click()
    Dim FileDir As String
    
    FileDir = App.Path + "\" + App.EXEName + ".dat"
    ServerName = UCase(InputBox("Enter Server Name", "Add Novell Server"))
    
    Dim nodX As Node
    Dim i As Integer
    Dim intIndex
   
    'Create an object variable for the ListItem object
    Dim itmR As ListItem
    Dim intCount As Integer
    'Create an object variable for the ColumnHeader object
    
    If ServerName <> "" Then
        
        
        Set itmR = ListView1.ListItems.Add(, , ServerName)
        itmR.SubItems(1) = "Offline"
        '----------------------------------
        'Add server to dat file
        '----------------------------------
        If FileExists(FileDir) Then
            Open (FileDir) For Append As #1
            DoEvents
            Print #1, ServerName
            Close #1
        Else
            MsgBox "Error Loading the " + App.EXEName + ".dat file: " & Err.Description & ". A new one will be created.", vbCritical + vbSystemModal, "Error"
            CreateDat
        End If
        
    End If
    

End Sub

Private Sub BtnRemoveServer_Click()
    On Error Resume Next
    With Me.ListView1
    If .ListItems.count > 0 Then
      .ListItems.Remove .SelectedItem.Index
    End If
    If .ListItems.count > 0 Then
      .ListItems(.SelectedItem.Index).Selected = True
    End If
    .SetFocus
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Update Layout in Ini
Dim FileExist
Dim FileDir As String
Dim Save As Long

FileDir = App.Path + "\" + App.EXEName + ".ini"

If Me.WindowState <> vbMinimized And FileExists(FileDir) Then
Dim OLeft, OTop, OWidth, OHeight
        OLeft = Me.Left
        OTop = Me.Top
        OWidth = Me.Width
        OHeight = Me.Height
                
                
    '----------------------------------
    'Save Form Layout
    '----------------------------------
    KeySection = "Layout"
    KeyKey = "ObjectsLeft"
    KeyValue = OLeft
    SaveIni

    KeySection = "Layout"
    KeyKey = "ObjectsTop"
    KeyValue = OTop
    SaveIni

    KeySection = "Layout"
    KeyKey = "ObjectsWidth"
    KeyValue = OWidth
    SaveIni

    KeySection = "Layout"
    KeyKey = "ObjectsHeight"
    KeyValue = OHeight
    SaveIni

ObjectShowing = False
'----------------------------------------------
End If

'Save Dat file
FileDir = App.Path + "\" + App.EXEName + ".dat"
On Error Resume Next

Open FileDir For Output As #1
    For Save = 0 To ListView1.ListItems.count
        Print #1, ListView1.ListItems(Save)
    Next Save
Close #1

End Sub

Private Sub BtnUpdate_Click()
 tmrConnect.Enabled = True
    
    'Set the pictures to working
    
    staStatus.Panels(1).Text = "Updating Status..."
    staStatus.Panels(1).Picture = imgAnimate2(0).Picture
    tmrAnimate2.Enabled = True
End Sub
Private Sub SaveIni()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\" & App.EXEName & ".ini"
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function to save Server Layout", vbExclamation)
End If

End Sub
Private Sub Form_Load()


    Dim nodX As Node
    Dim i As Integer
    Dim intIndex
   
    'Create an object variable for the ListItem object
    Dim itmR As ListItem
    Dim intCount As Integer
    'Create an object variable for the ColumnHeader object
    Dim clmX As ColumnHeader
    
    
  
    
    'Add ColumnHeaders
    'The width of the columns
    'is the width of the control divided by the
    'number of ColumnHeader objects
   
    'Set BorderStyle property
    ListView1.BorderStyle = ccFixedSingle
    
    'Clear the views
    ListView1.ListItems.Clear



'Status Bar Code
'Set the AnimateStatus to 0
    
    'Place the ProgressBar into the panel
    prStatus.Left = staStatus.Panels(3).Left + 40
    prStatus.Top = staStatus.Top + 60
    prStatus.Width = staStatus.Width - staStatus.Panels(3).Left - 80
    prStatus.Height = staStatus.Height - 90

'-------------------------------------------------------------------




End Sub


Private Sub ListView1_DblClick()
Dim SelectedServer
Dim y

Dim frmS As frmServerStatistics
Set frmS = New frmServerStatistics
    
SelectedServer = ListView1.SelectedItem.Text
ServerName = SelectedServer

frmS.Caption = " Server " & ServerName & " Statistics"
frmS.Show


End Sub

Private Sub tmrAnimate1_Timer()
  'If the Animation loop has finished restart it
    If AnimateStatus1 = imgAnimate1.count Then
        AnimateStatus1 = 0
    End If
    'Replace the actual StatusBar picture with the next one
    staStatus.Panels(1).Picture = imgAnimate1(AnimateStatus1).Picture
    AnimateStatus1 = AnimateStatus1 + 1
End Sub

Private Sub tmrAnimate2_Timer()
    'If the Animation loop has finished restart it
    If AnimateStatus2 = imgAnimate2.count Then
        AnimateStatus2 = 0
    End If
    'Replace the actual StatusBar picture with the next one
    staStatus.Panels(1).Picture = imgAnimate2(AnimateStatus2).Picture
    AnimateStatus2 = AnimateStatus2 + 1
End Sub

Private Sub tmrConnect_Timer()
  
    If prStatus.Value >= 99.0001 Then
        prStatus.Value = 100
        GoTo DoneConnecting
    End If
    'Add 1 to the value of the ProgressBar
    prStatus.Value = prStatus.Value + 1
    
    
DoneConnecting:
    'Display the percent value of the progressbar in the StatusBar
    staStatus.Panels(2).Text = Round(prStatus.Value) & "%"
    'If the ProgressBar reaches 100 the computer is "connected"
    If prStatus.Value >= 100 Then
    'Set the properties of the controls to "connected"
    prStatus.Value = prStatus.Min
    tmrAnimate2.Enabled = False
    tmrConnect.Enabled = False
    staStatus.Panels(1).Picture = imgEmpty.Picture
    staStatus.Panels(1).Text = "Updated: " & Time
    
    End If
End Sub

