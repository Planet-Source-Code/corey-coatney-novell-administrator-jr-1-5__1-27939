VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memo"
   ClientHeight    =   5460
   ClientLeft      =   5850
   ClientTop       =   3600
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Select Recipient(s)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   4455
      Begin VB.CommandButton btnSendMail 
         Caption         =   "Send"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         ItemData        =   "frmDocument.frx":030A
         Left            =   120
         List            =   "frmDocument.frx":0320
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         ItemData        =   "frmDocument.frx":034E
         Left            =   2040
         List            =   "frmDocument.frx":035B
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton SendTypeNotes 
         Caption         =   "Notes"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   10
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton SendTypeNextel 
         Caption         =   "Nextel"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton SendTypeBoth 
         Caption         =   "Both"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Technologists:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Groups:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Urgent Issue"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Firewall Down"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Server Down"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtsendto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      MaxLength       =   700
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox txtsubject 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "FYI"
      Top             =   240
      Width           =   3375
   End
   Begin RichTextLib.RichTextBox txtmessage 
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   1191
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":037E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Recipient(s):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const EMBED_ATTACHMENT As Integer = 1454

Sub GetRecipients()
    Dim i, t, count
    Dim Marked, Notes, Nextel
    
    count = 0
    i = 0
    t = 0
    
   '------------------------------------
   
   Select Case SendType
   
        Case 1 'Notes Mail
    
            For i = 0 To 7
                
                Marked = Tech(i).Marked
                Notes = Tech(i).NotesName
                Nextel = Tech(i).Nextel
                
                
                If (count >= 1) And Marked = True Then
                    txtsendto.Text = txtsendto.Text & ", " & Notes
                    count = count + 1
                End If
    
                If (count < 1) And Marked = True Then
                    txtsendto.Text = txtsendto.Text & Notes
                    count = count + 1
                End If
            Next i
    
            
        Case 2 ' Nextel
            For i = 0 To 7
                Marked = Tech(i).Marked
                Notes = Tech(i).NotesName
                Nextel = Tech(i).Nextel
    
                If (count >= 1) And Marked = True Then
                    txtsendto.Text = txtsendto.Text & ", " & Nextel
                    count = count + 1
                End If
    
                If (count < 1) And Marked = True Then
                    txtsendto.Text = txtsendto.Text & Nextel
                    count = count + 1
                End If
            Next i
    
    Case 3 ' Both
 
            For i = 0 To 7
                Marked = Tech(i).Marked
                Notes = Tech(i).NotesName
                Nextel = Tech(i).Nextel
    
                If (count >= 1) And Marked = True Then
                    txtsendto.Text = txtsendto.Text & ", " & Notes & ", " & Nextel
                    count = count + 1
                End If
    
                If (count < 1) And Marked = True Then
                    txtsendto.Text = txtsendto.Text & Notes & ", " & Nextel
                    count = count + 1
                End If
            Next i
    End Select
'Set all Send variables back to false
        '-----------------------------------------------------
            For t = 0 To 7
                Tech(t).Marked = False
            Next t
    
    
End Sub


Private Sub btnSendMail_Click()
Dim X

'GetRecipients
'This feature is used to loop through the stored technologist data
'This feature will be updated in the next release.


'Check to see if required fields are empty
If txtsendto.Text = "" Then
    y = MsgBox("There are No Recipients typed in the field. Please select a recipient from the list.", vbInformation, "Missing Recipients")
    Exit Sub
End If

If txtsubject.Text = " " Then
    y = MsgBox("No Subject entered. Please enter a subject.", vbInformation, "Missing Subject")
    txtsubject.SetFocus
    Exit Sub
End If

If txtmessage.Text = " " Then
    y = MsgBox("No Message entered. Please enter a message.", vbInformation, "Missing Message")
    txtmessage.SetFocus
    Exit Sub
End If


'Prepare Lotus Info------------------------------------------------
'Tested and works with Notes 4.6 and Notes R5 clients on Win95/2000
    
    Dim oSess As Object
    Dim oDB As Object
    Dim oDoc As Object
    Dim oItem As Object
    Dim direct As Object
    Dim var As Variant
    Dim flag As Boolean

    MousePointer = 11
    
    'Update Status Bar
    'Form1.StatusBar1.SimpleText = "Openening Lotus Notes..."

    Set oSess = CreateObject("notes.notessession")
    Set oDB = oSess.getdatabase("", "")
    Call oDB.openmail
    
    flag = True

    If Not (oDB.isopen) Then flag = oDB.Open("", "")

    If Not flag Then
        MsgBox "Cant't open mail file: " & oDB.Server & " " & oDB.filepath
    End If
    
    'Update Status Bar
    'Form1.StatusBar1.SimpleText = "Building Message"

    Set oDoc = oDB.createdocument
    Set oItem = oDoc.createrichtextitem("BODY")
    oDoc.Form = "Memo"
    oDoc.Subject = txtsubject.Text
    oDoc.sendto = txtsendto.Text
    oDoc.body = txtmessage.Text
    oDoc.postdate = Date

    'Update Status bar
    'Form1.StatusBar1.SimpleText = "Attaching Database " & Form1.txtfilepath

    'this line here is for sending attachement remove the ' if you want to send attachements
    'Call oItem.embedobject(1454, "", Form1.txtfilepath)

    'deze staat er waarchijnlijk teveel...
    'oDoc.visable = True

    'Update Status Bar
    'Form1.StatusBar1.SimpleText = "Sending message"

    oDoc.Send False


exit_sendattachement:
        On Error Resume Next
        Set oSess = Nothing
        Set oDB = Nothing
        Set oDoc = Nothing
        Set oItem = Nothing

    'Update Status Bar
    'Form1.StatusBar1.SimpleText = "Done!"

    MousePointer = 1

Unload Me
X = MsgBox("Your Message has been sent.", vbOKOnly, "Send Sucessful")


End Sub



Private Sub Form_Load()
SendType = 1 ' Notes Mail

Form_Resize
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    'txtmessage.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 1600
    txtmessage.RightMargin = txtmessage.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Set all Send variables back to false
    For t = 0 To 20
       Tech(t).Marked = False
    Next t
End Sub

Private Sub List1_Click()
Dim Send As Boolean
Dim SelectedIndex

SelectedIndex = List1.ListIndex

Send = Tech(SelectedIndex).Marked

Select Case Send
    Case True
        Tech(SelectedIndex).Marked = False
    Case False
        Tech(SelectedIndex).Marked = True
End Select

End Sub

Private Sub List2_Click()
Dim p
p = MsgBox("This option is curently not available.", vbInformation, "Feature not available")
End Sub

Private Sub Option1_Click()
txtsubject.Text = "Server Downtime Notice"
txtmessage.Text = "Attention: The following server(s) are down:"
txtmessage.SetFocus
End Sub

Private Sub Option2_Click()
txtsubject.Text = "Firewall Downtime Notice"
txtmessage.Text = "Attention: The following firewall is currently down. Customers may experience difficulty dialing in remotely and connecting to Notes."
txtmessage.SetFocus
End Sub

Private Sub SendTypeBoth_Click(Index As Integer)
SendType = 3
End Sub

Private Sub SendTypeNextel_Click(Index As Integer)
SendType = 2
End Sub

Private Sub SendTypeNotes_Click(Index As Integer)
SendType = 1
End Sub
