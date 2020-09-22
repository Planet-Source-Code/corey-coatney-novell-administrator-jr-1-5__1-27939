VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   3735
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5220
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2640
         Left            =   120
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   2610
         ScaleWidth      =   4950
         TabIndex        =   1
         Top             =   120
         Width           =   4980
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   2880
         Width           =   1890
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Tag             =   "Warning"
         Top             =   3360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.major & "." & App.minor & "." & App.revision
    
    
End Sub

