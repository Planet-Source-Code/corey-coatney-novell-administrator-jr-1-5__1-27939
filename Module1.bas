Attribute VB_Name = "Module1"
Declare Function WritePrivateProfileString _
Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationname As String, ByVal _
lpKeyName As Any, ByVal lsString As Any, _
ByVal lplFilename As String) As Long

Declare Function GetPrivateProfileString Lib _
"kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationname As String, ByVal _
lpKeyName As String, ByVal lpDefault As _
String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As _
String) As Long

Public fMainForm As frmMain


Global gServerName As String
Global ServerName As String
Global sFilename As String
Global ObjectShowing As Boolean

'-------Critical Server Stats Variables--------------------
Global statUptime As String
Global statOther As String
Global statCPU As String
'----------------------------------------------------------

Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String
Dim DatFilePresent As Boolean

Sub Main()
    On Error GoTo ErrTrap:
    DatFilePresent = False
    
    frmSplash.Show
    '-----------------------------------------
    Dim y
   
   'Initialize variables---------------------------------
    statCPU = "N/A"
    statOther = "N/A"
    statUptime = "N/A"
    '---------------------------------
    
    '--------------------------------------------------------
    'check to see if INI file exists
    '--------------------------------------------------------
    sFilename = App.Path + "\" + App.EXEName + ".ini"
    
    If FileExists(sFilename) Then
        'If File Exists, then Pull values
        frmSplash.lblstatus.Caption = "Status: Loading Ini file..."
        Call LoadIni
    Else
        frmSplash.lblstatus.Caption = "Status: Creating Ini file..."
        'y = MsgBox("The file " + sFilename + " has not been detected. Either this is the first time you have run the program, or your existing ini file is corrupt. Click OK and a new Ini file will be created." & sFilename, vbInformation, "First Run - " + App.EXEName + ".ini Setup")
        Call CreateIni
    End If
        
    '--------------------------------------------------------
    'check to see if DAT file exists
    '--------------------------------------------------------
   sFilename = App.Path + "\" + App.EXEName + ".dat"
    
    If FileExists(sFilename) Then
        frmSplash.lblstatus.Caption = "Status: Loading Data file..."
        'If File Exists, then Pull values
        DatFilePresent = True
            
    Else
        'y = MsgBox("The file " + sFilename + " has not been detected. Either this is the first time you have run the program, or your existing Dat file is corrupt. Click OK and a new Dat file will be created." & sFilename, vbInformation, "First Run - " + App.EXEName + ".dat Setup")
        'Update Splash screen message
        frmSplash.lblstatus.Caption = "Status: Creating Data file..."
        Call CreateDat
        DatFilePresent = True
    End If
    
'------------------------------------------
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
        
    
'-----------------------------------------------
'Get Layout Settings
'-----------------------------------------------
    
    
    KeySection = "Layout"
    KeyKey = "MainTop"
    LoadIni
    fMainForm.Top = KeyValue
    
    KeySection = "Layout"
    KeyKey = "MainLeft"
    LoadIni
    fMainForm.Left = KeyValue
    
    KeySection = "Layout"
    KeyKey = "MainWidth"
    LoadIni
    fMainForm.Width = KeyValue
    
    KeySection = "Layout"
    KeyKey = "MainHeight"
    LoadIni
    fMainForm.Height = KeyValue
    '-----------------------------------------------
    
    fMainForm.Show
    
    
    If DatFilePresent Then LoadDat
    
    Unload frmSplash
    
 Exit Sub
    
ErrTrap:
MsgBox "Error Starting the " & App.EXEName & " " & Err.Description, vbCritical + vbSystemModal, "Error"
Resume Next

End Sub
Function FileExists(sFilename As String)
    Dim Files As String
    Files = Dir(sFilename)


    If Files = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Private Sub CreateIni()
Dim FileDir As String

On Error GoTo ErrTrap:
    FileDir = App.Path + "\" + App.EXEName + ".ini"
    
    Open (FileDir) For Output As #1
        DoEvents
        Print #1, "' Netware Admin Jr."
        Print #1, "' Warning: There is no error checking on the this files variables"
        Print #1, "' If you make a change, be sure that you know what you are doing"
        Print #1, "' or the program may not function properly. If do run into issues,"
        Print #1, "' Delete this file and a new one will be created for you."
        Print #1, " "
        Print #1, " "
        Print #1, " "
        Print #1, "[Program]"
        Print #1, " "
        Print #1, "[Layout]"
        'Main Program Location on screen
        Print #1, "MainTop=1000"
        Print #1, "MainLeft=1000"
        Print #1, "MainHeight=7500"
        Print #1, "MainWidth=7470"
        Print #1, " "
        Print #1, "ObjectsTop=500"
        Print #1, "ObjectsLeft=1010"
        Print #1, "ObjectsHeight=4725"
        Print #1, "ObjectsWidth=4455"
        Print #1, " "
        
        Print #1, "[Options]"
        Print #1, "Username="
        Print #1, "Status="
        Print #1, " "
        Print #1, "[AutoUpdate]"
        Print #1, "URL=http://www.websitename/nwadminjr/updates"
        Print #1, " "
    Close #1

Exit Sub
ErrTrap:
MsgBox "Error creating the " + App.EXEName + ".ini file: " & Err.Description, vbCritical + vbSystemModal, "Error"

End Sub

Sub CreateDat()
Dim FileDir As String

On Error GoTo ErrTrap:
    
    FileDir = App.Path + "\" + App.EXEName + ".dat"
    
    Open (FileDir) For Output As #1
        DoEvents
        
        Print #1, "CINFS0001"
        Print #1, "COLFS0001"
        Print #1, "CLEFS0001"
        Print #1, "DAYFS0001"
        Print #1, "DETFS0001"
        Close #1

Exit Sub
ErrTrap:
MsgBox "Error Creating the " + App.EXEName + ".Dat file: " & Err.Description, vbCritical + vbSystemModal, "Error"
Resume

End Sub


Sub LoadDat()
Dim FileExist
Dim FileDir As String
Dim strString As String
Dim itmP As ListItem

FileDir = App.Path + "\" + App.EXEName + ".dat"

On Error GoTo ErrTrap:

FileExist = Dir(FileDir)
If FileExist = "" Then
    MsgBox "There was no server list found. A default server list will be created named" & FileDir, vbCritical + vbSystemModal, "Error"
    Call CreateDat
    FileDir = App.Path + "\" + App.EXEName + ".dat"
End If

Open (FileDir) For Input As #1
    While Not EOF(1)
        DoEvents
        Input #1, strString
        
        Set itmP = frmObjects.ListView1.ListItems.Add(, , strString)
        itmP.SubItems(1) = "Offline"
    Wend
Close #1
Exit Sub

ErrTrap:
MsgBox "Error Loading the " + App.EXEName + ".dat file: " & Err.Description & " Please delete the existing file and a new one will be created.", vbCritical + vbSystemModal, "Error"
Resume

End Sub

Sub LoadIni()
Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\" & App.EXEName & ".ini"
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If
End Sub
