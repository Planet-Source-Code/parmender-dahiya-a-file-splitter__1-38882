VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOP SPLIT"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3240
      Picture         =   "frmmain.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   16
      Top             =   0
      Width           =   525
   End
   Begin VB.CommandButton cmdhelp 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Picture         =   "frmmain.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "How to use TOP SPLIT"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdabout 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      Picture         =   "frmmain.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "About TOP SPLIT"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdexit 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Picture         =   "frmmain.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit TOP SPLIT"
      Top             =   2760
      Width           =   735
   End
   Begin MSComDlg.CommonDialog opnsav 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progress 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Splitted status"
      Top             =   4080
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmddoit 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      Picture         =   "frmmain.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "SPLIT THE FILE"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox splnum 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   375
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   2
      ToolTipText     =   "Click to enable"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox splsize 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   375
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   1
      ToolTipText     =   "Enter size of one part.Double Click for floppy size."
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox savefile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "Enter name and path of splitted file for saving"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE AS..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Save splitted file as"
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "OPEN..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Open file to split"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox openfile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Enter name with path of file to split"
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      ToolTipText     =   "Source File size in bytes"
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Presented by PARMENDER DAHIYA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2040
      TabIndex        =   18
      ToolTipText     =   "Designer"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Splitted X of Y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   17
      ToolTipText     =   "Currently processing"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "KB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   885
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "BY NUMBER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   915
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "BY SIZE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   240
      TabIndex        =   10
      Top             =   910
      Width           =   735
   End
   Begin VB.Label current 
      Alignment       =   2  'Center
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      ToolTipText     =   "Currently processed part X of Y"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "TOP SPLIT"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private total As Double
Private onepart, srclength As Double
Private display, splname, batapp As String

Private Sub cmdabout_Click()
    MsgBox "TOP SPLIT - Version 1.0" & vbCrLf & _
    "Developed by PARMENDER DAHIYA" & vbCrLf & _
    "E mail :- ps_dahiya@yahoo.com" & vbCrLf & _
    "This is a beta version." & vbCrLf & _
    "No resposibility of author in case of malfunction or data loss." & vbCrLf & _
    "For bug reporting feel free to mail.", vbOKOnly + vbInformation, "TOP SPLIT - About"
End Sub

Private Sub cmddoit_Click()
Dim temp As Double
On Error GoTo splerr3

If splsize.Locked = False Then splsize.BackColor = &H80000005
If splnum.Locked = False Then splnum.BackColor = &H80000005
openfile.BackColor = &H80000005
savefile.BackColor = &H80000005

'Check that value in the box for size is correct numeric value
'It is checked only when splitting by size
If splsize.Locked = False Then
    If splsize.Text = "" Then     'When nothing entered
        MsgBox "Enter size of one split file.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
        splsize.BackColor = &HFFFF&
        splsize.SetFocus
        Exit Sub
    Else
        If IsNumeric(splsize.Text) = False Or chkint(splsize.Text) = False Then 'When other than numeric entered
            MsgBox "Only integer values allowed for split size.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
            splsize.BackColor = &HFFFF&
            splsize.SetFocus
            Exit Sub
        Else
            temp = splsize.Text
            If temp = 0 Then ' When zeros entered
                MsgBox "Split size can not be zero.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
                splsize.BackColor = &HFFFF&
                splsize.SetFocus
                Exit Sub
            Else ' Correct value entered
            onepart = splsize.Text * 1024
            End If
        End If
    End If
End If

'Check that value in the box for number is correct numeric value
'It is checked only when splitting by number

If splnum.Locked = False Then
    If splnum.Text = "" Then 'When nothing entered
        MsgBox "Enter number to split file into.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
        splnum.BackColor = &HFFFF&
        splnum.SetFocus
        Exit Sub
    Else
        If IsNumeric(splnum.Text) = False Or chkint(splnum.Text) = False Then
                    'When other than integer entered
            MsgBox "Only integer values allowed for number of split files.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
            splnum.BackColor = &HFFFF&
            splnum.SetFocus
            Exit Sub
        Else
            temp = splnum.Text
            If temp = 0 Or temp = 1 Then ' When zeros or one entered
                MsgBox "File can not be split into zero or one part.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
                splnum.BackColor = &HFFFF&
                splnum.SetFocus
                Exit Sub
            End If
        End If
    End If
End If

If openfile.Text = "" Then ' File name not entered for splitting
    MsgBox "Enter the name of the file to split.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
    openfile.BackColor = &HFFFF&
    openfile.SetFocus
    Exit Sub
End If

If savefile.Text = "" Then 'File name not entered for splitted file
    MsgBox "Enter the name of the file to save the split file.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
    savefile.BackColor = &HFFFF&
    savefile.SetFocus
    Exit Sub
End If

srclength = FileLen(openfile.Text) ' Get size of source file in bytes
Label8.Caption = Format(FileLen(openfile.Text) / 1024, "##,##0.000") & " KBytes"

'If size of source file is less than what we specified in "SPLIT BY SIZE"
If splsize.Locked = False And srclength <= onepart Then
    MsgBox "Size of source file is " & Format(srclength \ 1024, "##,##0.000") & " KBytes" & vbCrLf & _
    "Size of one part specified for splitting is " & splsize.Text & " KBytes" & vbCrLf & _
    "SPLITTING CAN NOT BE DONE.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
    splsize.BackColor = &HFFFF&
    splsize.SetFocus
    openfile.BackColor = &HFFFF&
    Exit Sub
End If


'Calculate the total number of parts to be splitted
'It is required only when size specified
If splsize.Locked = False Then
    If (srclength / (splsize.Text * 1024)) > (srclength \ (splsize.Text * 1024)) Then
        total = (srclength \ (splsize.Text * 1024)) + 1
    Else
        total = (srclength \ (splsize.Text * 1024))
    End If
End If


'Calculate size of one part of splitted file
'Need to calculate only when we specify the number instead of size
If splnum.Locked = False Then
    temp = splnum.Text
    onepart = srclength / splnum.Text
    total = splnum.Text
    lastmsg = "Last part will be of " & Format(onepart / 1024, "##,##0.000")
Else
    lastmsg = "Last part will be of " & Format((srclength - (onepart * (total - 1))) / 1024, "##,##0.000")
    If (srclength - (onepart * (total - 1))) / 1024 = 0 Then
        lastmsg = "Last part will be of " & Format(onepart / 1024, "##,##0.000")
    End If
End If

'Before splitting check whether the number of parts exceeds 200 or not??
If total > 200 Then
MsgBox "A file can not be splitted into more than 200 parts." & vbCrLf & _
    "SPLITTING CAN NOT BE DONE.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
    splnum.BackColor = &HFFFF&
    splnum.SetFocus
Exit Sub
End If

'Check that size of one oart should not be less than 100 bytes.

If onepart < 100 Then
MsgBox "A file can not be split in parts less than 100 bytes." & vbCrLf & _
    "SPLITTING CAN NOT BE DONE.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
Exit Sub
End If

'Display the total parts, their size and Confirm splitting
If (MsgBox("File will be split in " & total & " parts." & vbCrLf & _
    "All parts (except last) will be of " & Format(onepart / 1024, "##,##0.000") & " KBytes" & vbCrLf & _
    lastmsg & " KBytes" & vbCrLf & _
    "" & vbCrLf & _
    "PROCEED FOR SPLITTING ?", vbYesNo + vbQuestion, _
    "TOP SPLIT - Split Confirm")) = vbYes Then
    
    Call split 'If yes then split
    
    MsgBox "Specified file splitted in " & total & " parts." & vbCrLf & _
    "Batch file '" & splname & ".bat" & "' created." & vbCrLf & _
    "To join all the parts just double click on the batch file.", vbInformation + vbOKOnly, "TOP SPLIT - File Splitted"
    If (MsgBox("To append the just created batch file with the option:" & vbCrLf & _
    "'DELETE ALL THE PARTS AFTER JOINING ?' Click YES.", vbYesNo + vbQuestion, _
        "TOP SPLIT - Option")) = vbYes Then
        Call appendbatch
    End If
    'Now clear all the text boxes
    Call splsize_Click
    splsize.SetFocus
    splsize.Text = "1380"
    openfile.Text = ""
    savefile.Text = ""
    progress.Value = 0
    current.Caption = "0/0"
    i = 0
End If
Exit Sub

splerr3:
MsgBox "File '" & openfile.Text & "' not found or some unexpected error.", vbOKOnly + vbCritical, "TOP SPLIT - Error"
openfile.BackColor = &HFFFF&
openfile.SetFocus
End Sub


Private Sub cmdexit_Click()
If (MsgBox("Are you sure to close TOP SPLIT ?", vbYesNo + vbQuestion, "TOP SPLIT - Close Confirm")) = vbYes Then End
End Sub

Private Sub cmdhelp_Click()
frmmain.Hide
frmhelp.Show
End Sub

Private Sub cmdopen_Click()
On Error Resume Next
opnsav.DialogTitle = "TOP SPLIT - Select file to split..."
Label8.Caption = "No File Selected"
opnsav.Action = 1 'Prompt for file opening
openfile.Text = opnsav.FileName
Label8.Caption = Format(FileLen(openfile.Text) / 1024, "##,##0.000") & " KBytes"
opnsav.FileName = ""
openfile.BackColor = &H80000005
End Sub

Private Sub cmdsave_Click()
opnsav.DialogTitle = "TOP SPLIT - Save splitted files as..."
opnsav.Action = 2 'Prompt for file saving
savefile.Text = opnsav.FileName
opnsav.FileName = ""
openfile.BackColor = &H80000005
End Sub

Private Sub Form_Load()
splnum.Locked = True
splnum.BackColor = &H8000000F
current.Caption = "0/0"
splsize.Text = 1380
Label8.Caption = "No File Selected"
display = "Presented by PARMENDER DAHIYA"
App.TaskVisible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If (MsgBox("Are you sure to close TOP SPLIT ?", vbYesNo + vbQuestion, "TOP SPLIT - Close Confirm")) = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub splnum_Click()
'Enable the "BY NUMBER" text box and disable the "BY SIZE" text box
splsize.Text = ""
splsize.ToolTipText = "Click to enable"
splsize.Locked = True
splsize.BackColor = &H8000000F
splnum.Locked = False
splnum.BackColor = &H80000005
splnum.ToolTipText = "Enter the number of parts to split in"
End Sub

Private Sub splsize_Click()
'Disable the "BY NUMBER" text box and enable the "BY SIZE" text box
splnum.Text = ""
splnum.ToolTipText = "Click to enable"
splnum.Locked = True
splnum.BackColor = &H8000000F
splsize.Locked = False
splsize.BackColor = &H80000005
splsize.ToolTipText = "Enter size of one part.Double Click for floppy size."
End Sub

Private Function copy(src As String, dst As String, top As Long, bottom As Long)
On Error GoTo splerr2
Dim b(64000) As Byte
Dim c() As Byte

'This function reads from source file and then writes to the part specified

Open src For Binary As #1

test = ((bottom - top) Mod 64000) - ((bottom - top) \ 64000)

'Write to the part specified
ReDim c(test)

Open dst For Binary As #2
Seek #1, top
For i = 0 To ((bottom - top) \ 64000) - 1
Get #1, , b
Put #2, , b
Next i

'This will be needed when ((bottom - top) \ 64000) will not be a whole number
If ((bottom - top) Mod 64000) > 0 Then
Get #1, , c
Put #2, , c
End If
Close #1
Close #2

Exit Function

splerr2:
MsgBox "Unexpected Error while splitting. Application will close.", vbOKOnly + vbCritical, "TOP SPLIT - Severe Error"
End
End Function

Private Function split()
On Error GoTo splerr1
Static i As Integer
Dim last As Long
i = 0
cmdstr = "copy /b "

'Get the file name without path for the file "Save As"
Call getfilename(savefile.Text)
batapp = splname

Do
    current.Caption = i + 1 & "/" & total
    progress.Value = ((i + 1) / total) * 100
    'Manipulate the string to write to the batch file
    cmdstr = cmdstr + splname + "." + Trim(Str$(i)) + " + "
    'Create split part
    copy openfile.Text, splname + "." + Trim(Str$(i)), i * onepart + 1, (i + 1) * onepart
    i = i + 1
    DoEvents
Loop Until i >= total - 1

'Create last part if needed. This will be needed when srclength\onepart
'will not be a whole number


current.Caption = i + 1 & "/" & total
cmdstr = cmdstr + splname + "." + Trim(Str$(i)) + " "
'Create last split part
last = srclength
copy openfile.Text, splname + "." + Trim(Str$(i)), i * onepart + 1, last
progress.Value = ((i + 1) / total) * 100

'Write name of the batch file to join the splitted files
Call getfilename(openfile.Text)
cmdstr = cmdstr + splname
Open splname + ".bat" For Output As #1
Print #1, cmdstr
Close #1
Exit Function
splerr1:
MsgBox "Unexpected Error while splitting. Application will close.", vbOKOnly + vbCritical, "TOP SPLIT - Severe Error"
End

End Function
Function getfilename(tempname)
Dim length As Long
Dim i As Integer

splname = ""
length = Len(tempname)
i = 0

'This function gets the file name without its path
'For example if file name is "C:\Winnt\abcd.exe" then
'this function will retrieve only abcd.exe i.e. string right to the last "\"
'Also the embedded blanks in the name will be replaced by an underscore
Do
    i = i + 1
    If Mid(tempname, i, 1) = "\" Then
        splname = ""
    Else
        If (Mid(tempname, i, 1) = " ") Then
            splname = splname & "_"
        Else
            splname = splname & Mid(tempname, i, 1)
        End If
    End If
    DoEvents
Loop Until i >= length

End Function

Function appendbatch()

'This function appends (if user tell to say so) the batch file
'with the following entries:
'del part1
'del part2
'del part3
' and so on

Open splname + ".bat" For Append As #1
For i = 0 To (total - 1)
    Print #1, "del " + batapp + "." + Trim(Str$(i))
Next i
    Print #1, "del " + splname & ".bat"

Close #1
'MsgBox "Batch File appended with the 'DELETE PARTS AFTER JOINING' option." & vbCrLf & _
"While joining the files you will be prompted to delete files one by one.", vbOKOnly + vbInformation, "TOP SPLIT - Over"
End Function

Private Sub splsize_DblClick()
splsize.Text = 1380
End Sub

Private Sub Timer1_Timer()
Static s As Integer
    s = s + 1
    Timer1.Interval = 100
    Label6.Caption = Mid(display, s, 30)
    
    If s = 30 Then
        Label6.Caption = display
        Timer1.Interval = 1500
        s = 0
    End If

End Sub

Function chkint(chkinttemp As String) As Boolean

' to check whether the entered value is integer or not
For i = 1 To Len(chkinttemp)
If Mid(chkinttemp, i, 1) = "." Then
chkint = False
Exit Function
End If

Next i
chkint = True
End Function
