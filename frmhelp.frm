VERSION 5.00
Begin VB.Form frmhelp 
   Caption         =   "TOP SPLIT - HELP"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "frmhelp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5655
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2040
      Picture         =   "frmhelp.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2160
      Picture         =   "frmhelp.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   10
      Top             =   0
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   3480
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   5175
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "SHORTCUTS : If you select a shortcut to split then its target will be selected for splitting. So be aware of that."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   6015
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   520
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   4010
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Before joining the files, make sure that all the parts should be in the same path where batch file is there."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   3960
      Width           =   5295
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3410
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   2430
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1240
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   $"frmhelp.frx":0CC6
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label7 
      Caption         =   $"frmhelp.frx":0D97
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label Label6 
      Caption         =   $"frmhelp.frx":0E2A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "SPLIT INSTRUCTIONS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "User can split in two ways. 1. By size (in KBytes)         2. By number (Integer)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "JOIN INSTRUCTIONS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "See the  batch file with the name as that of original file. To Join just double click on it. This will create the original file."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Developed by : PARMENDER DAHIYA Email:- ps_dahiya@yahoo.com"
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
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmmain.Show
Unload Me
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Timer1_Timer()
If Label1.Visible = True Then
Timer1.Interval = 500
Label1.Visible = False
Exit Sub
End If
If Label1.Visible = False Then
Timer1.Interval = 1500
Label1.Visible = True
Exit Sub
End If
End Sub
