VERSION 5.00
Begin VB.Form Frm_Copy 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "x copy"
   ClientHeight    =   3750
   ClientLeft      =   735
   ClientTop       =   1860
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "XCOPY.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3750
   ScaleWidth      =   8085
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   5880
      TabIndex        =   23
      Top             =   2760
      Width           =   1455
   End
   Begin VB.FileListBox File2 
      Height          =   720
      Left            =   3000
      TabIndex        =   19
      Top             =   2400
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   720
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   2655
   End
   Begin VB.DirListBox Dir2 
      Height          =   810
      Left            =   3000
      TabIndex        =   15
      Top             =   600
      Width           =   2655
   End
   Begin VB.DriveListBox Drive2 
      Height          =   330
      Left            =   3000
      TabIndex        =   14
      Top             =   240
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5760
      TabIndex        =   10
      Top             =   0
      Width           =   2295
      Begin VB.CheckBox Check4 
         Caption         =   "Make Specified Dir for To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkRecourse 
         Caption         =   "Include Subdirs"
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
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Matching"
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
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   810
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtPattern 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtTo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtFrom 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton cmdCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Copy"
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
      Left            =   5880
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Files in directory  To:"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Files in directory  From:"
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
      TabIndex        =   20
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3000
      TabIndex        =   17
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&From:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Copy Status:"
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
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Copy files matching:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5760
      TabIndex        =   6
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3000
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&From:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Lbl_CopyInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   3360
      Width           =   6495
   End
End
Attribute VB_Name = "Frm_Copy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
s% = MsgBox("Warning using this option will copy all files with the specified name or extension from all directories on system. Do you want to enable this feature ?", 3 + 32 + 256)
  If s% = 2 Or s% = 7 Then
   Check1.Value = 0
   Exit Sub
  End If
txtPattern.Enabled = True
Label3.Enabled = True
txtPattern.SetFocus
Else
txtPattern.Enabled = False
Label3.Enabled = False
End If
End Sub


Private Sub Check4_Click()
If Check4.Value = 1 Then
Drive2.Enabled = False
Dir2.Enabled = False
txtTo.Enabled = True
Exit Sub
End If
txtTo = Dir2
txtTo.Enabled = False
End Sub

Private Sub cmdCopy_Click()
 cmdCopy.Enabled = False
   R = CopyFiles(CStr(txtFrom), CStr(txtTo), CInt(chkRecourse), CStr(txtPattern))

   If Not R Then
   Else
      Drive1.SetFocus
      cmdCopy.Enabled = True
      Lbl_CopyInfo.Caption = "Done"
   End If
End Sub


Private Sub cmdExit_Click()
Unload Me
End
End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
txtFrom = Dir1
End Sub

Private Sub Dir2_Change()
File2.Path = Dir2.Path
txtTo = Dir2
End Sub

Private Sub Drive1_Change()
 Dir1.Path = Drive1.Drive
End Sub


Private Sub Drive2_Change()
Dir2.Path = Drive2.Drive
End Sub

Private Sub Form_Load()
Show
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Drive1.SetFocus
Drive1.Drive = App.Path
Dir1.Path = App.Path
Drive2.Drive = App.Path
Dir2.Path = App.Path
End Sub


Private Sub txtFrom_GotFocus()
txtFrom = Dir1
End Sub


Private Sub txtTo_GotFocus()
If Check4.Value = 1 Then Exit Sub
txtTo = Dir2
End Sub


