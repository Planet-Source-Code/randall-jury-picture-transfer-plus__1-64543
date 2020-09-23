VERSION 5.00
Begin VB.Form TargetPath 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Delete Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6780
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2340
      TabIndex        =   4
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make New Folder"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   6300
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.DirListBox Dir1 
      Height          =   5040
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "TargetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


nametosearch = InputBox("Enter Last name to Search:", "Search By Last Name")

If nametosearch = "" Then Exit Sub
MkDir Dir1.Path & "\" & nametosearch

Dir1.Refresh
End Sub

Private Sub Command2_Click()
FormStartup.FileTarget.Path = Dir1.Path
FormStartup.LabelPathAndFilename.Caption = Dir1.Path


If FormStartup.FileTarget.ListCount - 1 > -1 Then FormStartup.FileTarget.ListIndex = 0




Unload Me

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()

On Error GoTo errhandler

RmDir (Dir1.List(Dir1.ListIndex))
Dir1.Refresh


Exit Sub
errhandler:
MsgBox ("Error Directory Not Empty")



End Sub

Private Sub Dir1_Change()
Label1.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()
Label1.Caption = Dir1.Path
Dir1.Path = FormStartup.FileTarget.Path
End Sub
