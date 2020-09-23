VERSION 5.00
Begin VB.Form FormStartup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Transfer Plus!        Ver 1.0a"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14235
   Icon            =   "Startup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   14235
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboSlideSpeed 
      Height          =   315
      Left            =   10350
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13560
      Top             =   0
   End
   Begin PictureTransferPlus.isButton CommandFitImageToWindow 
      Height          =   495
      Left            =   7200
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "Startup.frx":030A
      Style           =   5
      Caption         =   "Fit Image To Window"
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Picture Transfer Plus!"
      ToolTipIcon     =   1
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBox 
      Height          =   6015
      Left            =   5520
      ScaleHeight     =   5955
      ScaleWidth      =   7950
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Double Click For Full Screen"
      Top             =   2160
      Width           =   8010
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      Left            =   5520
      SmallChange     =   200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7920
      Width           =   7590
   End
   Begin VB.VScrollBar VS 
      Height          =   5685
      Left            =   13155
      SmallChange     =   200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox PicFrame 
      Height          =   5730
      Left            =   5520
      ScaleHeight     =   5670
      ScaleWidth      =   7545
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7605
      Begin VB.PictureBox PictureScroll 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7905
         Left            =   0
         ScaleHeight     =   7905
         ScaleWidth      =   9840
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Double Click For Full Screen"
         Top             =   0
         Width           =   9840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8820
      Left            =   0
      TabIndex        =   2
      Top             =   1380
      Width           =   4575
      Begin VB.FileListBox FileSource 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   600
         TabIndex        =   0
         Top             =   2175
         Width           =   3435
      End
      Begin VB.FileListBox FileTarget 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   600
         TabIndex        =   1
         Top             =   4785
         Width           =   3435
      End
      Begin PictureTransferPlus.isButton SourcePath 
         Height          =   300
         Left            =   600
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Icon            =   "Startup.frx":0326
         Style           =   5
         Caption         =   "Path"
         iNonThemeStyle  =   0
         Tooltiptitle    =   "Picture Transfer Plus!"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PictureTransferPlus.isButton TargetPath 
         Height          =   300
         Left            =   600
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Icon            =   "Startup.frx":0342
         Style           =   5
         Caption         =   "Path"
         iNonThemeStyle  =   0
         Tooltiptitle    =   "Picture Transfer Plus!"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "FileName"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   600
         TabIndex        =   15
         Top             =   120
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Label LabelType 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1320
         Width           =   3465
      End
      Begin VB.Label LabelBytes 
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1080
         Width           =   3465
      End
      Begin VB.Label LabelSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   840
         Width           =   3450
      End
      Begin VB.Label LabelTarget 
         BackStyle       =   0  'Transparent
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2355
         TabIndex        =   5
         Top             =   4185
         Width           =   1635
      End
      Begin VB.Label LabelSource 
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2355
         TabIndex        =   4
         Top             =   1635
         Width           =   1635
      End
      Begin VB.Image ImageRemove 
         Height          =   615
         Left            =   1980
         Picture         =   "Startup.frx":035E
         Top             =   6075
         Width           =   765
      End
      Begin VB.Image ImageAdd 
         Height          =   600
         Left            =   2025
         Picture         =   "Startup.frx":1C9C
         Top             =   3480
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   675
         Left            =   600
         Picture         =   "Startup.frx":2E5E
         Top             =   6060
         Width           =   3435
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   600
         Picture         =   "Startup.frx":A790
         Top             =   3435
         Width           =   3435
      End
      Begin VB.Image Image4 
         Height          =   9240
         Left            =   0
         Picture         =   "Startup.frx":11E12
         Top             =   -15
         Width           =   4575
      End
   End
   Begin PictureTransferPlus.isButton CommandFullSizeScrollable 
      Height          =   495
      Left            =   7200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "Startup.frx":9BA74
      Style           =   5
      Caption         =   "Full Size Image Scrollable"
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Picture Transfer Plus!"
      ToolTipIcon     =   1
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PictureTransferPlus.isButton FullScreen 
      Height          =   495
      Left            =   5640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Style           =   5
      Caption         =   "Full Screen"
      IconSize        =   25
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Picture Transfer Plus!"
      ToolTipIcon     =   1
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PictureTransferPlus.isButton SlideShow 
      Height          =   495
      Left            =   8760
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "Startup.frx":9BA90
      Style           =   5
      Caption         =   "Show Slide Show"
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Picture Transfer Plus!"
      ToolTipIcon     =   1
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PictureTransferPlus.isButton StopSlideShow 
      Height          =   495
      Left            =   8760
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "Startup.frx":9BAAC
      Style           =   5
      Caption         =   "Stop Slide Show"
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Picture Transfer Plus!"
      ToolTipIcon     =   1
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PictureTransferPlus.isButton PauseSlideShow 
      Height          =   495
      Left            =   7200
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "Startup.frx":9BAC8
      Style           =   5
      Caption         =   "Pause Slide Show"
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Picture Transfer Plus!"
      ToolTipIcon     =   1
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Delay"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   10365
      TabIndex        =   24
      Top             =   1425
      Width           =   3330
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0a"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   240
      Width           =   1530
   End
   Begin VB.Image Image13 
      Height          =   660
      Left            =   -75
      Picture         =   "Startup.frx":9BAE4
      Top             =   0
      Width           =   16935
   End
   Begin VB.Image Image8 
      Height          =   7050
      Left            =   4560
      Picture         =   "Startup.frx":C0176
      Top             =   1380
      Width           =   9660
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   6
      Top             =   795
      Width           =   1065
   End
   Begin VB.Label LabelPathAndFilename 
      BackStyle       =   0  'Transparent
      Caption         =   "No File Loaded..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1215
      TabIndex        =   3
      Top             =   960
      Width           =   13035
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image3 
      Height          =   705
      Left            =   -90
      Picture         =   "Startup.frx":19DCC0
      Top             =   675
      Width           =   14355
   End
   Begin VB.Menu MnuImage 
      Caption         =   "Image"
      Begin VB.Menu MnuFitImageToWindow 
         Caption         =   "Fit Image To Window"
      End
      Begin VB.Menu MnuFullizeScrollable 
         Caption         =   "Full Size Scrollable"
      End
      Begin VB.Menu MnuFullScreen 
         Caption         =   "Full Screen"
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FormStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aa As Integer
Dim bb As Integer
Dim p As Integer
Private ImageTypes(4) As String


Dim sFileTarget As String


Private Sub Form_Load()

ComboSlideSpeed.AddItem ("Fast")
ComboSlideSpeed.AddItem ("1 Second")
ComboSlideSpeed.AddItem ("5 Second")
ComboSlideSpeed.AddItem ("15 Second")
ComboSlideSpeed.AddItem ("30 Second")
ComboSlideSpeed.AddItem ("1 Minute")
ComboSlideSpeed.ListIndex = 0


CommandFullSizeScrollable.Visible = True
CommandFitImageToWindow.Visible = False
FullScreen.Visible = True
StopSlideShow.Visible = True
StopSlideShow.Visible = False
PauseSlideShow.Visible = False
ComboSlideSpeed.Visible = True

LabelFileName = ""
FileSource.Pattern = "*.bmp;*.jpg;*.gif"
FileTarget.Pattern = "*.bmp;*.jpg;*.gif"

LabelPathAndFilename.ForeColor = vbBlue
LabelSource.ForeColor = vbBlack
LabelTarget.ForeColor = vbBlack


Label5.ForeColor = vbRed



ImageAdd.Visible = False
ImageRemove.Visible = False



'FileTarget.Path = App.Path & "\savedpics"


'Use On Error Incase Path Not There
On Error GoTo ErrorHandler   ' Enable error-handling routine.

Dim SourcePath As String
Dim FrmTargetPath As String
Open App.Path & "\Path.Ini" For Input As #1
Input #1, SourcePath
Input #1, FrmTargetPath
Close #1
FileSource.Path = SourcePath
FileTarget.Path = FrmTargetPath

If SourcePath = "" Then LabelPathAndFilename.Caption = "No Path Selected"

'make sure there is pic files
If FileSource.ListCount > 0 Then
FileSource.ListIndex = 0
LabelPathAndFilename.Caption = FileSource.Path & "\" & FileSource.List(FileSource.ListIndex)
Else
LabelPathAndFilename.Caption = FileSource.Path
End If



'LabelPathAndFileName.Caption = FileTarget.Path & "\" & FileTarget.List(FileTarget.ListIndex)
If FrmTargetPath = "" Then LabelPathAndFilename.Caption = "No Path Selected"

ErrorHandler:




End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageAdd.Visible = False
ImageRemove.Visible = False
ImageRemove.Visible = False
End Sub



Private Sub AddFile()
'MsgBox (FileSource.Path & "\" & FileSource.List(FileSource.ListIndex))
'The source file the you want to copy
SourceFile = (FileSource.Path & "\" & FileSource.List(FileSource.ListIndex))
'The destination file name
TargetFile = FileTarget.Path & "\" & FileSource.List(FileSource.ListIndex)






FileCopy SourceFile, TargetFile

FileTarget.Refresh

End Sub

Private Sub Removefile()

If FileTarget.ListIndex < 0 Then Exit Sub


Kill FileTarget.Path & "\" & FileTarget.List(FileTarget.ListIndex)

FileTarget.Refresh

PictureScroll.Picture = LoadPicture("")


If FileSource.ListIndex < 0 Then
PictureScroll.Picture = LoadPicture(FileSource.List(FileSource.ListIndex))
End If


End Sub


Private Sub CommandFitImageToWindow_Click()

If LabelFileName.Caption = "" Then Exit Sub
CommandFullSizeScrollable.Visible = True
CommandFitImageToWindow.Visible = False
PicFrame.Visible = False
VS.Visible = False
HS.Visible = False
picBox.Visible = True
ResizeImage

End Sub

Private Sub CommandFullSizeScrollable_Click()
If LabelFileName.Caption = "" Then Exit Sub
CommandFullSizeScrollable.Visible = False
CommandFitImageToWindow.Visible = True
PicFrame.Visible = True
VS.Visible = True
HS.Visible = True
picBox.Visible = False

End Sub

Private Sub FileSource_Click()



'Show First File If All Deleted Show Blank
If FileSource.ListCount < 0 Then
HS.Value = 0
VS.Value = 0


End If


HS.Value = 0
VS.Value = 0



FileSource.ForeColor = vbBlue
FileTarget.ForeColor = vbBlack
LabelSource.ForeColor = vbRed
LabelTarget.ForeColor = vbBlack
Label5.Caption = "Source"



PictureScroll.Picture = LoadPicture(FileSource.Path & "\" & FileSource.List(FileSource.ListIndex))

SizeImage

LabelPathAndFilename.Caption = FileSource.Path & "\" & FileSource.List(FileSource.ListIndex)
LabelFileName.Caption = "Filename:" & Chr$(10) & FileSource.List(FileSource.ListIndex)
Imageinfo

ResizeImage

End Sub

Public Sub SizeImage()
If PictureScroll.Width > PicFrame.Width Then
    HS.Visible = True
    'PicFrame.Height = PicFrame.Height - HS.Height
Else
    HS.Visible = False
End If
If PictureScroll.Height > PicFrame.Height Then
    VS.Visible = True
    'PicFrame.Width = PicFrame.Width - VS.Width
Else
    
    VS.Visible = False
End If
'HS.Width = PicFrame.Width
'VS.Height = PicFrame.Height
HS.Max = PictureScroll.Width - HS.Width
VS.Max = PictureScroll.Height - VS.Height
End Sub

Private Sub Imageinfo()
    ImageTypes(0) = "Unknown"
    ImageTypes(1) = "GIF"
    ImageTypes(2) = "JPEG"
    ImageTypes(3) = "PNG"
    ImageTypes(4) = "BMP"
    ReadImageInfo (LabelPathAndFilename.Caption)
    a = ImageWidth
    
  
    LabelSize.Caption = "Dimensions: " & a & "w X " & ImageHeight & "h"
    LabelType = "Type: " & ImageTypes(ImageType)
    LabelBytes.Caption = "Size: " & FileSize & " (Bytes)"

End Sub




Private Sub FileSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageAdd.Visible = False
ImageRemove.Visible = False




End Sub

Private Sub FileTarget_Click()


'Show First File If All Deleted Show Blank
If FileSource.ListCount < 0 Then
HS.Value = 0
VS.Value = 0


End If


HS.Value = 0
VS.Value = 0



FileSource.ForeColor = vbBlack
FileTarget.ForeColor = vbBlue
LabelSource.ForeColor = vbBlack
LabelTarget.ForeColor = vbRed
Label5.Caption = "Target"



PictureScroll.Picture = LoadPicture(FileTarget.Path & "\" & FileTarget.List(FileTarget.ListIndex))

SizeImage

LabelPathAndFilename.Caption = FileTarget.Path & "\" & FileTarget.List(FileTarget.ListIndex)
LabelFileName.Caption = "Filename:" & Chr$(10) & FileTarget.List(FileTarget.ListIndex)
'FileTarget.Refresh
Imageinfo

ResizeImage
End Sub

Private Sub FileTarget_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageAdd.Visible = False
ImageRemove.Visible = False




End Sub

Private Sub Form_Activate()
ResizeImage

End Sub


Private Sub Form_Unload(Cancel As Integer)
Open App.Path & "\Path.Ini" For Output As #1
Write #1, FileSource.Path
Write #1, FileTarget.Path
Close #1




End Sub

Private Sub HS_Change()
PictureScroll.Left = -HS.Value

End Sub

Private Sub HS_Scroll()
PictureScroll.Left = -HS.Value

End Sub

Private Sub Image1_Click()
If FileSource.ListIndex < 0 Then Exit Sub

AddFile
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FileSource.ListIndex < 0 Then Exit Sub
ImageAdd.Visible = True
End Sub






Private Sub Slider1_Change()
LabelSlideSpeed.Caption = Slider1.Value
End Sub

Private Sub PauseSlideShow_Click()

Select Case PauseSlideShow.Caption

Case "Pause Slide Show"
PauseSlideShow.Caption = "Resume Slide Show"
Timer1.Enabled = False

Case "Resume Slide Show"
PauseSlideShow.Caption = "Resume Slide Show"
Timer1.Enabled = True
PauseSlideShow.Caption = "Pause Slide Show"
End Select
End Sub

Private Sub SlideShow_Click()
If LabelFileName.Caption = "" Then Exit Sub

CommandFitImageToWindow_Click

CommandFitImageToWindow.Visible = False
CommandFullSizeScrollable.Visible = False
FullScreen.Visible = False
SlideShow.Visible = False
StopSlideShow.Visible = True
PauseSlideShow.Visible = True
ComboSlideSpeed.Visible = False

If LabelSource.ForeColor = vbRed And FileSource.ListCount = 0 Then Exit Sub
If LabelTarget.ForeColor = vbRed And FileTarget.ListCount = 0 Then Exit Sub

Timer1.Enabled = False
p = 0
aa = 0

If LabelSource.ForeColor = vbRed Then
FileSource.ListIndex = 0
Label2.Caption = "Slide Show Picture " & FileSource.ListIndex + 1 & " of " & FileSource.ListCount
Timer1.Enabled = True
End If

If LabelTarget.ForeColor = vbRed Then
FileTarget.ListIndex = 0
Label2.Caption = "Slide Show Picture " & FileTarget.ListIndex + 1 & " of " & FileTarget.ListCount
Timer1.Enabled = True
End If



End Sub

Private Sub SourcePath_Click()

'RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project." + vbCrLf + _
"sfile = BrowseForFolder(" + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")"
sfile = BrowseForFolder(FileSource.Path, "RDEAN")
'If sfile <> "" Then MsgBox "You selected " + sfile



If sfile = "" Then Exit Sub

FileSource.Path = sfile

'Open App.Path & "\Path.Ini" For Output As #1
'Write #1, FileSource.Path
'Close #1

'make sure there is pic files
If FileSource.ListCount > 0 Then
FileSource.ListIndex = 0
LabelPathAndFilename.Caption = FileSource.Path & "\" & FileSource.List(FileSource.ListIndex)
Else
LabelSource.ForeColor = vbBlack





LabelPathAndFilename.Caption = FileSource.Path
End If



End Sub

Private Sub StopSlideShow_Click()
Timer1.Enabled = False
CommandFullSizeScrollable.Visible = True
CommandFitImageToWindow.Visible = False
FullScreen.Visible = True

StopSlideShow.Visible = False
SlideShow.Visible = True
PauseSlideShow.Visible = False
ComboSlideSpeed.Visible = True
PauseSlideShow.Caption = "Pause Slide Show"
End Sub

Private Sub TargetPath_Click()

picBox.Picture = LoadPicture("")
PictureScroll.Picture = LoadPicture("")
picBox.Refresh


FrmTargetPath.Show

End Sub

Private Sub FullScreen_Click()
If LabelFileName.Caption = "" Then Exit Sub


Viewer.Show
'End If

End Sub

Private Sub Image2_Click()
If FileTarget.ListIndex < 0 Then Exit Sub
Removefile
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FileTarget.ListIndex < 0 Then Exit Sub
ImageRemove.Visible = True
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageAdd.Visible = False
ImageRemove.Visible = False





End Sub








Private Sub ImageAdd_Click()
If FileSource.ListIndex < 0 Then Exit Sub
For xx = 0 To FileTarget.ListCount - 1
If FileTarget.List(xx) = FileSource.List(FileSource.ListIndex) Then
MsgBox ("Duplicate Filename" & Chr(10) & FileTarget.List(xx) & Chr(10) & "Can Not Copy"), vbCritical
'MsgBox ("Duplicate Filename" & Chr(10) & "Can Not Copy")
Exit Sub
End If


Next xx



AddFile
End Sub



Private Sub ImageRemove_Click()
If FileTarget.ListIndex < 0 Then Exit Sub
Removefile


'Show First File If All Deleted Show Blank
If FileTarget.ListCount > 0 Then
FileTarget.ListIndex = 0
Else
BlankShow
LabelPathAndFilename.Caption = ""
End If

FileSource.Refresh


End Sub

Private Sub ImageRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FileTarget.ListIndex < 0 Then Exit Sub
ImageRemove.Visible = True
End Sub



Private Sub MnuFitImageToWindow_Click()
CommandFitImageToWindow_Click
End Sub

Private Sub MnuFullizeScrollable_Click()
CommandFullSizeScrollable_Click
End Sub

Private Sub MnuFullScreen_Click()
FullScreen_Click
End Sub

Private Sub picBox_DblClick()
If LabelSource.ForeColor = vbBlack And LabelTarget.ForeColor = vbBlack Then Exit Sub
Viewer.Show

End Sub


Private Sub BlankShow()
'Viewer.Show
PictureScroll.ScaleHeight = 7905
PictureScroll.ScaleWidth = 9840
HS.Value = 0
VS.Value = 0



End Sub

Private Sub PictureScroll_DblClick()
If LabelSource.ForeColor = vbBlack And LabelTarget.ForeColor = vbBlack Then Exit Sub
Viewer.Show

End Sub



Private Sub txtType_Change()

End Sub

Private Sub Timer1_Timer()


Select Case ComboSlideSpeed.ListIndex


Case 0
Timer1.Interval = 100
bb = 1
Case 1
Timer1.Interval = 1000
    bb = 1
Case 2
Timer1.Interval = 1000
bb = 5
Case 3
Timer1.Interval = 1000
bb = 15
Case 4
Timer1.Interval = 1000
bb = 30
Case 5
Timer1.Interval = 1000
bb = 60
End Select




aa = aa + 1
p = aa


If LabelSource.ForeColor = vbRed And FileSource.ListIndex = FileSource.ListCount - 1 Then
Timer1.Enabled = False
'FileSource.ListIndex = 0
Label2.Caption = "Picture Delay"
CommandFullSizeScrollable.Visible = True
CommandFitImageToWindow.Visible = False
FullScreen.Visible = True
StopSlideShow.Visible = False
SlideShow.Visible = True
PauseSlideShow.Visible = False
ComboSlideSpeed.Visible = True

Exit Sub
End If

If LabelTarget.ForeColor = vbRed And FileTarget.ListIndex = FileTarget.ListCount - 1 Then
Timer1.Enabled = False
'Filetarget.ListIndex = 0
Label2.Caption = "Picture Delay"
CommandFullSizeScrollable.Visible = True
CommandFitImageToWindow.Visible = False
FullScreen.Visible = True
StopSlideShow.Visible = False
SlideShow.Visible = True
PauseSlideShow.Visible = False
ComboSlideSpeed.Visible = True

Exit Sub
End If




If p = bb And LabelSource.ForeColor = vbRed Then
FileSource.ListIndex = FileSource.ListIndex + 1
PictureScroll.Picture = LoadPicture(FileSource.Path & "\" & FileSource.List(FileSource.ListIndex))
aa = 0
p = 0
Label2.Caption = "Slide Show Picture " & FileSource.ListIndex + 1 & " of " & FileSource.ListCount
End If


If p = bb And LabelTarget.ForeColor = vbRed Then
FileTarget.ListIndex = FileTarget.ListIndex + 1
PictureScroll.Picture = LoadPicture(FileTarget.Path & "\" & FileTarget.List(FileTarget.ListIndex))
aa = 0
p = 0
'FileTarget.ListIndex = 0
Label2.Caption = "Slide Show Picture " & FileTarget.ListIndex + 1 & " of " & FileTarget.ListCount
End If



End Sub

Private Sub VS_Change()
PictureScroll.Top = -VS.Value

End Sub

Private Sub VS_Scroll()
PictureScroll.Top = -VS.Value

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// This program resizes the image's width and height based on the
'/// destination's width and height while maintaining aspect ratio.
'/// The image's aspect ratio is defined as : Aspect Ratio = Image's Height / Image's Width
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// NOTE : No error-handling included
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub ResizeImage()
      Dim ImageWidth As Single      'Original image's width
      Dim ImageHeight As Single     'Original image's height
      Dim ResizedWidth As Single     'Resized image's width
      Dim ResizedHeight As Single    'Resized image's height
      Dim DestWidth As Single         'Destination picturebox's width
      Dim DestHeight As Single         'Destination picturebox's height
      Dim AspectRatio As Single        'Image's aspect ratio ( NOTE : aspect ratio = height / width )
      
      'Destination picturebox's dimensions
      DestWidth = picBox.Width
      DestHeight = picBox.Height
      
      'Stores the image's original dimensions
      ImageWidth = PictureScroll.Width
      ImageHeight = PictureScroll.Height
      
      'Initializes the resized dimensions
      ResizedWidth = ImageWidth
      ResizedHeight = ImageHeight
                  
      'Calculate image's original aspect ratio and display it in lblOldAspectRatio
      AspectRatio = (ImageHeight / ImageWidth)
      lblOldAspectRatio = "Original Aspect Ratio : " & AspectRatio
      
      'Now resize the dimensions...
      Call AdjustImageDimensions(ResizedWidth, ResizedHeight, DestWidth, DestHeight)
      
      'Calculate image's new aspect ratio and display it in lblNewAspectRatio
      AspectRatio = (ResizedHeight / ResizedWidth)
      lblNewAspectRatio = "New Aspect Ratio : " & AspectRatio
      
      'Paint the image onto picBox
      picBox.Cls
      On Error Resume Next
      picBox.PaintPicture LoadPicture(LabelPathAndFilename.Caption), 0, 0, ResizedWidth, ResizedHeight
                                
End Sub
