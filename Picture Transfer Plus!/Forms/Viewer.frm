VERSION 5.00
Begin VB.Form Viewer 
   BorderStyle     =   0  'None
   ClientHeight    =   7035
   ClientLeft      =   105
   ClientTop       =   -195
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll 
      Enabled         =   0   'False
      Height          =   255
      Left            =   -15
      SmallChange     =   100
      TabIndex        =   2
      Top             =   6150
      Width           =   7215
   End
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   4455
      Left            =   7200
      SmallChange     =   100
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox View 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   7155
      TabIndex        =   0
      ToolTipText     =   "Double Click Picture To Exit Full Screen"
      Top             =   255
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   30
      TabIndex        =   3
      Top             =   15
      Width           =   1215
   End
   Begin VB.Image Img 
      Height          =   255
      Left            =   7065
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ShowPicture()






Img = FormStartup.PictureScroll.Picture






'Img = LoadPicture(FormStartup.LabelPathAndFileName.Caption)



View.Picture = LoadPicture
View.PaintPicture Img.Picture, 0, 0, Img.Width, Img.Height

If Img.Width > View.Width Then
    HScroll.Min = 0
    HScroll.Max = (Img.Width - View.Width) / 1000
    HScroll.Enabled = True
Else
    HScroll.Enabled = False
End If

If Img.Height > View.Height Then
    VScroll.Min = 0
    VScroll.Max = (Img.Height - View.Height) / 1000
    VScroll.Enabled = True
Else
    VScroll.Enabled = False
End If

End Sub

Private Sub fileext_Click()
End
End Sub

Private Sub Form_Load()
Label1.Caption = "Double Click Picture To Exit Full Screen... " & FormStartup.LabelPathAndFilename.Caption
ShowPicture
End Sub

Private Sub Form_Resize()
View.Left = 0
View.Top = 250
View.Width = Me.Width - VScroll.Width - 100
If Me.Height - HScroll.Height < 200 Then
    View.Height = Me.Height
Else
    View.Height = Me.Height - HScroll.Height '- 100
End If

VScroll.Height = View.Height - 100
HScroll.Width = View.Width

VScroll.Left = View.Width + 50
HScroll.Top = View.Height + 50
Label1.Width = View.Width
'LabelPathAndFileName.Top = View.Height + 350
Call ShowPicture





End Sub

Private Sub HScroll_Change()
View.Picture = LoadPicture
View.PaintPicture Img.Picture, -HScroll.Value * 1000, -VScroll.Value * 1000, Img.Width, Img.Height
End Sub



Private Sub HScroll_Scroll()
View.Picture = LoadPicture
View.PaintPicture Img.Picture, -HScroll.Value * 1000, -VScroll.Value * 1000, Img.Width, Img.Height

End Sub

Private Sub LabelPathAndFileName_Click()
Unload Me

End Sub

Private Sub View_DblClick()
Unload Me

End Sub

Private Sub VScroll_Change()
View.Picture = LoadPicture
View.PaintPicture Img.Picture, -HScroll.Value * 1000, -VScroll.Value * 1000, Img.Width, Img.Height - 200
End Sub

Private Sub VScroll_Scroll()
View.Picture = LoadPicture
View.PaintPicture Img.Picture, -HScroll.Value * 1000, -VScroll.Value * 1000, Img.Width, Img.Height - 20

End Sub
