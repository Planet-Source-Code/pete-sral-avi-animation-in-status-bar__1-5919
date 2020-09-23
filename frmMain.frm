VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMain 
   Caption         =   "Visit Pete's Programmer's AVI Files for more AVI files"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7275
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Animation Type..."
      Height          =   1575
      Left            =   1830
      TabIndex        =   8
      Top             =   1365
      Width           =   2415
      Begin ComCtl2.Animation Animation2 
         Height          =   435
         Left            =   945
         TabIndex        =   13
         Top             =   270
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         _Version        =   327681
         AutoPlay        =   -1  'True
         Center          =   -1  'True
         FullWidth       =   30
         FullHeight      =   29
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Alert"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Folder Update"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Find File"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   900
      End
      Begin ComCtl2.Animation Animation3 
         Height          =   435
         Left            =   1680
         TabIndex        =   14
         Top             =   630
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   327681
         AutoPlay        =   -1  'True
         Center          =   -1  'True
         FullWidth       =   33
         FullHeight      =   29
      End
      Begin ComCtl2.Animation Animation4 
         Height          =   435
         Left            =   1230
         TabIndex        =   15
         Top             =   1020
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         _Version        =   327681
         AutoPlay        =   -1  'True
         Center          =   -1  'True
         FullWidth       =   30
         FullHeight      =   29
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   975
      ScaleWidth      =   7275
      TabIndex        =   7
      ToolTipText     =   "Download more AVI files!!!"
      Top             =   0
      Width           =   7275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Avi in Status Bar..."
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
      Begin VB.OptionButton Option4 
         Caption         =   "Panel 4"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Panel 3"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Panel 2"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Panel 1"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   615
      Left            =   6045
      TabIndex        =   1
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      _Version        =   327681
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   65
      FullHeight      =   41
   End
   Begin ComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3435
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   882
            MinWidth        =   882
            Text            =   ":-)"
            TextSave        =   ":-)"
            Key             =   "Message"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9287
            Text            =   "Visit Pete's Programmer's AVI Files for more AVI files"
            TextSave        =   "Visit Pete's Programmer's AVI Files for more AVI files"
            Key             =   "Output"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1217
            MinWidth        =   176
            TextSave        =   "7:29 PM"
            Key             =   "Time"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmMain.frx":2A38
            TextSave        =   ""
            Key             =   "Anim"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "http://pjs-inc.com/vb-avi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iCurPanel As Integer


'note
'this code was inspired by James E. Toebes
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=2552
'I only take credit for using an AVI instead of the ProgressBar.

Private Sub Form_Load()

Animation1.Open App.Path & "\AVI\Alert.avi"
SetAviStatusBar 4

Animation2.Open App.Path & "\AVI\Alert.avi"
Animation3.Open App.Path & "\AVI\FolderUpdate.avi"
Animation4.Open App.Path & "\AVI\Find.avi"

End Sub


Private Sub SetAviStatusBar(piPanel As Integer)
'NOTE: Project Needs to reference
'  Microsoft Windows Common Controls
'  Microsoft Windows Common Controls 2

    Dim BdrWth As Single
    
    iCurPanel = piPanel
    
  
    ScaleMode = vbTwips
    
    'Determine Border Width - May need to be adjusted if a different border style is used
    BdrWth = (Width - ScaleWidth) / 4
    
    'Position Animation Control - "Anim" is the Panel Key we want to draw on top of
    Animation1.Move sbMain.Panels(piPanel).Left, sbMain.Top + BdrWth, sbMain.Panels(piPanel).Width, sbMain.Height - BdrWth
    
    'show after move
    
    
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label1.ForeColor = vbBlack
End Sub

Private Sub Form_Paint()
'James code was all in the form_paint event, I move it into a function.
'the code needs to be here so that on re-size the avi will follow the status bar.

SetAviStatusBar iCurPanel
End Sub

Private Sub Form_Resize()
'pjs
'when form is resized the image does not re-paint sometimes.
'Esp. when min, max, restored... so I put the function in
'the form re-size also???

SetAviStatusBar iCurPanel

End Sub

Private Sub Label1_Click()
     Picture1_Click
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = vbYellow
End Sub

Private Sub Option1_Click()
    SetAviStatusBar 1
End Sub

Private Sub Option2_Click()
    SetAviStatusBar 2
End Sub

Private Sub Option3_Click()
    SetAviStatusBar 3
End Sub

Private Sub Option4_Click()
    SetAviStatusBar 4
End Sub

Private Sub Option6_Click()
    Animation1.Open App.Path & "\AVI\Find.avi"
    SetAviStatusBar iCurPanel
End Sub

Private Sub Option7_Click()
    Animation1.Open App.Path & "\AVI\FolderUpdate.avi"
    SetAviStatusBar iCurPanel
End Sub

Private Sub Option8_Click()
    Animation1.Open App.Path & "\AVI\Alert.avi"
    SetAviStatusBar iCurPanel
End Sub

Private Sub Picture1_Click()
LoadWebPage "http://pjs-inc.com/vb-avi", Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbYellow
End Sub
