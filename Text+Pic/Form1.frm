VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   Caption         =   "  Text+Picture ::.."
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Fit"
      Height          =   315
      Left            =   1680
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "AutoSize"
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Font"
      Height          =   255
      Left            =   6600
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "File"
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "R"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "B"
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
      Left            =   4440
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Color"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cmg 
      Left            =   960
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   120
      Picture         =   "Form1.frx":15162
      ScaleHeight     =   3660
      ScaleWidth      =   7095
      TabIndex        =   0
      Top             =   1560
      Width           =   7125
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Put The text where you want."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   1
         Top             =   3120
         Width           =   2925
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      Height          =   195
      Left            =   3000
      TabIndex        =   13
      Top             =   240
      Width           =   285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Size:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text to print:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sc As Long 'scale
Dim MoveReady As Boolean


Private Sub Command1_Click()
savePic
End Sub

Private Sub Command10_Click()
pic.AutoSize = False
rez
End Sub

Private Sub Command2_Click()
cmg.ShowColor
Label1.ForeColor = cmg.Color
End Sub
Private Sub Command3_Click()
Label1.FontBold = True
End Sub
Private Sub Command4_Click()
Label1.FontItalic = True
End Sub
Private Sub Command5_Click()
Label1.FontBold = False
Label1.FontItalic = False
End Sub
Private Sub Command6_Click()
pic.Cls
End Sub
Private Sub Command7_Click()
On Error GoTo 1
cmg.ShowOpen
If cmg.FileName <> "" Then
   pic.Cls
   pic = LoadPicture(cmg.FileName)
   Text2 = cmg.FileName
End If
1: If Error(Err.Number) <> "" Then
      MsgBox Error(Err.Number)
   End If
End Sub


Private Sub Command8_Click()
'on my computer i recive an error that there is no font setup.what about you?
On Error GoTo 1
   cmg.ShowFont
   If cmg.FontName <> "" Then
      Label1.Font = cmg.FontName
   End If
1: If Error(Err.Number) <> "" Then
      MsgBox Error(Err.Number)
   End If
End Sub

Private Sub Command9_Click()
pic.AutoSize = True
rez
End Sub

Private Sub Form_Load()
MoveReady = False
End Sub

Private Sub Form_Resize()

pic.Width = Me.Width - 1200
pic.Height = Me.Height - 3000
rez
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveReady = True
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveReady = True Then
Label1.Left = X
Label1.Top = Y
DoEvents
End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

MoveReady = False
End Sub

Private Sub Text1_Change()
Label1 = Text1
End Sub
Private Sub savePic()
pic.CurrentY = Label1.Top
pic.CurrentX = Label1.Left
pic.FontBold = Label1.FontBold
pic.FontItalic = Label1.FontItalic
pic.FontSize = Label1.FontSize
pic.ForeColor = Label1.ForeColor
pic.Font = Label1.Font
pic.Print Label1
SavePicture pic.Image, "c:\test.bmp"
End Sub

Private Sub Text3_Change()
On Error Resume Next
Label1.FontSize = Text3
End Sub
Private Sub rez()
Dim picW, picH As Long
picW = pic.Width
picH = pic.Height
pic.Cls
pic.PaintPicture pic.Picture, 0, 0, picW, picH, 0, 0
End Sub
