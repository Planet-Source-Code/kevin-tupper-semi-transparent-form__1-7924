VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Text            =   "50"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      Top             =   4125
      Width           =   1095
   End
   Begin VB.PictureBox picDC 
      AutoRedraw      =   -1  'True
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Blend Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End
End Sub

Private Sub Command1_Click()
    Me.Cls
    Blend Me, picDC, Text1, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Me.Show
    Me.Refresh
End Sub

Private Sub Form_Load()
    DoEvents
    SetWindowPos Me.hWnd, -1, 100, 100, Me.ScaleWidth, Me.ScaleHeight, SWP_NOOWNERZORDER

    picDC.Width = Me.ScaleWidth + 10
    picDC.Height = Me.ScaleHeight + 10
    picDC.Left = 0
    picDC.Top = 0

    DeskHdc = GetDC(0)
    ret = BitBlt(picDC.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
    ret = ReleaseDC(0&, DeskHdc)
    Blend Me, picDC, 50, 0, 0, Me.ScaleWidth, Me.ScaleHeight

    Me.Show
    Me.Refresh
End Sub
