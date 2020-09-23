VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1995
   ClientLeft      =   5460
   ClientTop       =   3975
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMessage.frx":0000
   ScaleHeight     =   1995
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMessage.frx":CD64
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3630
      TabIndex        =   1
      Top             =   1590
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   480
      TabIndex        =   0
      Top             =   450
      Width           =   4335
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
SetFormPosition frmMessage.hwnd, True
End Sub

Private Sub Form_Initialize()
SetFormPosition frmMessage.hwnd, True
End Sub

Private Sub Form_Load()
SetFormPosition frmMessage.hwnd, True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = vbBlack
End Sub

Private Sub lblAdd_Click()
Unload Me
End Sub

Private Sub lblAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = vbWhite
End Sub
