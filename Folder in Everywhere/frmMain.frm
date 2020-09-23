VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Folder Creater"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12195
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5880
      Top             =   8190
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8625
      Left            =   0
      Picture         =   "frmMain.frx":1242
      ScaleHeight     =   8625
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2910
         TabIndex        =   25
         Text            =   "0"
         Top             =   810
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CheckBox chRename 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "Rename"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7470
         TabIndex        =   24
         Top             =   5430
         Width           =   945
      End
      Begin VB.CheckBox chCopy 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "Copy"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6510
         TabIndex        =   23
         Top             =   5430
         Width           =   765
      End
      Begin VB.CheckBox chCut 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "Cut"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5580
         TabIndex        =   22
         Top             =   5430
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2070
         Top             =   5670
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chControl 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "Control Panel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Top             =   6540
         Width           =   1365
      End
      Begin VB.CheckBox chDesk 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "Desktop"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6960
         TabIndex        =   20
         Top             =   6180
         Width           =   1365
      End
      Begin VB.CheckBox chComp 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "My Computer"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5550
         TabIndex        =   19
         Top             =   6180
         Width           =   1365
      End
      Begin VB.CheckBox chDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "Delete"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   5430
         Width           =   825
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00979797&
         Caption         =   "Program Options"
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   3720
         TabIndex        =   10
         Top             =   6930
         Width           =   4785
         Begin VB.CheckBox chAlways 
            Appearance      =   0  'Flat
            BackColor       =   &H00979797&
            Caption         =   "Always On Top"
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
            Left            =   390
            TabIndex        =   17
            Top             =   330
            Width           =   1935
         End
      End
      Begin VB.TextBox txtIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4650
         TabIndex        =   8
         Top             =   4470
         Width           =   3675
      End
      Begin VB.TextBox txtBrowse 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   3510
         Width           =   3675
      End
      Begin VB.OptionButton opFolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "Folder"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6300
         TabIndex        =   5
         Top             =   2550
         Width           =   1125
      End
      Begin VB.OptionButton opFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6300
         TabIndex        =   4
         Top             =   2130
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.ListBox lstName 
         Appearance      =   0  'Flat
         BackColor       =   &H00979797&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2505
         ItemData        =   "frmMain.frx":151B24
         Left            =   120
         List            =   "frmMain.frx":151B26
         TabIndex        =   2
         Top             =   780
         Width           =   2745
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4710
         TabIndex        =   1
         Top             =   1320
         Width           =   4005
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   6540
         TabIndex        =   18
         Top             =   8280
         Width           =   2475
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
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
         Left            =   9990
         TabIndex        =   16
         Top             =   5010
         Width           =   1605
      End
      Begin VB.Label lblDelete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete"
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
         Left            =   9990
         TabIndex        =   15
         Top             =   4020
         Width           =   1605
      End
      Begin VB.Label lblClear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clear"
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
         Left            =   9990
         TabIndex        =   14
         Top             =   3030
         Width           =   1605
      End
      Begin VB.Label lblChange 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Change"
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
         Left            =   9990
         TabIndex        =   13
         Top             =   1950
         Width           =   1605
      End
      Begin VB.Label lblAdd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
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
         Left            =   9960
         TabIndex        =   12
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblIcon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8490
         TabIndex        =   9
         Top             =   4470
         Width           =   375
      End
      Begin VB.Label lblBrowse 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8520
         TabIndex        =   7
         Top             =   3510
         Width           =   375
      End
      Begin VB.Label lblSelected 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   330
         TabIndex        =   3
         Top             =   4410
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Grenik Poghosyan
'E-Mail:gmah2005@mail.ru
Dim cx As Single, cy As Single, dx As Single, dy As Single
Dim bDrag As Boolean
Dim Folder As String
Dim AlwaysOn As Integer
Dim Checked As Boolean
Dim FileName As String
Dim n As Integer
Dim Hapavum As String
Dim regName As String
Private Sub chAlways_Click()
If chAlways.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder", "AlwaysOnTop", 0
SetFormPosition Me.hwnd, False
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder", "AlwaysOnTop", 1
SetFormPosition Me.hwnd, True
End If
End Sub

Private Sub chComp_Click()
If chComp.Value = 0 And chDesk.Value = 0 And chControl.Value = 0 Then
Checked = False
Else
Checked = True
End If
End Sub

Private Sub chControl_Click()
If chComp.Value = 0 And chDesk.Value = 0 And chControl.Value = 0 Then
Checked = False
Else
Checked = True
End If
End Sub
Private Sub chDesk_Click()
If chComp.Value = 0 And chDesk.Value = 0 And chControl.Value = 0 Then
Checked = False
Else
Checked = True
End If
End Sub
Private Sub Form_Load()
If GetSetting("txtt", "txtt", "count") = "" Then
n = 0
Else
n = Int(GetSetting("txtt", "txtt", "count"))
End If
For i = 1 To n
lstName.AddItem GetSetting("txtt", "txtt", "Item" & i)
Next
Dim St As String, c, CC
pic.Move 0, 0
For c = 0 To 2
CC = CC + (pic.Width \ 100) * 10
Next c
Main1
Timer1_Timer
Folder = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder", "Caspra")
AlwaysOn = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder", "AlwaysOnTop")
Me.Left = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder", "X")
Me.Top = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder", "Y")
Text1.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder", "Counter")
If Text1.Text = "" Then
Text1.Text = "0"
End If
If AlwaysOn = 1 Then
chAlways.Value = 1
chAlways_Click
Else
chAlways.Value = 0
End If
If Folder = "" Then
SaveKey HKEY_CURRENT_USER, "Software\Caspra\Folder"
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder", "Caspra", App.Path & "\" & App.EXEName & ".exe"
End If
End Sub
Private Sub lblAdd_Click()
lblSelected.Caption = ""
Dim sWord As String
Dim sProperWord As String
Dim ssWord As String
Dim ssProperWord As String
FileName = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Name")
sWord = FileName
sProperWord = UCase$(Left$(sWord, 1))
sProperWord = sProperWord & LCase$(Mid$(sWord, 2))
FileName = sProperWord
ssWord = txtName
ssProperWord = UCase$(Left$(ssWord, 1))
ssProperWord = ssProperWord & LCase$(Mid$(ssWord, 2))
txtName = ssProperWord
If txtName.Text <> "" And Checked = True And txtBrowse.Text <> "" And txtIcon.Text <> "" Then
If FileName = txtName.Text Then
frmMessage.Label1.Caption = "A File with the name you specified already exists."
frmMessage.Show 1
Else
lstName.AddItem txtName.Text
n = n + 1
SaveSetting "txtt", "txtt", "count", Str(n)
For i = 1 To n
SaveSetting "txtt", "txtt", "item" & i, lstName.List(i - 1)
Next
SaveKey HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Name", txtName.Text
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Browse", txtBrowse.Text
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Icon", txtIcon.Text
Text1.Text = (-Text1.Text - 1) * -1
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder", "Counter", Text1.Text
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}", "", txtName.Text
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\DefaultIcon", "", txtIcon.Text
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\InProcServer32", "", "Shell32.dll"
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\InProcServer32", "ThreadingModel", "Apartment"
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder", txtName.Text, "{00020D75-0000-0000-C000-00000000004" & Text1.Text
If chDelete.Value = 0 And chCut.Value = 0 And chCopy.Value = 0 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 0
End If
If chDelete.Value = 0 And chCut.Value = 0 And chCopy.Value = 1 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 1
End If
If chDelete.Value = 0 And chCut.Value = 1 And chCopy.Value = 0 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 2
End If
If chDelete.Value = 0 And chCut.Value = 1 And chCopy.Value = 1 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 3
End If
If chDelete.Value = 0 And chCut.Value = 0 And chCopy.Value = 0 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 16
End If
If chDelete.Value = 0 And chCut.Value = 0 And chCopy.Value = 1 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 17
End If
If chDelete.Value = 0 And chCut.Value = 1 And chCopy.Value = 0 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 18
End If
If chDelete.Value = 0 And chCut.Value = 1 And chCopy.Value = 1 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 19
End If
If chDelete.Value = 1 And chCut.Value = 0 And chCopy.Value = 0 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 32
End If
If chDelete.Value = 1 And chCut.Value = 0 And chCopy.Value = 1 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 33
End If
If chDelete.Value = 1 And chCut.Value = 1 And chCopy.Value = 0 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 34
End If
If chDelete.Value = 1 And chCut.Value = 1 And chCopy.Value = 1 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 35
End If
If chDelete.Value = 1 And chCut.Value = 0 And chCopy.Value = 0 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 48
End If
If chDelete.Value = 1 And chCut.Value = 0 And chCopy.Value = 1 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 49
End If
If chDelete.Value = 1 And chCut.Value = 1 And chCopy.Value = 0 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 50
End If
If chDelete.Value = 1 And chCut.Value = 1 And chCopy.Value = 1 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 51
End If
If chComp.Value = 1 Then
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}", "", ""
End If
If chDesk.Value = 1 Then
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}", "", ""
End If
If chControl.Value = 1 Then
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}", "", ""
End If
If opFile.Value = True Then
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\Shell\Open\Command", "", txtBrowse.Text
Else
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\Shell\Open\Command", "", "explorer.exe " & txtBrowse.Text
End If
If chComp.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "My computer", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "My computer", "1"
End If
If chDesk.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Desktop", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Desktop", "1"
End If
If chControl.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Control Panel", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Control Panel", "1"
End If
If chDelete.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Delete", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Delete", "1"
End If
If chCut.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Cut", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Cut", "1"
End If
If chCopy.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Copy", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Copy", "1"
End If
If chRename.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Rename", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Rename", "1"
End If
End If
Else
frmMessage.Label1.Caption = "You must fill all textboxes and check one of places."
frmMessage.Show 1
End If
End Sub

Private Sub lblAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.BackColor = &HC0C0C0
End Sub

Private Sub lblBrowse_Click()
If opFolder.Value = True Then
txtBrowse.Text = BrowseForFolder(Me.hwnd, "Choose Folder ....")
Else
On Error GoTo ErrHandler
CommonDialog1.CancelError = True
CommonDialog1.FilterIndex = 2
CommonDialog1.Filter = "Programs (*.exe)|*.exe|"
CommonDialog1.ShowOpen
txtBrowse.Text = CommonDialog1.FileName
Exit Sub
ErrHandler:
Exit Sub
End If
End Sub

Private Sub lblBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBrowse.BackColor = &HC0C0C0
End Sub

Private Sub lblChange_Click()
If txtName.Text <> "" And txtBrowse.Text <> "" And txtIcon.Text <> "" And Checked = True Then
If txtName.Text = lstName.Text Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Name", txtName.Text
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Browse", txtBrowse.Text
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Icon", txtIcon.Text
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}", "", txtName.Text
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\DefaultIcon", "", txtIcon.Text
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\InProcServer32", "", "Shell32.dll"
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\InProcServer32", "ThreadingModel", "Apartment"
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder", txtName.Text, "{00020D75-0000-0000-C000-00000000004" & Text1.Text
If chDelete.Value = 0 And chCut.Value = 0 And chCopy.Value = 0 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 0
End If
If chDelete.Value = 0 And chCut.Value = 0 And chCopy.Value = 1 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 1
End If
If chDelete.Value = 0 And chCut.Value = 1 And chCopy.Value = 0 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 2
End If
If chDelete.Value = 0 And chCut.Value = 1 And chCopy.Value = 1 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 3
End If
If chDelete.Value = 0 And chCut.Value = 0 And chCopy.Value = 0 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 16
End If
If chDelete.Value = 0 And chCut.Value = 0 And chCopy.Value = 1 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 17
End If
If chDelete.Value = 0 And chCut.Value = 1 And chCopy.Value = 0 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 18
End If
If chDelete.Value = 0 And chCut.Value = 1 And chCopy.Value = 1 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 19
End If
If chDelete.Value = 1 And chCut.Value = 0 And chCopy.Value = 0 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 32
End If
If chDelete.Value = 1 And chCut.Value = 0 And chCopy.Value = 1 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 33
End If
If chDelete.Value = 1 And chCut.Value = 1 And chCopy.Value = 0 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 34
End If
If chDelete.Value = 1 And chCut.Value = 1 And chCopy.Value = 1 And chRename.Value = 0 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 35
End If
If chDelete.Value = 1 And chCut.Value = 0 And chCopy.Value = 0 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 48
End If
If chDelete.Value = 1 And chCut.Value = 0 And chCopy.Value = 1 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 49
End If
If chDelete.Value = 1 And chCut.Value = 1 And chCopy.Value = 0 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 50
End If
If chDelete.Value = 1 And chCut.Value = 1 And chCopy.Value = 1 And chRename.Value = 1 Then
SaveDword HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\ShellFolder", "Attributes", 51
End If
If chComp.Value = 1 Then
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}", "", ""
End If
If chDesk.Value = 1 Then
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}", "", ""
End If
If chControl.Value = 1 Then
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}", "", ""
End If
If opFile.Value = True Then
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\Shell\Open\Command", "", txtBrowse.Text
Else
SaveString HKEY_CLASSES_ROOT, "CLSID\{00020D75-0000-0000-C000-00000000004" & Text1.Text & "}\Shell\Open\Command", "", "explorer.exe " & txtBrowse.Text
End If
If chComp.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "My computer", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "My computer", "1"
End If
If chDesk.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Desktop", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Desktop", "1"
End If
If chControl.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Control Panel", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Control Panel", "1"
End If
If chDelete.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Delete", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Delete", "1"
End If
If chCut.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Cut", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Cut", "1"
End If
If chCopy.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Copy", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Copy", "1"
End If
If chRename.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Rename", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Rename", "1"
End If
If chComp.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "My computer", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "My computer", "1"
End If
If chDesk.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Desktop", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Desktop", "1"
End If
If chControl.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Control Panel", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Control Panel", "1"
End If
If chDelete.Value = 0 Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Deleteble", "0"
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder\" & txtName.Text, "Deleteble", "1"
End If
frmMessage.Label1.Caption = "Information has been Changed!!"
frmMessage.Show 1
Else
frmMessage.Label1.Caption = "The name of Folder or File can not be changed"
frmMessage.Show 1
End If
Else
frmMessage.Label1.Caption = "There is nothing to change"
frmMessage.Show 1
End If
End Sub

Private Sub lblChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblChange.BackColor = &HC0C0C0
End Sub

Private Sub lblClear_Click()
txtName.Text = ""
txtBrowse.Text = ""
txtIcon.Text = ""
chDelete.Value = 0
chRename.Value = 0
chCut.Value = 0
chCopy.Value = 0
chComp.Value = 0
chDesk.Value = 0
chControl.Value = 0
lblSelected.Caption = ""
End Sub

Private Sub lblClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClear.BackColor = &HC0C0C0
End Sub
Private Sub lblClose_Click()
Unload Me
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.BackColor = &HC0C0C0
End Sub

Private Sub lblDelete_Click()
On Error GoTo CantDelete
If lblSelected.Caption <> "" Then
regName = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder", lblSelected.Caption) & "}"
lstName.RemoveItem lstName.ListIndex
n = n - 1
SaveSetting "txtt", "txtt", "count", Str(n)
For i = 1 To n
SaveSetting "txtt", "txtt", "item" & i, lstName.List(i - 1)
Next
DeleteKey HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lblSelected.Caption
DeleteKey HKEY_CLASSES_ROOT, "CLSID\" & regName & "\DefaultIcon"
DeleteKey HKEY_CLASSES_ROOT, "CLSID\" & regName & "\Shell\Open\Command"
DeleteKey HKEY_CLASSES_ROOT, "CLSID\" & regName & "\Shell\Open"
DeleteKey HKEY_CLASSES_ROOT, "CLSID\" & regName & "\Shell"
DeleteKey HKEY_CLASSES_ROOT, "CLSID\" & regName & "\InProcServer32"
DeleteKey HKEY_CLASSES_ROOT, "CLSID\" & regName & "\ShellFolder"
DeleteKey HKEY_CLASSES_ROOT, "CLSID\" & regName
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\" & regName
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\" & regName
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\" & regName
DeleteValue HKEY_CURRENT_USER, "Software\Caspra\Folder", lblSelected.Caption
lblSelected.Caption = ""
lblClear_Click
Else
frmMessage.Label1.Caption = "Before clicking Delete you must select item"
frmMessage.Show 1
End If
Exit Sub
CantDelete:
Exit Sub
End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDelete.BackColor = &HC0C0C0
End Sub

Private Sub lblIcon_Click()
On Error GoTo ErrHandler
CommonDialog1.CancelError = True
CommonDialog1.FilterIndex = 2
CommonDialog1.Filter = "Icons (*.ico)|*.ico|"
CommonDialog1.ShowOpen
txtIcon.Text = CommonDialog1.FileName
Exit Sub
ErrHandler:
Exit Sub
End Sub

Private Sub lblIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblIcon.BackColor = &HC0C0C0
End Sub

Private Sub lstName_Click()
lblSelected.Caption = lstName.Text
txtName.Text = lstName.Text
txtBrowse.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "Browse")
txtIcon.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "Icon")
chDelete.Value = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "Delete")
chCut.Value = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "Cut")
chCopy.Value = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "Copy")
chRename.Value = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "Rename")
chComp.Value = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "My Computer")
chDesk.Value = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "Desktop")
chControl.Value = GetString(HKEY_CURRENT_USER, "Software\Caspra\Folder\" & lstName.Text, "Control Panel")
Hapavum = Right(txtBrowse.Text, 3)
If Hapavum = "exe" Then
opFile.Value = True
Else
opFolder.Value = True
End If
End Sub

Private Sub opFile_Click()
If lblSelected = "" Then
txtBrowse.Text = ""
End If
End Sub

Private Sub opFolder_Click()
If lblSelected = "" Then
txtBrowse.Text = ""
End If
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
bDrag = True
cx = Me.Left
cy = Me.Top
dx = X
dy = Y
lblSelected.Caption = ""
End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bDrag Then
cx = cx + X - dx
cy = cy + Y - dy
Me.Move cx, cy
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder", "X", Me.Left
SaveString HKEY_CURRENT_USER, "Software\Caspra\Folder", "Y", Me.Top
End If
lblBrowse.BackColor = 8421504
lblIcon.BackColor = 8421504
lblAdd.BackColor = 8421504
lblChange.BackColor = 8421504
lblClear.BackColor = 8421504
lblDelete.BackColor = 8421504
lblClose.BackColor = 8421504
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bDrag = False
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = Date & " " & Time
End Sub

Private Sub txtName_Click()
lblSelected.Caption = ""
End Sub
