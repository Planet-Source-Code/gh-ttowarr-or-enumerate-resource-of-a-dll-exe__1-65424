VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Enumerate"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView tView 
      Height          =   5015
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8837
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImgLst"
      Appearance      =   1
   End
   Begin VB.ListBox LstFile 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2655
   End
   Begin VB.FileListBox File 
      Height          =   480
      Left            =   120
      Pattern         =   "*.exe;*.dll"
      System          =   -1  'True
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.DirListBox Dir 
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin ComctlLib.ImageList ImgLst 
      Left            =   120
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0712
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir_Change()
Dim i As Integer

File.Path = Dir.Path
LstFile.Clear

For i = 0 To File.ListCount
    LstFile.AddItem File.List(i), i
Next i
End Sub

Private Sub Drive_Change()
Dir.Path = Drive.Drive
End Sub

Private Sub LstFile_Click()
EnumResData File.Path & "\" & LstFile.List(LstFile.ListIndex)
End Sub
