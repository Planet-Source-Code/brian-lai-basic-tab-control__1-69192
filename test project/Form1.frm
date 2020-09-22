VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6495
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   480
      Width           =   6015
      Begin VB.CommandButton Command3 
         Caption         =   "Edit Tab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Tab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add a tab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label LBL 
         BackStyle       =   0  '³z©ú
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label LBL 
         BackStyle       =   0  '³z©ú
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label LBL 
         BackStyle       =   0  '³z©ú
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   3735
      End
   End
   Begin Project1.Tab Tab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Tab1.AddTab
End Sub

Private Sub Command2_Click()
    Tab1.RemoveTab Tab1.ActiveTab
End Sub

Private Sub Command3_Click()
    Tab1.TabCaption(Tab1.ActiveTab) = InputBox("Enter new caption:", , Tab1.TabCaption(Tab1.ActiveTab))
End Sub

Private Sub Form_Load()
    Tab1.AddTabs "Hello", "Hello2", "Hello World"
    Tab1.ActiveTab = 2
End Sub

Private Sub Tab1_Click(tIndex As Integer)
    LBL(0).Caption = "Selected tab: No." & tIndex
    LBL(1).Caption = "Tab's tag: " & Tab1.TabTag(tIndex)
    LBL(2).Caption = "Tab's text: " & Tab1.TabCaption(tIndex)
End Sub
