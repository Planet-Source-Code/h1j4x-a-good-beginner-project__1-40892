VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic Beginner Helper"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdText 
      Caption         =   "Text"
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCloseAll 
      Caption         =   "Close All"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenAll 
      Caption         =   "Open All"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdCommonDialog 
      Caption         =   "Common Dialogs"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdTimers 
      Caption         =   "Timers"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdButtons 
      Caption         =   "Buttons"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdListBoxes 
      Caption         =   "ListBoxes"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdMsgBoxes 
      Caption         =   "Message Boxes"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdExplorer 
      Caption         =   "Explorer"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program was created by H1j4x on 11/19/02.
' It is intended to help beginners with the basics of VB programming.
' If you have any questions or comments, drop me a line at st3v3_chambers@hotmail.com
' If you like this code, please rate it at psc :o)

Private Sub cmdButtons_Click()
    frmButtons.Show 'Show the Buttons form
End Sub

Private Sub cmdCloseAll_Click()
    CloseAll 'Call to module to close all except main form
End Sub

Private Sub cmdCommonDialog_Click()
    frmCmnDialog.Show 'Show the Common Dialog form
End Sub

Private Sub cmdExit_Click()
    UnloadAll 'Call to module to unload everything
End Sub

Private Sub cmdExplorer_Click()
    frmExplorer.Show 'Show the Explorer form
End Sub

Private Sub cmdListBoxes_Click()
    frmListBoxes.Show 'Show the List Boxes form
End Sub

Private Sub cmdMsgBoxes_Click()
    frmMsgBoxes.Show 'Show the Message Boxes form
End Sub

Private Sub cmdOpenAll_Click()
    ShowAll 'Call to module to open everything
End Sub

Private Sub cmdText_Click()
    frmtext.Show 'Show the Text form
End Sub

Private Sub cmdTimers_Click()
    frmTimers.Show 'Show the Timers form
End Sub
