VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStuff 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cut"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopy_Click()
    Clipboard.SetText txtStuff.SelText 'Copy the selected text to the clipboard
End Sub

Private Sub cmdCut_Click()
    Clipboard.SetText txtStuff.SelText 'Copt the selected text to the clipboard
    txtStuff.SelText = "" 'and delete it from the text box
End Sub

Private Sub cmdDelete_Click()
    txtStuff.SelText = "" 'Delete the selected text
End Sub

Private Sub cmdPaste_Click()
    txtStuff.SelText = Clipboard.GetText 'Paste info from clipboard
End Sub

