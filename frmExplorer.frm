VERSION 5.00
Begin VB.Form frmExplorer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Explorer"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   4335
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path 'When the directory is changed, update the filelist box
End Sub

Private Sub Drive1_Change()
    On Error GoTo handler
        Dir1.Path = Drive1.Drive 'Update Directory box to show contents on new drive
            Exit Sub
handler:
    MsgBox "No disc in drive", vbCritical, "Error" 'Used when there is no disc in the drive
End Sub
