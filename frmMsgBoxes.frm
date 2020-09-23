VERSION 5.00
Begin VB.Form frmMsgBoxes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Boxes"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton OptRetryIgnore 
      Caption         =   "Abort/Retry/Ignore"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton OptOKCancel 
      Caption         =   "OK/Cancel"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton OptYesNo 
      Caption         =   "Yes/No"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton OptOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "frmMsgBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If OptYesNo.Value = True Then 'If Option button is ticked then display the message box
        If MsgBox("This is a yes/no message box.", vbYesNo, "Test") = vbYes Then 'If user clicked Yes Then...
            MsgBox "You clicked Yes", vbDefaultButton1, "Result"
                Else 'Or if he clicked No Then...
                    MsgBox "You click No", vbDefaultButton1, "Result"
        End If
    End If
    
    If OptOKCancel.Value = True Then
        MsgBox "This is an OK/Cancel message box.", vbOKCancel, "Test"
    End If
    
    If OptOK.Value = True Then
        MsgBox "This is an OK message box.", vbOKOnly, "Test"
    End If
    
    If OptRetryIgnore.Value = True Then
        MsgBox "This is an Abort/Retry/Ignore message box.", vbAbortRetryIgnore, "Test"
    End If
End Sub

Private Sub Form_Load()
    UnTick 'Call to module to un-tick all the radio buttons on form load
End Sub
