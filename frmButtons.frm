VERSION 5.00
Begin VB.Form frmButtons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buttons"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Close"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdSizableForm 
      Caption         =   "Unload Form"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdTxtChange 
      Caption         =   "Disable TextBox"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtChange 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "TextBox Info..."
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Caption"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    Select Case cmdChange.Caption 'Select what you'll be changing
        Case "Change Caption" 'The original caption of the button
            cmdChange.Caption = "Changed :o)" 'The changed caption of the button when first clicked
        Case "Changed :o)" 'The changed caption of the button ready to be clicked
            cmdChange.Caption = "Change Caption" 'Then once clicked, change the caption back to default caption
    End Select 'End Select
End Sub

Private Sub cmdSizableForm_Click()
    Unload Me 'Unload this form
End Sub

Private Sub cmdTxtChange_Click()
    Select Case cmdTxtChange.Caption
        Case "Disable TextBox"
            txtChange.Enabled = False
                cmdTxtChange.Caption = "Enable TextBox"
        Case "Enable TextBox"
            txtChange.Enabled = True
                cmdTxtChange.Caption = "Disable TextBox"
    End Select
        
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub
