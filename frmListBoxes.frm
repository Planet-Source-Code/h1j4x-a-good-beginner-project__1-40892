VERSION 5.00
Begin VB.Form frmListBoxes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Boxes"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Misc input..."
      Height          =   1455
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdSomething 
         Caption         =   "Something"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdThisFormTitle 
         Caption         =   "This form's title"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdMainFormTitle 
         Caption         =   "Main form title"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   615
   End
   Begin VB.ListBox Lst1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmListBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInput_Click()
    If txtInput.Text = "" Then 'If the input textbox has nothing in it, then display a message box
        MsgBox "Please enter something in the text box to add to the listbox.", vbInformation, "Information"
            Else 'Or if there is stuff in the input textbox, then add it to the listbox
                Lst1.AddItem txtInput.Text 'Add item from input textbox
                    txtInput.Text = "" 'Clear the input textbox
    End If
End Sub

Private Sub cmdMainFormTitle_Click()
    Lst1.AddItem frmMain.Caption 'Add main form's title
End Sub

Private Sub cmdSomething_Click()
    Lst1.AddItem "Something" 'Add something ;)
End Sub

Private Sub cmdThisFormTitle_Click()
    Lst1.AddItem frmListBoxes.Caption 'Add this form's title
End Sub
