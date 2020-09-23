VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCmnDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Common Dialogs"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3625
      _Version        =   393217
      TextRTF         =   $"frmCmnDialog.frx":0000
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CmnDlg1 
      Left            =   2880
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCmnDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
    CmnDlg1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*" 'This sets the filter for common dialog
        CmnDlg1.ShowOpen 'Make common dialog show the open dialog
            RTB1.LoadFile (CmnDlg1.FileName) 'Loads the selected file you chose into the rich text box
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next 'If error occurs, ignore it
        Dim FileName As String 'Global variable...can be used anywhere in the program
            CmnDlg1.Filter = "Text Files (*.txt) |*.txt| All Files (*.*) |*.*|" 'Filter...same as above
                CmnDlg1.Action = 2 'Tells common dialog which action to perform...In this case, it's "save as"
                    FileName = CmnDlg1.FileName 'Makes the variable = the common dialog filename you specified
                        F = FreeFile
                    Open FileName For Output As #F 'Opens the file for output
                Print #F, "Beginner Info..." & vbNewLine & vbNewLine & RTB1.Text 'This really isn't needed besides the RTB1.Text, but all it does it automatically enter some stuff into the output file before what you typed in the rich text box
            Close #F 'Close the file
End Sub
