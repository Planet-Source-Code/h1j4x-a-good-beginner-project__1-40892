VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Timers"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Progress Bar"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   75
         Left            =   120
         Top             =   720
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "Enable"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time && Date"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   720
      End
      Begin VB.CommandButton cmdDisplayTime 
         Caption         =   "Start Time && Date"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblTime 
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is basically the same code as the code from the buttons form just with timers...
Private Sub cmdDisplayTime_Click()
    Select Case cmdDisplayTime.Caption 'Select what you'll be changing
        Case "Start Time && Date"
            Timer1.Enabled = True 'Enable the timer...which will display time and date
                cmdDisplayTime.Caption = "Stop"
        Case "Stop"
            Timer1.Enabled = False 'Disable the timer...which will stop the display of time and date
                lblTime.Caption = ""
                    cmdDisplayTime.Caption = "Start Time && Date"
    End Select
End Sub

Private Sub cmdEnable_Click()
    Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = Now 'This tells that the label will display the date and time when timer is enabled
End Sub

Private Sub Timer2_Timer()
    ProgressBar1.Value = ProgressBar1.Value + 1 'When timer enabled, make the progress bar advance 1 bar
        If ProgressBar1.Value = 100 Then 'If progress bar is full, then reset it
            ProgressBar1.Value = 0 'Resets progress bar
            Timer2.Enabled = False 'Disables timer until called for again
                End If
End Sub
