Attribute VB_Name = "mdlCommands"
'This is a module...basically a big hunk of code...
'anything in here can be used through the entire program.

Sub UnTick()
    frmMsgBoxes.OptYesNo.Value = False
    frmMsgBoxes.OptOKCancel.Value = False
    frmMsgBoxes.OptOK.Value = False
    frmMsgBoxes.OptRetryIgnore.Value = False
End Sub

Sub UnloadAll()
    Unload frmButtons
    Unload frmCmnDialog
    Unload frmExplorer
    Unload frmListBoxes
    Unload frmMain
    Unload frmMsgBoxes
    Unload frmtext
    Unload frmTimers
End Sub

Sub ShowAll()
    frmButtons.Show
    frmCmnDialog.Show
    frmExplorer.Show
    frmListBoxes.Show
    frmMsgBoxes.Show
    frmtext.Show
    frmTimers.Show
End Sub

Sub CloseAll()
    Unload frmButtons
    Unload frmCmnDialog
    Unload frmExplorer
    Unload frmListBoxes
    Unload frmMsgBoxes
    Unload frmtext
    Unload frmTimers
End Sub
