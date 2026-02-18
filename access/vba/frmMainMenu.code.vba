Option Compare Database
Option Explicit

Private Sub cmdOpenItem_Click()
    DoCmd.OpenForm "frmItem"
End Sub

Private Sub cmdOpenBOM_Click()
    DoCmd.OpenForm "frmBOM"
End Sub

Private Sub cmdOpenReport_Click()
    If TempVarExists("BOMRunID") Then
        DoCmd.OpenReport "rptBOM", acViewPreview
    Else
        MsgBoxU U("0627 0628 062A 062F 0627 0020 06CC 06A9 06CC 0020 0627 0632 0020 0042 004F 004D 0647 0627 0020 0631 0627 0020 062F 0631 0020 0641 0631 0645 0020 0645 062F 06CC 0631 06CC 062A 0020 0042 004F 004D 0020 0628 0627 0632 0020 06A9 0646 06CC 062F 0020 0648 0020 06AF 0632 06CC 0646 0647 0020 0686 0627 067E 0020 0042 004F 004D 0020 0631 0627 0020 0628 0632 0646 06CC 062F 002E"), vbInformation
        DoCmd.OpenForm "frmBOM"
    End If
End Sub

Private Sub cmdCloseMenu_Click()
    DoCmd.Close acForm, Me.Name
End Sub
