Option Compare Database
Option Explicit

Private Sub cboFGItemID_AfterUpdate()
    If Me.NewRecord Then
        Me!VersionNo = GetNextBOMVersionNo(Me!FGItemID)
        Me!VersionLabel = "v" & Me!VersionNo
    End If
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me!IsActive = False
    Me!ActiveKey = Null
    Me!CreatedOn = Now()
    Me!CreatedBy = Environ$("USERNAME")
End Sub

Private Sub Form_Current()
    On Error Resume Next
    Me.cboFGItemID.Locked = Not Me.NewRecord
    Me.sfmBOMLines.Enabled = Not IsNull(Me!BOMHeaderID)
End Sub

Private Sub cmdSave_Click()
    SaveCurrentRecord
    Me.sfmBOMLines.Enabled = True
End Sub

Private Sub cmdActivateBOM_Click()
    On Error GoTo EH
    SaveCurrentRecord
    ActivateBOM Me!BOMHeaderID, Date
    Me.Requery
    MsgBoxU U("0646 0633 062E 0647 0020 0042 004F 004D 0020 0628 0627 0020 0645 0648 0641 0642 06CC 062A 0020 0641 0639 0627 0644 0020 0634 062F 002E"), vbInformation
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdCopyBOM_Click()
    On Error GoTo EH
    SaveCurrentRecord

    Dim newID As Long
    newID = CopyBOM(Me!BOMHeaderID)

    DoCmd.OpenForm Me.Name, , , "BOMHeaderID=" & newID
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdPrintBOM_Click()
    On Error GoTo EH
    SaveCurrentRecord

    Dim runID As String
    runID = BuildBOMExplosion(Me!BOMHeaderID, Date)

    SetTempVarValue "BOMRunID", runID

    DoCmd.OpenReport "rptBOM", acViewPreview
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub SaveCurrentRecord()
    If Me.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
    End If
End Sub
