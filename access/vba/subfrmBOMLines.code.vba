Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me!LineNo = Nz(DMax("LineNo", "tblBOMLines", "BOMHeaderID=" & Me!BOMHeaderID), 0) + 1
    If IsNull(Me!ScrapPct) Then Me!ScrapPct = 0
End Sub

Private Sub cboComponentItemID_AfterUpdate()
    Me!UOMID = Me!cboComponentItemID.Column(3)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim parentID As Long
    parentID = Nz(Me.Parent!FGItemID, 0)

    If parentID = 0 Then Exit Sub

    If Nz(Me!ComponentItemID, 0) = parentID Then
        MsgBoxU U("0642 0637 0639 0647 0020 0646 0645 06CC 200C 062A 0648 0627 0646 062F 0020 0628 0627 0020 06A9 0627 0644 0627 06CC 0020 0648 0627 0644 062F 0020 0028 06A9 0627 0644 0627 06CC 0020 0646 0647 0627 06CC 06CC 002F 0646 06CC 0645 200C 0633 0627 062E 062A 0647 0029 0020 06CC 06A9 0633 0627 0646 0020 0628 0627 0634 062F 002E"), vbExclamation
        Cancel = True
        Exit Sub
    End If

    If Not ValidateNoCycle(parentID, Me!ComponentItemID, Date) Then
        MsgBoxU U("062F 0631 0020 0633 0627 062E 062A 0627 0631 0020 0042 004F 004D 0020 062D 0644 0642 0647 0020 0028 0043 0079 0063 006C 0065 0029 0020 062A 0634 062E 06CC 0635 0020 062F 0627 062F 0647 0020 0634 062F 061B 0020 0630 062E 06CC 0631 0647 0020 0627 0646 062C 0627 0645 0020 0646 0634 062F 002E"), vbCritical
        Cancel = True
        Exit Sub
    End If

    If Nz(Me!QtyPer, 0) <= 0 Then
        MsgBoxU U("0645 0642 062F 0627 0631 0020 0051 0074 0079 0050 0065 0072 0020 0628 0627 06CC 062F 0020 0628 0632 0631 06AF 200C 062A 0631 0020 0627 0632 0020 0635 0641 0631 0020 0628 0627 0634 062F 002E"), vbExclamation
        Cancel = True
        Exit Sub
    End If

    If Nz(Me!ScrapPct, 0) < 0 Or Nz(Me!ScrapPct, 0) > 100 Then
        MsgBoxU U("062F 0631 0635 062F 0020 0636 0627 06CC 0639 0627 062A 0020 0028 0053 0063 0072 0061 0070 0050 0063 0074 0029 0020 0628 0627 06CC 062F 0020 0628 06CC 0646 0020 0030 0020 062A 0627 0020 0031 0030 0030 0020 0628 0627 0634 062F 002E"), vbExclamation
        Cancel = True
        Exit Sub
    End If

    If Not IsNull(Me!EffectiveFrom) And Not IsNull(Me!EffectiveTo) Then
        If Me!EffectiveTo < Me!EffectiveFrom Then
            MsgBoxU U("062A 0627 0631 06CC 062E 0020 067E 0627 06CC 0627 0646 0020 0627 0639 062A 0628 0627 0631 0020 0628 0627 06CC 062F 0020 0628 0632 0631 06AF 200C 062A 0631 0020 06CC 0627 0020 0645 0633 0627 0648 06CC 0020 062A 0627 0631 06CC 062E 0020 0634 0631 0648 0639 0020 0627 0639 062A 0628 0627 0631 0020 0628 0627 0634 062F 002E"), vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = 3022 Then
        MsgBoxU U("062B 0628 062A 0020 062A 06A9 0631 0627 0631 06CC 0020 0642 0637 0639 0647 0020 062F 0631 0020 06CC 06A9 0020 0042 004F 004D 0020 0645 062C 0627 0632 0020 0646 06CC 0633 062A 002E"), vbExclamation
        Response = acDataErrContinue
    End If
End Sub
