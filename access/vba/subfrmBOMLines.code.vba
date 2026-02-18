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
        MsgBox "Component cannot be equal to parent FG/Assembly.", vbExclamation
        Cancel = True
        Exit Sub
    End If

    If Not ValidateNoCycle(parentID, Me!ComponentItemID, Date) Then
        MsgBox "Cycle detected in BOM structure. Save is blocked.", vbCritical
        Cancel = True
        Exit Sub
    End If

    If Nz(Me!QtyPer, 0) <= 0 Then
        MsgBox "QtyPer must be greater than zero.", vbExclamation
        Cancel = True
        Exit Sub
    End If

    If Nz(Me!ScrapPct, 0) < 0 Or Nz(Me!ScrapPct, 0) > 100 Then
        MsgBox "ScrapPct must be between 0 and 100.", vbExclamation
        Cancel = True
        Exit Sub
    End If

    If Not IsNull(Me!EffectiveFrom) And Not IsNull(Me!EffectiveTo) Then
        If Me!EffectiveTo < Me!EffectiveFrom Then
            MsgBox "EffectiveTo must be greater than or equal to EffectiveFrom.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = 3022 Then
        MsgBox "Duplicate component is not allowed in the same BOM.", vbExclamation
        Response = acDataErrContinue
    End If
End Sub
