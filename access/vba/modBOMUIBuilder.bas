Attribute VB_Name = "modBOMUIBuilder"
Option Compare Database
Option Explicit

Public Sub AutoSetupAllWithUI()
    On Error GoTo EH

    AutoSetupBOM True
    BuildBOMUIObjects
    DoCmd.OpenForm "frmMainMenu"

    MsgBoxU U("0631 0627 0647 200C 0627 0646 062F 0627 0632 06CC 0020 0628 062E 0634 0020 062F 0627 062F 0647 0020 0648 0020 0631 0627 0628 0637 0020 06A9 0627 0631 0628 0631 06CC 0020 0628 0627 0020 0645 0648 0641 0642 06CC 062A 0020 0627 0646 062C 0627 0645 0020 0634 062F 002E"), vbInformation
    Exit Sub
EH:
    MsgBoxU U("0631 0627 0647 200C 0627 0646 062F 0627 0632 06CC 0020 06A9 0627 0645 0644 0020 0042 004F 004D 0020 0646 0627 0645 0648 0641 0642 0020 0628 0648 062F 003A 0020") & Err.Description, vbCritical
End Sub

Public Sub BuildBOMUIObjects()
    On Error GoTo EH

    CloseAllOpenFormsAndReports

    BuildForm_SubBOMLines
    BuildForm_Item
    BuildForm_BOM
    BuildForm_MainMenu
    BuildReport_BOM

    Exit Sub
EH:
    MsgBoxU U("0633 0627 062E 062A 0020 0622 0628 062C 06A9 062A 200C 0647 0627 06CC 0020 0631 0627 0628 0637 0020 06A9 0627 0631 0628 0631 06CC 0020 0646 0627 0645 0648 0641 0642 0020 0628 0648 062F 003A 0020") & Err.Description, vbCritical
End Sub

Private Sub CloseAllOpenFormsAndReports()
    On Error Resume Next

    Dim i As Long
    For i = Forms.Count - 1 To 0 Step -1
        DoCmd.Close acForm, Forms(i).Name, acSaveNo
    Next i

    For i = Reports.Count - 1 To 0 Step -1
        DoCmd.Close acReport, Reports(i).Name, acSaveNo
    Next i

    On Error GoTo 0
End Sub

Private Sub BuildForm_Item()
    Const TARGET As String = "frmItem"

    DeleteFormIfExists TARGET

    Dim frm As Form
    Set frm = CreateForm()

    Dim tmpName As String
    tmpName = frm.Name

    frm.RecordSource = "tblItems"
    frm.DefaultView = 0
    frm.Caption = U("062A 0639 0631 06CC 0641 0020 06A9 0627 0644 0627 0020 002F 0020 0642 0637 0639 0647")
    frm.Width = 12000

    AddLabel tmpName, U("06A9 062F 0020 06A9 0627 0644 0627"), 500, 500, 1800, 300
    AddTextBox tmpName, "txtItemCode", "ItemCode", 2600, 450, 3000, 360

    AddLabel tmpName, U("0634 0631 062D 0020 06A9 0627 0644 0627"), 500, 1000, 1800, 300
    AddTextBox tmpName, "txtItemDescription", "ItemDescription", 2600, 950, 6000, 360

    AddLabel tmpName, U("0648 0627 062D 062F"), 500, 1500, 1800, 300
    Dim cboUOM As Control
    Set cboUOM = AddComboBox(tmpName, "cboUOMID", "UOMID", 2600, 1450, 2800, 360)
    cboUOM.RowSource = "SELECT UOMID, UOMCode, UOMName FROM tblUOM WHERE IsActive=True ORDER BY UOMCode;"
    cboUOM.BoundColumn = 1
    cboUOM.ColumnCount = 3
    cboUOM.ColumnWidths = "0cm;2.5cm;4cm"

    AddLabel tmpName, U("0646 0648 0639 0020 06A9 0627 0644 0627"), 500, 2000, 1800, 300
    Dim cboType As Control
    Set cboType = AddComboBox(tmpName, "cboItemTypeID", "ItemTypeID", 2600, 1950, 2800, 360)
    cboType.RowSource = "SELECT ItemTypeID, TypeCode, TypeName FROM tblItemType ORDER BY TypeCode;"
    cboType.BoundColumn = 1
    cboType.ColumnCount = 3
    cboType.ColumnWidths = "0cm;2.5cm;4cm"

    AddLabel tmpName, U("0641 0639 0627 0644"), 500, 2500, 1800, 300
    AddCheckBox tmpName, "chkIsActive", "IsActive", 2600, 2450, 500, 360

    AddLabel tmpName, U("0647 0632 06CC 0646 0647 0020 0627 0633 062A 0627 0646 062F 0627 0631 062F"), 500, 3000, 1800, 300
    AddTextBox tmpName, "txtStdCost", "StdCost", 2600, 2950, 1800, 360

    AddLabel tmpName, U("062A 0627 0631 06CC 062E 0020 0627 06CC 062C 0627 062F"), 500, 3600, 1800, 300
    Dim txtCreatedOn As Control
    Set txtCreatedOn = AddTextBox(tmpName, "txtCreatedOn", "CreatedOn", 2600, 3550, 2500, 360)
    txtCreatedOn.Locked = True

    AddLabel tmpName, U("0627 06CC 062C 0627 062F 0020 06A9 0646 0646 062F 0647"), 500, 4100, 1800, 300
    Dim txtCreatedBy As Control
    Set txtCreatedBy = AddTextBox(tmpName, "txtCreatedBy", "CreatedBy", 2600, 4050, 2500, 360)
    txtCreatedBy.Locked = True

    AddLabel tmpName, U("062A 0627 0631 06CC 062E 0020 0648 06CC 0631 0627 06CC 0634"), 6000, 3600, 1800, 300
    Dim txtModifiedOn As Control
    Set txtModifiedOn = AddTextBox(tmpName, "txtModifiedOn", "ModifiedOn", 7900, 3550, 2500, 360)
    txtModifiedOn.Locked = True

    AddLabel tmpName, U("0648 06CC 0631 0627 06CC 0634 0020 06A9 0646 0646 062F 0647"), 6000, 4100, 1800, 300
    Dim txtModifiedBy As Control
    Set txtModifiedBy = AddTextBox(tmpName, "txtModifiedBy", "ModifiedBy", 7900, 4050, 2500, 360)
    txtModifiedBy.Locked = True

    DoCmd.Save acForm, tmpName
    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename TARGET, acForm, tmpName

    AttachFormCodeAndEvents TARGET, "vba\frmItem.code.vba", Array( _
        "Form.OnBeforeInsert", _
        "Form.OnBeforeUpdate" _
    )
End Sub

Private Sub BuildForm_SubBOMLines()
    Const TARGET As String = "subfrmBOMLines"

    DeleteFormIfExists TARGET

    Dim frm As Form
    Set frm = CreateForm()

    Dim tmpName As String
    tmpName = frm.Name

    frm.RecordSource = "qryBOMLinesWithItem"
    frm.DefaultView = 1
    frm.Caption = U("0627 0642 0644 0627 0645 0020 0042 004F 004D")
    frm.Width = 18000

    AddLabel tmpName, U("0631 062F 06CC 0641"), 300, 50, 900, 250
    AddTextBox tmpName, "txtLineNo", "LineNo", 300, 350, 900, 320

    AddLabel tmpName, U("0642 0637 0639 0647"), 1400, 50, 2200, 250
    Dim cboComp As Control
    Set cboComp = AddComboBox(tmpName, "cboComponentItemID", "ComponentItemID", 1400, 350, 2500, 320)
    cboComp.RowSource = "SELECT i.ItemID, i.ItemCode, i.ItemDescription, i.UOMID FROM tblItems AS i WHERE i.IsActive=True ORDER BY i.ItemCode;"
    cboComp.BoundColumn = 1
    cboComp.ColumnCount = 4
    cboComp.ColumnWidths = "0cm;3cm;6cm;0cm"

    AddLabel tmpName, U("06A9 062F 0020 0642 0637 0639 0647"), 4100, 50, 2200, 250
    Dim txtCompCode As Control
    Set txtCompCode = AddTextBox(tmpName, "txtComponentCode", "ComponentCode", 4100, 350, 2200, 320)
    txtCompCode.Locked = True

    AddLabel tmpName, U("0634 0631 062D 0020 0642 0637 0639 0647"), 6400, 50, 2600, 250
    Dim txtCompDesc As Control
    Set txtCompDesc = AddTextBox(tmpName, "txtComponentDescription", "ComponentDescription", 6400, 350, 3800, 320)
    txtCompDesc.Locked = True

    AddLabel tmpName, U("0645 0635 0631 0641"), 10300, 50, 1200, 250
    AddTextBox tmpName, "txtQtyPer", "QtyPer", 10300, 350, 1200, 320

    AddLabel tmpName, U("0648 0627 062D 062F"), 11650, 50, 1200, 250
    Dim cboUOM As Control
    Set cboUOM = AddComboBox(tmpName, "cboUOMID", "UOMID", 11650, 350, 1600, 320)
    cboUOM.RowSource = "SELECT UOMID, UOMCode FROM tblUOM WHERE IsActive=True ORDER BY UOMCode;"
    cboUOM.BoundColumn = 1
    cboUOM.ColumnCount = 2
    cboUOM.ColumnWidths = "0cm;2.5cm"

    AddLabel tmpName, U("0636 0627 06CC 0639 0627 062A 0020 0025"), 13400, 50, 1200, 250
    AddTextBox tmpName, "txtScrapPct", "ScrapPct", 13400, 350, 1200, 320

    AddLabel tmpName, U("0627 0632 0020 062A 0627 0631 06CC 062E"), 14700, 50, 1200, 250
    AddTextBox tmpName, "dtEffectiveFrom", "EffectiveFrom", 14700, 350, 1500, 320

    AddLabel tmpName, U("062A 0627 0020 062A 0627 0631 06CC 062E"), 16300, 50, 1200, 250
    AddTextBox tmpName, "dtEffectiveTo", "EffectiveTo", 16300, 350, 1500, 320

    DoCmd.Save acForm, tmpName
    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename TARGET, acForm, tmpName

    AttachFormCodeAndEvents TARGET, "vba\subfrmBOMLines.code.vba", Array( _
        "Form.OnBeforeInsert", _
        "Form.OnBeforeUpdate", _
        "Form.OnError", _
        "cboComponentItemID.OnAfterUpdate" _
    )
End Sub

Private Sub BuildForm_BOM()
    Const TARGET As String = "frmBOM"

    DeleteFormIfExists TARGET

    Dim frm As Form
    Set frm = CreateForm()

    Dim tmpName As String
    tmpName = frm.Name

    frm.RecordSource = "tblBOMHeader"
    frm.DefaultView = 0
    frm.Caption = U("0633 0631 0622 06CC 0646 062F 0020 0042 004F 004D")
    frm.Width = 19000

    AddLabel tmpName, U("06A9 0627 0644 0627 06CC 0020 0646 0647 0627 06CC 06CC"), 500, 500, 1500, 300
    Dim cboFG As Control
    Set cboFG = AddComboBox(tmpName, "cboFGItemID", "FGItemID", 2200, 450, 5000, 360)
    cboFG.RowSource = "SELECT i.ItemID, i.ItemCode, i.ItemDescription FROM tblItems AS i " & _
                      "INNER JOIN tblItemType AS t ON i.ItemTypeID=t.ItemTypeID " & _
                      "WHERE i.IsActive=True AND t.TypeCode IN ('FG','SA') ORDER BY i.ItemCode;"
    cboFG.BoundColumn = 1
    cboFG.ColumnCount = 3
    cboFG.ColumnWidths = "0cm;3cm;7cm"

    AddLabel tmpName, U("0634 0645 0627 0631 0647 0020 0646 0633 062E 0647"), 7600, 500, 1400, 300
    Dim txtVerNo As Control
    Set txtVerNo = AddTextBox(tmpName, "txtVersionNo", "VersionNo", 9100, 450, 1000, 360)
    txtVerNo.Locked = True

    AddLabel tmpName, U("0646 0633 062E 0647"), 10300, 500, 1000, 300
    Dim txtVer As Control
    Set txtVer = AddTextBox(tmpName, "txtVersionLabel", "VersionLabel", 11400, 450, 1000, 360)
    txtVer.Locked = True

    AddLabel tmpName, U("0641 0639 0627 0644"), 12600, 500, 800, 300
    Dim chkActive As Control
    Set chkActive = AddCheckBox(tmpName, "chkIsActive", "IsActive", 13400, 450, 500, 360)
    chkActive.Locked = True

    Dim btnSave As Control
    Set btnSave = AddButton(tmpName, "cmdSave", U("0630 062E 06CC 0631 0647"), 500, 1000, 1600, 450)

    Dim btnActivate As Control
    Set btnActivate = AddButton(tmpName, "cmdActivateBOM", U("0641 0639 0627 0644 200C 0633 0627 0632 06CC 0020 0042 004F 004D"), 2200, 1000, 2200, 450)

    Dim btnCopy As Control
    Set btnCopy = AddButton(tmpName, "cmdCopyBOM", U("06A9 067E 06CC 0020 0042 004F 004D"), 4500, 1000, 1800, 450)

    Dim btnPrint As Control
    Set btnPrint = AddButton(tmpName, "cmdPrintBOM", U("0686 0627 067E 0020 0042 004F 004D"), 6400, 1000, 1800, 450)

    Dim sfm As Control
    Set sfm = CreateControl(tmpName, acSubform, acDetail, "", "", 500, 1800, 18000, 7000)
    sfm.Name = "sfmBOMLines"
    sfm.SourceObject = "Form.subfrmBOMLines"
    sfm.LinkMasterFields = "BOMHeaderID"
    sfm.LinkChildFields = "BOMHeaderID"

    DoCmd.Save acForm, tmpName
    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename TARGET, acForm, tmpName

    AttachFormCodeAndEvents TARGET, "vba\frmBOM.code.vba", Array( _
        "Form.OnCurrent", _
        "Form.OnBeforeInsert", _
        "cboFGItemID.OnAfterUpdate", _
        "cmdSave.OnClick", _
        "cmdActivateBOM.OnClick", _
        "cmdCopyBOM.OnClick", _
        "cmdPrintBOM.OnClick" _
    )
End Sub

Private Sub BuildForm_MainMenu()
    Const TARGET As String = "frmMainMenu"

    DeleteFormIfExists TARGET

    Dim frm As Form
    Set frm = CreateForm()

    Dim tmpName As String
    tmpName = frm.Name

    frm.RecordSource = ""
    frm.DefaultView = 0
    frm.Caption = U("0645 0646 0648 06CC 0020 0627 0635 0644 06CC 0020 0633 06CC 0633 062A 0645 0020 0042 004F 004D")
    frm.Width = 9000
    frm.Section(acDetail).BackColor = RGB(245, 248, 252)

    Dim lblTitle As Control
    Set lblTitle = AddLabel(tmpName, U("0633 06CC 0633 062A 0645 0020 0645 062F 06CC 0631 06CC 062A 0020 0042 004F 004D"), 500, 500, 5000, 500)
    lblTitle.FontSize = 16
    lblTitle.FontBold = True

    Dim btnItem As Control
    Set btnItem = AddButton(tmpName, "cmdOpenItem", U("062A 0639 0631 06CC 0641 0020 06A9 0627 0644 0627"), 500, 1400, 2200, 500)

    Dim btnBOM As Control
    Set btnBOM = AddButton(tmpName, "cmdOpenBOM", U("0645 062F 06CC 0631 06CC 062A 0020 0042 004F 004D"), 2800, 1400, 2200, 500)

    Dim btnRpt As Control
    Set btnRpt = AddButton(tmpName, "cmdOpenReport", U("0686 0627 067E 0020 06AF 0632 0627 0631 0634 0020 0042 004F 004D"), 5100, 1400, 2200, 500)

    Dim btnClose As Control
    Set btnClose = AddButton(tmpName, "cmdCloseMenu", U("062E 0631 0648 062C"), 500, 2100, 2200, 500)

    DoCmd.Save acForm, tmpName
    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename TARGET, acForm, tmpName

    AttachFormCodeAndEvents TARGET, "vba\frmMainMenu.code.vba", Array( _
        "cmdOpenItem.OnClick", _
        "cmdOpenBOM.OnClick", _
        "cmdOpenReport.OnClick", _
        "cmdCloseMenu.OnClick" _
    )
End Sub

Private Sub BuildReport_BOM()
    Const TARGET As String = "rptBOM"

    DeleteReportIfExists TARGET

    Dim rpt As Report
    Set rpt = CreateReport()

    Dim tmpName As String
    tmpName = rpt.Name

    rpt.RecordSource = "qryBOMPrintDataset"
    rpt.Caption = U("06AF 0632 0627 0631 0634 0020 0042 004F 004D")
    rpt.Width = 19000

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 500, 300, 2000, 300
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("06A9 062F 0020 06A9 0627 0644 0627 06CC 0020 0646 0647 0627 06CC 06CC")
    AddReportTextBox tmpName, "txtFGCode", "FGCode", acPageHeader, 2600, 300, 2000, 300

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 5000, 300, 2200, 300
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0634 0631 062D 0020 06A9 0627 0644 0627 06CC 0020 0646 0647 0627 06CC 06CC")
    AddReportTextBox tmpName, "txtFGDesc", "FGDescription", acPageHeader, 7300, 300, 5000, 300

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 500, 700, 1600, 300
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0646 0633 062E 0647")
    AddReportTextBox tmpName, "txtVersion", "VersionLabel", acPageHeader, 2600, 700, 1400, 300

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 4300, 700, 1200, 300
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0648 0636 0639 06CC 062A")
    AddReportTextBox tmpName, "txtStatus", U("003D 0049 0049 0066 0028 005B 0049 0073 0041 0063 0074 0069 0076 0065 005D 002C 0022 0641 0639 0627 0644 0022 002C 0022 063A 06CC 0631 0641 0639 0627 0644 0022 0029"), acPageHeader, 5600, 700, 1600, 300

    AddReportTextBox tmpName, "txtPrintDate", "=Now()", acPageHeader, 7600, 700, 2600, 300
    AddReportTextBox tmpName, "txtUser", "=CurrentUser()", acPageHeader, 10400, 700, 2600, 300

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 500, 1150, 700, 250
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0633 0637 062D")
    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 1300, 1150, 2500, 250
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("06A9 062F 0020 0642 0637 0639 0647")
    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 3900, 1150, 3000, 250
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0634 0631 062D 0020 0642 0637 0639 0647")
    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 9200, 1150, 1200, 250
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0645 0635 0631 0641")
    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 10500, 1150, 1000, 250
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0648 0627 062D 062F")
    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 11600, 1150, 1000, 250
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0636 0627 06CC 0639 0627 062A 0020 0025")
    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 12700, 1150, 1300, 250
    rpt.Controls(rpt.Controls.Count - 1).Caption = U("0645 0642 062F 0627 0631 0020 062A 062C 0645 0639 06CC")

    AddReportTextBox tmpName, "txtLevelNo", "LevelNo", acDetail, 500, 400, 700, 300
    AddReportTextBox tmpName, "txtComponentCode", "ComponentCode", acDetail, 1300, 400, 2500, 300
    AddReportTextBox tmpName, "txtComponentDescription", "ComponentDescription", acDetail, 3900, 400, 5000, 300
    AddReportTextBox tmpName, "txtQtyPer", "QtyPer", acDetail, 9200, 400, 1200, 300
    AddReportTextBox tmpName, "txtUOMCode", "UOMCode", acDetail, 10500, 400, 1000, 300
    AddReportTextBox tmpName, "txtScrapPct", "ScrapPct", acDetail, 11600, 400, 1000, 300
    AddReportTextBox tmpName, "txtExtQty", "ExtQty", acDetail, 12700, 400, 1300, 300

    DoCmd.Save acReport, tmpName
    DoCmd.Close acReport, tmpName, acSaveYes
    DoCmd.Rename TARGET, acReport, tmpName

    AttachReportCodeAndEvents TARGET, "vba\rptBOM.code.vba", Array( _
        "Detail.OnFormat" _
    )
End Sub

Private Sub AttachFormCodeAndEvents(ByVal formName As String, ByVal codeRelativePath As String, ByVal eventBindings As Variant)
    On Error Resume Next
    DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0

    DoCmd.OpenForm formName, acDesign

    Dim frm As Form
    Set frm = Forms(formName)
    frm.HasModule = True

    Dim codeText As String
    codeText = NormalizeCodeText(ReadTextFile(CurrentProject.Path & "\" & codeRelativePath))
    If Len(codeText) > 0 Then
        frm.Module.AddFromString codeText
    End If

    Dim i As Long
    For i = LBound(eventBindings) To UBound(eventBindings)
        BindEventToken frm, CStr(eventBindings(i))
    Next i

    DoCmd.Close acForm, formName, acSaveYes
End Sub

Private Sub AttachReportCodeAndEvents(ByVal reportName As String, ByVal codeRelativePath As String, ByVal eventBindings As Variant)
    On Error Resume Next
    DoCmd.Close acReport, reportName, acSaveNo
    On Error GoTo 0

    DoCmd.OpenReport reportName, acViewDesign

    Dim rpt As Report
    Set rpt = Reports(reportName)
    rpt.HasModule = True

    Dim codeText As String
    codeText = NormalizeCodeText(ReadTextFile(CurrentProject.Path & "\" & codeRelativePath))
    If Len(codeText) > 0 Then
        rpt.Module.AddFromString codeText
    End If

    Dim i As Long
    For i = LBound(eventBindings) To UBound(eventBindings)
        BindReportEventToken rpt, CStr(eventBindings(i))
    Next i

    DoCmd.Close acReport, reportName, acSaveYes
End Sub

Private Sub BindEventToken(ByVal frm As Form, ByVal token As String)
    Dim parts() As String
    parts = Split(token, ".")
    If UBound(parts) <> 1 Then Exit Sub

    On Error Resume Next
    If StrComp(parts(0), "Form", vbTextCompare) = 0 Then
        CallByName frm, parts(1), VbLet, "[Event Procedure]"
    Else
        CallByName frm.Controls(parts(0)), parts(1), VbLet, "[Event Procedure]"
    End If
    On Error GoTo 0
End Sub

Private Sub BindReportEventToken(ByVal rpt As Report, ByVal token As String)
    Dim parts() As String
    parts = Split(token, ".")
    If UBound(parts) <> 1 Then Exit Sub

    On Error Resume Next
    If StrComp(parts(0), "Detail", vbTextCompare) = 0 Then
        CallByName rpt.Section(acDetail), parts(1), VbLet, "[Event Procedure]"
    Else
        CallByName rpt.Controls(parts(0)), parts(1), VbLet, "[Event Procedure]"
    End If
    On Error GoTo 0
End Sub

Private Function AddLabel(ByVal formName As String, ByVal caption As String, ByVal leftPos As Long, _
                          ByVal topPos As Long, ByVal width As Long, ByVal height As Long) As Control
    Set AddLabel = CreateControl(formName, acLabel, acDetail, "", "", leftPos, topPos, width, height)
    AddLabel.Caption = caption
End Function

Private Function AddTextBox(ByVal formName As String, ByVal controlName As String, ByVal controlSource As String, _
                            ByVal leftPos As Long, ByVal topPos As Long, ByVal width As Long, ByVal height As Long) As Control
    Set AddTextBox = CreateControl(formName, acTextBox, acDetail, "", controlSource, leftPos, topPos, width, height)
    AddTextBox.Name = controlName
End Function

Private Function AddComboBox(ByVal formName As String, ByVal controlName As String, ByVal controlSource As String, _
                             ByVal leftPos As Long, ByVal topPos As Long, ByVal width As Long, ByVal height As Long) As Control
    Set AddComboBox = CreateControl(formName, acComboBox, acDetail, "", controlSource, leftPos, topPos, width, height)
    AddComboBox.Name = controlName
End Function

Private Function AddCheckBox(ByVal formName As String, ByVal controlName As String, ByVal controlSource As String, _
                             ByVal leftPos As Long, ByVal topPos As Long, ByVal width As Long, ByVal height As Long) As Control
    Set AddCheckBox = CreateControl(formName, acCheckBox, acDetail, "", controlSource, leftPos, topPos, width, height)
    AddCheckBox.Name = controlName
End Function

Private Function AddButton(ByVal formName As String, ByVal controlName As String, ByVal caption As String, _
                           ByVal leftPos As Long, ByVal topPos As Long, ByVal width As Long, ByVal height As Long) As Control
    Set AddButton = CreateControl(formName, acCommandButton, acDetail, "", "", leftPos, topPos, width, height)
    AddButton.Name = controlName
    AddButton.Caption = caption
End Function

Private Function AddReportTextBox(ByVal reportName As String, ByVal controlName As String, ByVal controlSource As String, _
                                  ByVal section As AcSection, ByVal leftPos As Long, ByVal topPos As Long, _
                                  ByVal width As Long, ByVal height As Long) As Control
    Set AddReportTextBox = CreateReportControl(reportName, acTextBox, section, "", controlSource, leftPos, topPos, width, height)
    AddReportTextBox.Name = controlName
End Function

Private Sub DeleteFormIfExists(ByVal formName As String)
    On Error Resume Next
    DoCmd.Close acForm, formName, acSaveNo
    If HasForm(formName) Then
        DoCmd.DeleteObject acForm, formName
    End If
    On Error GoTo 0
End Sub

Private Sub DeleteReportIfExists(ByVal reportName As String)
    On Error Resume Next
    DoCmd.Close acReport, reportName, acSaveNo
    If HasReport(reportName) Then
        DoCmd.DeleteObject acReport, reportName
    End If
    On Error GoTo 0
End Sub

Private Function HasForm(ByVal formName As String) As Boolean
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllForms
        If StrComp(ao.Name, formName, vbTextCompare) = 0 Then
            HasForm = True
            Exit Function
        End If
    Next ao
End Function

Private Function HasReport(ByVal reportName As String) As Boolean
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllReports
        If StrComp(ao.Name, reportName, vbTextCompare) = 0 Then
            HasReport = True
            Exit Function
        End If
    Next ao
End Function

Private Function ReadTextFile(ByVal filePath As String) As String
    If Dir(filePath, vbNormal) = "" Then Exit Function

    On Error Resume Next
    ReadTextFile = ReadTextFileWithCharset(filePath, "utf-8")
    If Err.Number = 0 And Len(ReadTextFile) > 0 Then Exit Function
    Err.Clear

    ReadTextFile = ReadTextFileWithCharset(filePath, "unicode")
    If Err.Number = 0 And Len(ReadTextFile) > 0 Then Exit Function
    Err.Clear

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False)
    ReadTextFile = ts.ReadAll
    ts.Close
    On Error GoTo 0
End Function

Private Function ReadTextFileWithCharset(ByVal filePath As String, ByVal charsetName As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Mode = 3
    stm.Charset = charsetName
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileWithCharset = stm.ReadText(-1)
    stm.Close
End Function

Private Function NormalizeCodeText(ByVal codeText As String) As String
    If Len(codeText) = 0 Then Exit Function

    ' Normalize line endings and strip UTF-8 BOM if present.
    codeText = Replace$(codeText, vbCrLf, vbLf)
    codeText = Replace$(codeText, vbCr, vbLf)
    codeText = Replace$(codeText, ChrW$(65279), "")

    Dim lines() As String
    lines = Split(codeText, vbLf)

    Dim i As Long
    Dim outText As String
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then GoTo AppendLine

        If LCase$(Left$(lineText, 9)) = "attribute" Then GoTo ContinueLoop

        ' Any Option statement at file header must be removed before AddFromString
        ' because class modules already contain their own Option lines.
        If LCase$(Left$(lineText, 7)) = "option " Then GoTo ContinueLoop

AppendLine:
        outText = outText & lines(i) & vbCrLf
ContinueLoop:
    Next i

    NormalizeCodeText = outText
End Function
