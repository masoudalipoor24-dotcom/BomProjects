Attribute VB_Name = "modBOMUIBuilder"
Option Compare Database
Option Explicit

Public Sub AutoSetupAllWithUI()
    On Error GoTo EH

    AutoSetupBOM True
    BuildBOMUIObjects

    MsgBox "BOM backend + UI setup completed.", vbInformation
    Exit Sub
EH:
    MsgBox "AutoSetupAllWithUI failed: " & Err.Description, vbCritical
End Sub

Public Sub BuildBOMUIObjects()
    On Error GoTo EH

    BuildForm_SubBOMLines
    BuildForm_Item
    BuildForm_BOM
    BuildReport_BOM

    Exit Sub
EH:
    MsgBox "BuildBOMUIObjects failed: " & Err.Description, vbCritical
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
    frm.Caption = "Item Master"
    frm.Width = 12000

    AddLabel tmpName, "Item Code", 500, 500, 1800, 300
    AddTextBox tmpName, "txtItemCode", "ItemCode", 2600, 450, 3000, 360

    AddLabel tmpName, "Description", 500, 1000, 1800, 300
    AddTextBox tmpName, "txtItemDescription", "ItemDescription", 2600, 950, 6000, 360

    AddLabel tmpName, "UOM", 500, 1500, 1800, 300
    Dim cboUOM As Control
    Set cboUOM = AddComboBox(tmpName, "cboUOMID", "UOMID", 2600, 1450, 2800, 360)
    cboUOM.RowSource = "SELECT UOMID, UOMCode, UOMName FROM tblUOM WHERE IsActive=True ORDER BY UOMCode;"
    cboUOM.BoundColumn = 1
    cboUOM.ColumnCount = 3
    cboUOM.ColumnWidths = "0cm;2.5cm;4cm"

    AddLabel tmpName, "Item Type", 500, 2000, 1800, 300
    Dim cboType As Control
    Set cboType = AddComboBox(tmpName, "cboItemTypeID", "ItemTypeID", 2600, 1950, 2800, 360)
    cboType.RowSource = "SELECT ItemTypeID, TypeCode, TypeName FROM tblItemType ORDER BY TypeCode;"
    cboType.BoundColumn = 1
    cboType.ColumnCount = 3
    cboType.ColumnWidths = "0cm;2.5cm;4cm"

    AddLabel tmpName, "Active", 500, 2500, 1800, 300
    AddCheckBox tmpName, "chkIsActive", "IsActive", 2600, 2450, 500, 360

    AddLabel tmpName, "Std Cost", 500, 3000, 1800, 300
    AddTextBox tmpName, "txtStdCost", "StdCost", 2600, 2950, 1800, 360

    AddLabel tmpName, "Created On", 500, 3600, 1800, 300
    Dim txtCreatedOn As Control
    Set txtCreatedOn = AddTextBox(tmpName, "txtCreatedOn", "CreatedOn", 2600, 3550, 2500, 360)
    txtCreatedOn.Locked = True

    AddLabel tmpName, "Created By", 500, 4100, 1800, 300
    Dim txtCreatedBy As Control
    Set txtCreatedBy = AddTextBox(tmpName, "txtCreatedBy", "CreatedBy", 2600, 4050, 2500, 360)
    txtCreatedBy.Locked = True

    AddLabel tmpName, "Modified On", 6000, 3600, 1800, 300
    Dim txtModifiedOn As Control
    Set txtModifiedOn = AddTextBox(tmpName, "txtModifiedOn", "ModifiedOn", 7900, 3550, 2500, 360)
    txtModifiedOn.Locked = True

    AddLabel tmpName, "Modified By", 6000, 4100, 1800, 300
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
    frm.DefaultView = 2
    frm.Caption = "BOM Lines"
    frm.Width = 18000

    AddTextBox tmpName, "txtLineNo", "LineNo", 300, 300, 900, 320

    Dim cboComp As Control
    Set cboComp = AddComboBox(tmpName, "cboComponentItemID", "ComponentItemID", 1400, 300, 2500, 320)
    cboComp.RowSource = "SELECT i.ItemID, i.ItemCode, i.ItemDescription, i.UOMID FROM tblItems AS i WHERE i.IsActive=True ORDER BY i.ItemCode;"
    cboComp.BoundColumn = 1
    cboComp.ColumnCount = 4
    cboComp.ColumnWidths = "0cm;3cm;6cm;0cm"

    Dim txtCompCode As Control
    Set txtCompCode = AddTextBox(tmpName, "txtComponentCode", "ComponentCode", 4100, 300, 2200, 320)
    txtCompCode.Locked = True

    Dim txtCompDesc As Control
    Set txtCompDesc = AddTextBox(tmpName, "txtComponentDescription", "ComponentDescription", 6400, 300, 3800, 320)
    txtCompDesc.Locked = True

    AddTextBox tmpName, "txtQtyPer", "QtyPer", 10300, 300, 1200, 320

    Dim cboUOM As Control
    Set cboUOM = AddComboBox(tmpName, "cboUOMID", "UOMID", 11650, 300, 1600, 320)
    cboUOM.RowSource = "SELECT UOMID, UOMCode FROM tblUOM WHERE IsActive=True ORDER BY UOMCode;"
    cboUOM.BoundColumn = 1
    cboUOM.ColumnCount = 2
    cboUOM.ColumnWidths = "0cm;2.5cm"

    AddTextBox tmpName, "txtScrapPct", "ScrapPct", 13400, 300, 1200, 320
    AddTextBox tmpName, "dtEffectiveFrom", "EffectiveFrom", 14700, 300, 1500, 320
    AddTextBox tmpName, "dtEffectiveTo", "EffectiveTo", 16300, 300, 1500, 320

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
    frm.Caption = "BOM Header"
    frm.Width = 19000

    AddLabel tmpName, "FG Item", 500, 500, 1500, 300
    Dim cboFG As Control
    Set cboFG = AddComboBox(tmpName, "cboFGItemID", "FGItemID", 2200, 450, 5000, 360)
    cboFG.RowSource = "SELECT i.ItemID, i.ItemCode, i.ItemDescription FROM tblItems AS i " & _
                      "INNER JOIN tblItemType AS t ON i.ItemTypeID=t.ItemTypeID " & _
                      "WHERE i.IsActive=True AND t.TypeCode IN ('FG','SA') ORDER BY i.ItemCode;"
    cboFG.BoundColumn = 1
    cboFG.ColumnCount = 3
    cboFG.ColumnWidths = "0cm;3cm;7cm"

    AddLabel tmpName, "Version No", 7600, 500, 1400, 300
    Dim txtVerNo As Control
    Set txtVerNo = AddTextBox(tmpName, "txtVersionNo", "VersionNo", 9100, 450, 1000, 360)
    txtVerNo.Locked = True

    AddLabel tmpName, "Version", 10300, 500, 1000, 300
    Dim txtVer As Control
    Set txtVer = AddTextBox(tmpName, "txtVersionLabel", "VersionLabel", 11400, 450, 1000, 360)
    txtVer.Locked = True

    AddLabel tmpName, "Active", 12600, 500, 800, 300
    Dim chkActive As Control
    Set chkActive = AddCheckBox(tmpName, "chkIsActive", "IsActive", 13400, 450, 500, 360)
    chkActive.Locked = True

    Dim btnSave As Control
    Set btnSave = AddButton(tmpName, "cmdSave", "Save", 500, 1000, 1600, 450)

    Dim btnActivate As Control
    Set btnActivate = AddButton(tmpName, "cmdActivateBOM", "Activate BOM", 2200, 1000, 2200, 450)

    Dim btnCopy As Control
    Set btnCopy = AddButton(tmpName, "cmdCopyBOM", "Copy BOM", 4500, 1000, 1800, 450)

    Dim btnPrint As Control
    Set btnPrint = AddButton(tmpName, "cmdPrintBOM", "Print BOM", 6400, 1000, 1800, 450)

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

Private Sub BuildReport_BOM()
    Const TARGET As String = "rptBOM"

    DeleteReportIfExists TARGET

    Dim rpt As Report
    Set rpt = CreateReport()

    Dim tmpName As String
    tmpName = rpt.Name

    rpt.RecordSource = "qryBOMPrintDataset"
    rpt.Caption = "BOM Report"
    rpt.Width = 19000

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 500, 300, 2000, 300
    rpt.Controls(rpt.Controls.Count - 1).Caption = "FG Code"
    AddReportTextBox tmpName, "txtFGCode", "FGCode", acPageHeader, 2600, 300, 2000, 300

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 5000, 300, 2200, 300
    rpt.Controls(rpt.Controls.Count - 1).Caption = "FG Description"
    AddReportTextBox tmpName, "txtFGDesc", "FGDescription", acPageHeader, 7300, 300, 5000, 300

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 500, 700, 1600, 300
    rpt.Controls(rpt.Controls.Count - 1).Caption = "Version"
    AddReportTextBox tmpName, "txtVersion", "VersionLabel", acPageHeader, 2600, 700, 1400, 300

    CreateReportControl tmpName, acLabel, acPageHeader, "", "", 4300, 700, 1200, 300
    rpt.Controls(rpt.Controls.Count - 1).Caption = "Status"
    AddReportTextBox tmpName, "txtStatus", "=IIf([IsActive],""Active"",""Inactive"")", acPageHeader, 5600, 700, 1600, 300

    AddReportTextBox tmpName, "txtPrintDate", "=Now()", acPageHeader, 7600, 700, 2600, 300
    AddReportTextBox tmpName, "txtUser", "=CurrentUser()", acPageHeader, 10400, 700, 2600, 300

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

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False)
    ReadTextFile = ts.ReadAll
    ts.Close
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
