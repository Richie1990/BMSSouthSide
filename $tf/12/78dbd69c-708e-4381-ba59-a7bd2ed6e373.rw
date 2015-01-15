Public Class frmLandlord

    ' Dimension Form Variables
    Dim objLandlord As New clsLandlord
    Dim bDataChanged As Boolean

    Private Sub frmLandlord_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If bDataChanged = True Then
            If MsgBox("Data not Saved - Are you Sure ?", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "BMSSortmyPC - User Information") = vbCancel Then
                e.Cancel = True
            End If
        End If

    End Sub

    Private Sub frmLandlord_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Convert.ToChar(27) Then
            Me.Close()
        End If

    End Sub

    Private Sub frmLandlord_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadLandlordTree()
        '        LoadLandlordJobList()
        bDataChanged = False

    End Sub

    Public Sub DisplayLandlord()
        '** Display Landlord Data

        ' Error Checking
        On Error GoTo Err_DisplayLandlord


        ' Display Record
        With objLandlord
            .LandlordID = Val(lblLandlordIDValue.Text)
            .LoadLandlord()
            txtTitle1.Text = .Title1
            txtForename1.Text = .Forename1
            txtSurname1.Text = .Surname1
            dtpDOB1.Value = .DOB1
            txtPlaceofBirth2.Text = .PlaceofBirth2
            txtTitle2.Text = .Title2
            txtForename2.Text = .Forename2
            txtSurname2.Text = .Surname2
            dtpDOB2.Value = .DOB2
            txtPlaceofBirth2.Text = .PlaceofBirth2
            txtCommsTitle.Text = .CommsTitle
            txtLandlordName.Text = .LandlordName
            txtAddress1.Text = .Address1
            txtAddress2.Text = .Address2
            txtAddress3.Text = .Address3
            txtTown.Text = .Town
            txtPostCode.Text = .PostCode
            txtPhoneNo.Text = .PhoneNo
            txtMobileNo.Text = .MobileNo
            txtEMailAddress.Text = .EMailAddress
            rtbLandlordNotes.Text = .LandlordNotes
            For iRCnt = 0 To cmbSource.Items.Count - 1
                If cmbSource.Items.Item(iRCnt).ToString = .SourceDescription Then
                    cmbSource.SelectedIndex = iRCnt
                End If
            Next
            '            optLStatus.Checked = .LStatus
        End With

Err_DisplayLandlord:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "DisplayLandlord", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "DisplayLandlord" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub tvwLandlordListBusiness_AfterSelect(sender As Object, e As TreeViewEventArgs)
        '** Load Data Records

        ' Error Checking
        On Error GoTo Err_tvwLandlordListBusiness_AfterSelect

        ' Dimension Local Variables
        Dim sNode, sNodeType As String

        ' Get Data from Node
        sNode = tvwLandlordList.SelectedNode.Tag.ToString

        ' Set Menus & Buttons
        mnuNewLandlord.Enabled = True


        lblLandlordIDValue.Text = sNode.Substring(2, 4)
        DisplayLandlord()
        lblSource.Visible = False
        cmbSource.Visible = True

        pnlLandlord.Visible = True
        pnlLandlord.Enabled = True
        txtAddress1.Enabled = False
        txtAddress2.Enabled = False
        txtAddress3.Enabled = False
        txtTown.Enabled = False
        txtPostCode.Enabled = False
        lblSource.Visible = False
        cmbSource.Visible = False
        LoadLandlordJobList()
        tsbSaveLandlord.Enabled = True
        mnuSaveLandlord.Enabled = True
        tsbRemoveLandlord.Enabled = True
        mnuRemoveLandlord.Enabled = True
        tsbNewProperty.Enabled = True
        mnuJob.Enabled = True
        tsbEditProperty.Enabled = False
        WriteAuditLogRecord(Me.Name, "tvwLandlordList_AfterSelect", "ACTION", "Load Landlord" & sNode)


Err_tvwLandlordListBusiness_AfterSelect:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tvwLandlordListBusiness_AfterSelect", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tvwLandlordListBusiness_AfterSelect" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub


    Private Sub mnuNewLandlord_Click(sender As Object, e As EventArgs) Handles mnuNewLandlord.Click
        tsbNewLandlord.PerformClick()

    End Sub


    Sub LoadLandlordTree()
        '**Loads Landlord List

        ' Error Checking
        On Error GoTo Err_LoadLandlordTree

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter
        Dim uNode As TreeNode
        Dim iLandlordID As Integer = 0
        Dim iLandlordCnt As Integer = 0
        Dim iContactID As Integer = 0
        Dim iContactCnt As Integer = 0

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Landlord List (Based on Search Value)
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SearchValue", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@SearchValue").Value = txtFilter.Text
            .CommandText = "Landlord_LoadRecords"
            uRecSnap = .Execute
        End With

        ' Suppress TreeView Repaint / Clear TreeView
        tvwLandlordList.BeginUpdate()
        tvwLandlordList.Nodes.Clear()
        tvwLandlordList.ShowNodeToolTips = True

        ' Populate List
        Do Until uRecSnap.EOF
            uNode = tvwLandlordList.Nodes.Add("L" & Format(uRecSnap("LandlordID").Value, "0000"), uRecSnap("LandlordName").Value)
            uNode.Tag = "L:" & Format(uRecSnap("LandlordID").Value, "0000") & ":01:"
            uRecSnap.MoveNext()
        Loop
        uRecSnap.Close()

        ' Repaint TreeView.
        tvwLandlordList.EndUpdate()
        tvwLandlordList.Refresh()

        ' Close Connection
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_LoadLandlordTree:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "LoadLandlordTree", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadLandlordTree" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Sub LoadLandlordJobList()
        '**Loads Landlord Job List

        ' Error Checking
        On Error GoTo Err_LoadLandlordJobList

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Landlord Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@LandlordID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LandlordID").Value = Val(lblLandlordIDValue.Text)
            .CommandText = "Landlord_LoadJobList"
            uRecSnap = .Execute
        End With


        ' Populate List

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_LoadLandlordJobList:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "LoadLandlordJobList", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadLandlordJobList" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub


    Private Sub mnuFileClose_Click(sender As Object, e As EventArgs) Handles mnuFileClose.Click
        Me.Close()

    End Sub

    Private Sub mnuSaveLandlord_Click(sender As Object, e As EventArgs) Handles mnuSaveLandlord.Click
        tsbSaveLandlord.PerformClick()

    End Sub

    Private Sub mnuRemoveLandlord_Click(sender As Object, e As EventArgs) Handles mnuRemoveLandlord.Click
        tsbRemoveLandlord.PerformClick()

    End Sub

    Private Sub txtForename_TextChanged(sender As Object, e As EventArgs)
        tsbSaveLandlord.Enabled = True
        mnuSaveLandlord.Enabled = True

    End Sub





    Private Sub tsbNewLandlord_Click(sender As Object, e As EventArgs) Handles tsbNewLandlord.Click
        ' Clear Data
        lblLandlordIDValue.Text = "0000"
        txtTitle1.Text = ""
        txtForename1.Text = ""
        txtSurname1.Text = ""
        txtAddress1.Text = ""
        txtAddress1.Enabled = True
        txtAddress2.Text = ""
        txtAddress2.Enabled = True
        txtAddress3.Text = ""
        txtAddress3.Enabled = True
        txtTown.Text = ""
        txtTown.Enabled = True
        txtPostCode.Text = ""
        txtPostCode.Enabled = True
        txtPhoneNo.Text = ""
        txtMobileNo.Text = ""
        txtEMailAddress.Text = ""
        cmbSource.SelectedIndex = 0

        rtbLandlordNotes.Text = ""
        rtbLandlordNotes.Enabled = True


        tsbNewLandlord.Enabled = False
        mnuNewLandlord.Enabled = False
        tsbSaveLandlord.Enabled = True
        mnuSaveLandlord.Enabled = True
        tsbRemoveLandlord.Enabled = False
        mnuRemoveLandlord.Enabled = False
        pnlLandlord.Enabled = True


    End Sub

    Private Sub tsbClose_Click(sender As Object, e As EventArgs) Handles tsbClose.Click
        Me.Close()

    End Sub


    Private Sub tsbSaveLandlord_Click(sender As Object, e As EventArgs) Handles tsbSaveLandlord.Click
        '** Save Landlord Data

        ' Error Checking
        On Error GoTo Err_SaveLandlord

        ' Dimension Local Variables
        Dim bErr As Boolean
        Dim bSaved As Boolean

        ' Clear Previous Errors
        For Each ControlChild In Me.Controls
            If TypeOf ControlChild Is Label Then
                ControlChild.forecolor = Color.Black
            End If
        Next

        ' Check Data
        If txtForename1.Text = "" Then
            bErr = True
            lblForenameTitle1.ForeColor = Color.Red
        Else
            lblForenameTitle1.ForeColor = Color.Black
        End If
        If txtSurname1.Text = "" Then
            bErr = True
            lblSurname1.ForeColor = Color.Red
        Else
            lblSurname1.ForeColor = Color.Black
        End If
        If cmbSource.SelectedIndex < 0 Then
            bErr = True
            lblSource.ForeColor = Color.Red
        Else
            lblSource.ForeColor = Color.Black
        End If


        ' Check for Errors
        If bErr Then
            lblMessagePanel.Text = "** Warning - Errors in Data" & Chr(13) & "Please refer to red boxes above"
            lblMessagePanel.Visible = True

        Else

            With objLandlord
                .LandlordID = Val(lblLandlordIDValue.Text)
                .Title1 = txtTitle1.Text
                .Forename1 = txtForename1.Text
                .Surname1 = txtSurname1.Text
                .DOB1 = dtpDOB1.Value
                .PlaceofBirth1 = txtPlaceofBirth1.Text
                .Title2 = txtTitle2.Text
                .Forename2 = txtForename2.Text
                .Surname2 = txtSurname2.Text
                .DOB2 = dtpDOB2.Value
                .PlaceofBirth2 = txtPlaceofBirth2.Text
                .CommsTitle = txtCommsTitle.Text
                .LandlordName = txtLandlordName.Text
                .Address1 = txtAddress1.Text
                .Address2 = txtAddress2.Text
                .Address3 = txtAddress3.Text
                .Town = txtTown.Text
                .PostCode = txtPostCode.Text
                .PhoneNo = txtPhoneNo.Text
                .MobileNo = txtMobileNo.Text
                .EMailAddress = txtEMailAddress.Text
                .CommsStatus = ""
                .LandlordNotes = rtbLandlordNotes.Text
                .SourceID = cmbSource.SelectedValue
                .JoinDate = dtpDateJoined.Value
                .BankName = txtBankName.Text
                .Sortcode = Val(txtSortcode.Text)
                .AccountNo = Val(txtAccountNo.Text)
                .LStatus = True
                bSaved = .SaveLandlord()
            End With


            ' Store Information in Class Object
            If bSaved = True Then
                WriteAuditLogRecord(Me.Name, "SaveLandlordContact", "OK", "Save Landlord" & objLandlord.LandlordID)
                lblLandlordIDValue.Text = Format(objLandlord.LandlordID, "0000")
                lblMessagePanel.Visible = False
                tsbSaveLandlord.Enabled = False
                mnuSaveLandlord.Enabled = False
                tsbNewLandlord.Enabled = True
                mnuNewLandlord.Enabled = True
                pnlLandlord.Enabled = False
                tsbNewProperty.Enabled = False
                mnuJob.Enabled = False
                bDataChanged = False
                MsgBox("Landlord / Contact Data saved successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Contact Information")
                LoadLandlordTree()
            Else
                WriteAuditLogRecord(Me.Name, "SaveLandlord", "FAIL", "Save Landlord")
                txtSurname1.Select()
                MsgBox("Landlord / Contact Update Failed", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Contact Information")
            End If
        End If

Err_SaveLandlord:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "SaveLandlord", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveLandlord" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If


    End Sub

    Private Sub tsbRemoveLandlord_Click(sender As Object, e As EventArgs) Handles tsbRemoveLandlord.Click
        '** Remove Selected Landlord / Contact

        ' Error Checking
        On Error GoTo Err_cmdRemoveCategory_Click

        If MsgBox("Removing Business Landlord: " & txtTitle1.Text & Chr(13) & "** This will also remove all Property Records", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "AztecCRM - Confirmation") = MsgBoxResult.Ok Then
            With objLandlord
                .LandlordID = Val(lblLandlordIDValue.Text)
                .RemoveLandlord()
            End With
            MsgBox("Landlord Removed Successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Confirmation")
        End If


        ' Clear Data
        lblLandlordIDValue.Text = "0000"
        txtTitle1.Text = ""
        txtForename1.Text = ""
        txtSurname1.Text = ""
        txtAddress1.Text = ""
        txtAddress1.Enabled = True
        txtAddress2.Text = ""
        txtAddress2.Enabled = True
        txtAddress3.Text = ""
        txtAddress3.Enabled = True
        txtTown.Text = ""
        txtTown.Enabled = True
        txtPostCode.Text = ""
        txtPostCode.Enabled = True
        txtPhoneNo.Text = ""
        txtMobileNo.Text = ""
        txtEMailAddress.Text = ""
        cmbSource.SelectedIndex = 0
        rtbLandlordNotes.Text = ""
        rtbLandlordNotes.Enabled = True

        tsbNewProperty.Enabled = False
        mnuJob.Enabled = False

        LoadLandlordTree()

Err_cmdRemoveCategory_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdRemoveCategory_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdRemoveCategory_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If


    End Sub





    Private Sub mnuNewJob_Click(sender As Object, e As EventArgs) Handles mnuNewProperty.Click
        tsbNewProperty.PerformClick()
    End Sub

    Private Sub mnuEditJob_Click(sender As Object, e As EventArgs) Handles mnuEditProperty.Click
        tsbEditProperty.PerformClick()
    End Sub

    Private Sub tsbNewContact_Click(sender As Object, e As EventArgs)
        ' Clear Data
        txtTitle1.Text = ""
        txtForename1.Text = ""
        txtSurname1.Text = ""
        txtAddress1.Text = ""
        txtAddress1.Enabled = False
        txtAddress2.Text = ""
        txtAddress2.Enabled = False
        txtAddress3.Text = ""
        txtAddress3.Enabled = False
        txtTown.Text = ""
        txtTown.Enabled = False
        txtPostCode.Text = ""
        txtPostCode.Enabled = False
        txtPhoneNo.Text = ""
        txtMobileNo.Text = ""
        txtEMailAddress.Text = ""
        cmbSource.Visible = False
        cmbSource.SelectedIndex = 0
        rtbLandlordNotes.Text = ""
        rtbLandlordNotes.Enabled = False

        tsbNewLandlord.Enabled = False
        mnuNewLandlord.Enabled = False
        tsbSaveLandlord.Enabled = True
        mnuSaveLandlord.Enabled = True
        tsbRemoveLandlord.Enabled = False
        mnuRemoveLandlord.Enabled = False
        pnlLandlord.Visible = True
        pnlLandlord.Enabled = True

    End Sub



    Private Sub pnlBusiness_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub txtFilter_TextChanged(sender As Object, e As EventArgs) Handles txtFilter.TextChanged
        LoadLandlordTree()

    End Sub



    Private Sub lblFilter_Click(sender As Object, e As EventArgs) Handles lblFilter.Click

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        PopulateCombo(cmbSource)

    End Sub

    Private Sub txtCompany_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub txtCAddress1_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub txtCAddress2_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub txtCAddress3_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub txtCAddress4_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub txtCPostcode_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub txtCPhoneNo_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub txtWebsite_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub cmbCSource_SelectedIndexChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub cmbCLandlordType_SelectedIndexChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub rtbCLandlordNotes_TextChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub

    Private Sub txtTitle_TextChanged(sender As Object, e As EventArgs) Handles txtTitle1.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtForename_TextChanged_1(sender As Object, e As EventArgs) Handles txtForename1.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtSurname_TextChanged(sender As Object, e As EventArgs) Handles txtSurname1.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtAddress1_TextChanged(sender As Object, e As EventArgs) Handles txtAddress1.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtAddress2_TextChanged(sender As Object, e As EventArgs) Handles txtAddress2.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtAddress3_TextChanged(sender As Object, e As EventArgs) Handles txtAddress3.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtAddress4_TextChanged(sender As Object, e As EventArgs) Handles txtTown.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtPostCode_TextChanged(sender As Object, e As EventArgs) Handles txtPostCode.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtPhoneNo_TextChanged(sender As Object, e As EventArgs) Handles txtPhoneNo.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtMobileNo_TextChanged(sender As Object, e As EventArgs) Handles txtMobileNo.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtEMailAddress_TextChanged(sender As Object, e As EventArgs) Handles txtEMailAddress.TextChanged
        bDataChanged = True

    End Sub

    Private Sub cmbSource_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSource.SelectedIndexChanged
        bDataChanged = True

    End Sub

    Private Sub cmbLandlordType_SelectedIndexChanged(sender As Object, e As EventArgs)
        bDataChanged = True

    End Sub


    Private Sub tvwLandlordList_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvwLandlordList.AfterSelect

        ' Error Checking
        On Error GoTo Err_tvwLandlordList_AfterSelect

        ' Dimension Local Variables
        Dim sNode, sNodeType As String

        ' Get Data from Node
        sNode = tvwLandlordList.SelectedNode.Tag.ToString

        lblLandlordIDValue.Text = sNode.Substring(2, 4)
        DisplayLandlord()
        pnlLandlord.Enabled = True
        mnuSaveLandlord.Enabled = True
        tsbSaveLandlord.Enabled = True
        bDataChanged = False

Err_tvwLandlordList_AfterSelect:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tvwLandlordList_AfterSelect", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tvwLandlordList_AfterSelect" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub tsbClient_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles tsbClient.ItemClicked

    End Sub
End Class