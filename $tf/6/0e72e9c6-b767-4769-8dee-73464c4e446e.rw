Public Class frmSupplier

    ' Dimension Form Variables
    Dim objSupplier As New clsSupplier
    Dim bDataChanged As Boolean

    Private Sub frmSupplier_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If bDataChanged = True Then
            If MsgBox("Data not Saved - Are you Sure ?", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "BMSSortmyPC - User Information") = vbCancel Then
                e.Cancel = True
            End If
        End If

    End Sub

    Private Sub frmSupplier_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadSupplierTree()
        LoadSupplierOrderList()
        bDataChanged = False

    End Sub

    Private Sub DisplaySupplier()
        '** Display Supplier Data

        ' Error Checking
        On Error GoTo Err_DisplaySupplier

        With objSupplier
            lblSupplierIDValue.Text = Format(.SupplierID, "0000")
            txtCompany.Text = .Company
            txtAddress1.Text = .Address1
            txtAddress2.Text = .Address2
            txtAddress3.Text = .Address3
            txtAddress4.Text = .Address4
            txtPostcode.Text = .PostCode
            txtWebsite.Text = .Website
            rtbSupplierNotes.Text = .SupplierNotes
            optLive.Checked = .CLive
        End With
        bDataChanged = False

Err_DisplaySupplier:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "DisplaySupplier", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "DisplaySupplier" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub DisplaySupplierContact()
        '** Display Supplier Contact Data

        ' Error Checking
        On Error GoTo Err_DisplaySupplierContact

        With objSupplier
            lblSupplierIDValue.Text = Format(.SupplierID, "0000")
            lblContactIDValue.Text = Format(.ContactID, "00")
            txtTitle.Text = .Title
            txtForename.Text = .Forename
            txtSurname.Text = .Surname
            txtPhoneNo.Text = .PhoneNo
            txtMobileNo.Text = .MobileNo
            txtEmailAddress.Text = .EMailAddress
        End With

Err_DisplaySupplierContact:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "DisplaySupplierContact", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "DisplaySupplierContact" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub mnuNewSupplier_Click(sender As Object, e As EventArgs) Handles mnuNewSupplier.Click
        tsbNewSupplier.PerformClick()

    End Sub

    Sub LoadSupplierTree()
        '**Loads Supplier List

        ' Error Checking
        On Error GoTo Err_LoadSupplierTree

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter
        Dim uNode As TreeNode
        Dim iSupplierID As Integer = 0
        Dim iSupplierCnt As Integer = 0
        Dim iContactID As Integer = 0
        Dim iContactCnt As Integer = 0

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Supplier List (Based on Search Value)
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SearchValue", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@SearchValue").Value = txtSearchValue.Text
            .CommandText = "Supplier_LoadRecords"
            uRecSnap = .Execute
        End With

        ' Suppress TreeView Repaint / Clear TreeView
        tvwSupplier.BeginUpdate()
        tvwSupplier.Nodes.Clear()
        tvwSupplier.ShowNodeToolTips = True

        ' Populate List
        Do Until uRecSnap.EOF
            If iSupplierID <> uRecSnap("SupplierID").Value Then
                uNode = tvwSupplier.Nodes.Add("S" & Format(uRecSnap("SupplierID").Value, "0000"), uRecSnap("Company").Value)
                uNode.Tag = "S:" & Format(uRecSnap("SupplierID").Value, "0000") & ":00:"
                iSupplierID = uRecSnap("SupplierID").Value
                iSupplierCnt = iSupplierCnt + 1
                iContactID = 0
                iContactCnt = 0
            End If
            If (IsDBNull(uRecSnap("ContactID").Value)) = False Then
                If iContactID <> uRecSnap("ContactID").Value Then
                    uNode = tvwSupplier.Nodes(iSupplierCnt - 1).Nodes.Add("C" & Format(uRecSnap("ContactID").Value, "00"), uRecSnap("ListName").Value)
                    uNode.Tag = "C:" & Format(uRecSnap("SupplierID").Value, "0000") & ":" & Format(uRecSnap("ContactID").Value, "00") & ":"
                    iContactID = uRecSnap("ContactID").Value
                    iContactCnt = iContactCnt + 1
                End If
            End If
            uRecSnap.MoveNext()
        Loop
        uRecSnap.Close()

        ' Repaint TreeView.
        tvwSupplier.EndUpdate()
        tvwSupplier.Refresh()

        ' Close Connection
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_LoadSupplierTree:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "LoadSupplierTree", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadSupplierTree" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Sub LoadSupplierOrderList()
        '**Loads Supplier Order List

        ' Error Checking
        On Error GoTo Err_LoadSupplierOrderList

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Supplier Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = Val(lblSupplierIDValue.Text)
            .CommandText = "Supplier_LoadOrderList"
            uRecSnap = .Execute
        End With


        ' Populate List
        lstOrders.Items.Clear()
        Do Until uRecSnap.EOF
            lstOrders.Items.Add(Format(uRecSnap("OrderID").Value, "0000") & " - " & uRecSnap("OrderDescription").Value)
            uRecSnap.MoveNext()
        Loop

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_LoadSupplierOrderList:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "LoadSupplierOrderList", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadSupplierOrderList" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub


    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub cmdNewSupplier_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub mnuFileClose_Click(sender As Object, e As EventArgs) Handles mnuFileClose.Click
        Me.Close()

    End Sub

    Private Sub mnuSaveSupplier_Click(sender As Object, e As EventArgs) Handles mnuSaveSupplier.Click
        tsbSaveSupplier.PerformClick()

    End Sub

    Private Sub mnuRemoveSupplier_Click(sender As Object, e As EventArgs) Handles mnuRemoveSupplier.Click
        tsbRemoveSupplier.PerformClick()

    End Sub

    Private Sub txtForename_TextChanged(sender As Object, e As EventArgs)
        tsbSaveSupplier.Enabled = True
        mnuSaveSupplier.Enabled = True

    End Sub

    Private Sub mnuJob_Click(sender As Object, e As EventArgs) Handles mnuOrder.Click
        tsbNewOrder.PerformClick()
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub cmdNewJob_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub cmdEditJob_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub lstOrders_DoubleClick(sender As Object, e As EventArgs) Handles lstOrders.DoubleClick

        ' Store OrderID
        If IsDBNull(lstOrders.SelectedItem) Then
            tsbEditOrder.Enabled = False
            Exit Sub
        End If
        If lstOrders.SelectedItem <> "" Then
            tsbEditOrder.Enabled = True
        End If
        tsbEditOrder.PerformClick()

    End Sub

    Private Sub lstOrders_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstOrders.SelectedIndexChanged

        ' Store OrderID
        If IsDBNull(lstOrders.SelectedItem) Then
            tsbEditOrder.Enabled = False
            Exit Sub
        End If
        If lstOrders.SelectedItem <> "" Then
            tsbEditOrder.Enabled = True
        End If

    End Sub


    Private Sub cmdSearch_Click(sender As Object, e As EventArgs) Handles cmdSearch.Click

    End Sub


    Private Sub tsbNewSupplier_Click(sender As Object, e As EventArgs) Handles tsbNewSupplier.Click
        ' Clear Data
        tvwSupplier.Tag = "S"
        lblSupplierIDValue.Text = "0000"
        lblContactIDValue.Text = "00"
        txtCompany.Text = ""
        txtAddress1.Text = ""
        txtAddress2.Text = ""
        txtAddress3.Text = ""
        txtAddress4.Text = ""
        txtPostcode.Text = ""
        txtPhoneNo.Text = ""
        txtWebsite.Text = ""
        txtTitle.Text = ""
        txtForename.Text = ""
        txtSurname.Text = ""
        txtPhoneNo.Text = ""
        txtMobileNo.Text = ""
        txtEmailAddress.Text = ""
        rtbSupplierNotes.Text = ""

        tsbNewSupplier.Enabled = False
        mnuNewSupplier.Enabled = False
        tsbSaveSupplier.Enabled = True
        mnuSaveSupplier.Enabled = True
        tsbRemoveSupplier.Enabled = False
        mnuRemoveSupplier.Enabled = False
        pnlSupplier.Enabled = True


    End Sub

    Private Sub tsbClose_Click(sender As Object, e As EventArgs) Handles tsbClose.Click
        Me.Close()

    End Sub




    Private Sub tsbSaveSupplier_Click(sender As Object, e As EventArgs) Handles tsbSaveSupplier.Click
        '** Save Supplier Data

        ' Error Checking
        On Error GoTo Err_SaveSupplier

        ' Dimension Local Variables
        Dim bErr As Boolean
        Dim bSaved As Boolean

        ' Clear Previous Errors
        For Each ControlChild In Me.Controls
            If TypeOf ControlChild Is Label Then
                ControlChild.forecolor = Color.Black
            End If
        Next

        Select Case tvwSupplier.Tag
            Case "S"
                If txtCompany.Text = "" Then
                    bErr = True
                    lblCompany.ForeColor = Color.Red
                End If

            Case "C"
                If txtForename.Text = "" Then
                    bErr = True
                    lblForenameTitle.ForeColor = Color.Red
                End If

        End Select

        ' Check for Errors
        If bErr Then
            lblMessagePanel.Text = "** Warning - Errors in Data" & Chr(13) & "Please refer to red boxes above"
            lblMessagePanel.Visible = True

        Else

            Select Case tvwSupplier.Tag
                Case "S"
                    With objSupplier
                        .SupplierID = Val(lblSupplierIDValue.Text)
                        .Company = txtCompany.Text
                        .Address1 = txtAddress1.Text
                        .Address2 = txtAddress2.Text
                        .Address3 = txtAddress3.Text
                        .Address4 = txtAddress4.Text
                        .PostCode = txtPostcode.Text
                        .Website = txtWebsite.Text
                        .SupplierNotes = rtbSupplierNotes.Text
                        .CLive = (optLive.Checked = True)
                        bSaved = .SaveSupplier()
                    End With

                Case "C"
                    With objSupplier
                        .SupplierID = Val(lblSupplierIDValue.Text)
                        .ContactID = Val(lblContactIDValue.Text)
                        .Title = txtTitle.Text
                        .Forename = txtForename.Text
                        .Surname = txtSurname.Text
                        .PhoneNo = txtPhoneNo.Text
                        .MobileNo = txtMobileNo.Text
                        .EMailAddress = txtEmailAddress.Text
                        bSaved = .SaveSupplierContact()
                    End With

            End Select

            ' Store Information in Class Object
            If bSaved = True Then
                WriteAuditLogRecord(Me.Name, "SaveSupplierUser", "OK", "Save Supplier" & objSupplier.SupplierID)
                lblSupplierIDValue.Text = objSupplier.SupplierID
                lblMessagePanel.Visible = False
                tsbSaveSupplier.Enabled = False
                mnuSaveSupplier.Enabled = False
                tsbNewSupplier.Enabled = True
                mnuNewSupplier.Enabled = True
                tsbNewOrder.Enabled = False
                mnuOrder.Enabled = False
                MsgBox("Supplier / User Data saved successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - User Information")
                LoadSupplierTree()
                pnlSupplier.Enabled = False
                pnlContact.Enabled = False
            Else
                Select Case tvwSupplier.Tag
                    Case "S"
                        WriteAuditLogRecord(Me.Name, "SaveSupplier", "FAIL", "Save Supplier")
                        txtSurname.Select()
                        MsgBox("Supplier Update Failed", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - User Information")

                    Case "C"
                        WriteAuditLogRecord(Me.Name, "SaveSupplierContact", "FAIL", "Save Supplier Contact")
                        txtSurname.Select()
                        MsgBox("Supplier Contact Update Failed", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - User Information")
                End Select
            End If
        End If

Err_SaveSupplier:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "SaveSupplier", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveSupplier" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If


    End Sub

    Private Sub tsbRemoveSupplier_Click(sender As Object, e As EventArgs) Handles tsbRemoveSupplier.Click
        '** Remove Selected Supplier

        ' Error Checking
        On Error GoTo Err_tsbRemoveSupplier_Click

        ' Check with User
        Select Case tvwSupplier.Tag
            Case "S"
                If MsgBox("Removing Supplier: " & txtCompany.Text, MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "AztecCRM - Confirmation") = MsgBoxResult.Ok Then
                    With objSupplier
                        .SupplierID = Val(lblSupplierIDValue.Text)
                        If .SupplierRemoveItem() Then
                            MsgBox("Supplier Removed Successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Confirmation")
                        End If
                    End With
                    LoadSupplierTree()
                End If

            Case "C"
                If MsgBox("Removing Supplier Contact: " & txtForename.Text & " " & txtSurname.Text, MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "AztecCRM - Confirmation") = MsgBoxResult.Ok Then
                    With objSupplier
                        .SupplierID = Val(lblSupplierIDValue.Text)
                        .ContactID = Val(lblContactIDValue.Text)
                        If .SupplierContactRemoveItem() Then
                            MsgBox("Supplier Contact Removed Successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Confirmation")
                        End If
                    End With
                    LoadSupplierTree()
                End If
        End Select
        tsbSaveSupplier.Enabled = False
        mnuSaveSupplier.Enabled = False
        tsbNewSupplier.Enabled = True
        mnuNewSupplier.Enabled = True
        tsbNewOrder.Enabled = False
        mnuOrder.Enabled = False

Err_tsbRemoveSupplier_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tsbRemoveSupplier_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tsbRemoveSupplier_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If


    End Sub


    Private Sub tvwSupplier_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvwSupplier.AfterSelect
        '** Load Data Records

        ' Error Checking
        On Error GoTo Err_tvwSupplier_AfterSelect

        ' Dimension Local Variables
        Dim sNode, sNodeType As String

        ' Get Data from Node
        sNode = tvwSupplier.SelectedNode.Tag.ToString
        sNodeType = sNode.Substring(0, 1)
        tvwSupplier.Tag = sNodeType

        ' Set Menus & Buttons
        mnuNewSupplier.Enabled = True
        pnlSupplier.Enabled = True
        pnlContact.Enabled = True


        Select Case sNodeType
            Case "S"
                ' Extract Selected Record
                objSupplier.SupplierID = sNode.Substring(2, 4)
                objSupplier.LoadSupplier() ' Initiate Class Data Load
                DisplaySupplier()
                pnlSupplier.Visible = True
                pnlContact.Visible = False
                tsbNewContact.Enabled = True
                mnuNewContact.Enabled = True

            Case "C"
                ' Extract Selected Record
                objSupplier.SupplierID = sNode.Substring(2, 4)
                objSupplier.ContactID = sNode.Substring(7, 2)
                objSupplier.LoadSupplierContact() ' Initiate Class Data Load
                DisplaySupplierContact()
                pnlSupplier.Visible = False
                pnlContact.Visible = True
                pnlContact.Left = 396
                pnlContact.Top = 104
                tsbNewContact.Enabled = False
                mnuNewContact.Enabled = False

        End Select
        LoadSupplierOrderList()
        tsbSaveSupplier.Enabled = True
        mnuSaveSupplier.Enabled = True
        tsbRemoveSupplier.Enabled = True
        mnuRemoveSupplier.Enabled = True
        tsbNewOrder.Enabled = True
        mnuOrder.Enabled = True
        pnlSupplier.Enabled = True
        WriteAuditLogRecord(Me.Name, "tvwSupplier_AfterSelect", "ACTION", "Load Supplier" & sNode)


Err_tvwSupplier_AfterSelect:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tvwSupplier_AfterSelect", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tvwSupplier_AfterSelect" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub


    Private Sub tsbAddContact_Click(sender As Object, e As EventArgs) Handles tsbNewContact.Click
        ' Clear Data
        tvwSupplier.Tag = "C"
        pnlSupplier.Visible = False
        pnlContact.Visible = True
        pnlContact.Left = 396
        pnlContact.Top = 104
        lblContactIDValue.Text = "00"
        txtTitle.Text = ""
        txtForename.Text = ""
        txtSurname.Text = ""
        txtPhoneNo.Text = ""
        txtMobileNo.Text = ""
        txtEmailAddress.Text = ""

        tsbNewSupplier.Enabled = False
        mnuNewSupplier.Enabled = False
        tsbNewContact.Enabled = False
        mnuNewContact.Enabled = False
        tsbSaveSupplier.Enabled = True
        mnuSaveSupplier.Enabled = True
        tsbRemoveSupplier.Enabled = False
        mnuRemoveSupplier.Enabled = False
        pnlContact.Enabled = True

    End Sub

    Private Sub txtTitle_TextChanged(sender As Object, e As EventArgs) Handles txtTitle.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtForename_TextChanged_1(sender As Object, e As EventArgs) Handles txtForename.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtSurname_TextChanged(sender As Object, e As EventArgs) Handles txtSurname.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtEmailAddress_TextChanged(sender As Object, e As EventArgs) Handles txtEmailAddress.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtPhoneNo_TextChanged(sender As Object, e As EventArgs) Handles txtPhoneNo.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtMobileNo_TextChanged(sender As Object, e As EventArgs) Handles txtMobileNo.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtCompany_TextChanged(sender As Object, e As EventArgs) Handles txtCompany.TextChanged
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

    Private Sub txtAddress4_TextChanged(sender As Object, e As EventArgs) Handles txtAddress4.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtPostcode_TextChanged(sender As Object, e As EventArgs) Handles txtPostcode.TextChanged
        bDataChanged = True

    End Sub

    Private Sub txtWebsite_TextChanged(sender As Object, e As EventArgs) Handles txtWebsite.TextChanged
        bDataChanged = True

    End Sub
End Class