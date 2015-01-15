Imports Microsoft.Office.Interop

Public Class frmSalesInvoice

    Dim objPayment As New clsPayment
    Dim objJob As New clsJob

    Private Sub tsbClose_Click(sender As Object, e As EventArgs) Handles tsbClose.Click
        Me.Close()

    End Sub

    Private Sub tsbEMailInvoice_Click(sender As Object, e As EventArgs) Handles tsbEMailInvoice.Click
        ' Error Checking
        On Error GoTo Err_cmdEMailInvoice_Click

        If tsbEMailInvoice.Tag = "" Then
            MsgBox("No EMail Address on File", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AztecCRM - Action Failed")
            Exit Sub
        End If

        ' Create Report & Export It
        If cmbPaymentType.SelectedIndex = 1 Then
            Dim ListReport = New FastReport.Report
            ListReport.Load(SystemPath & "Reports\ReportValuation - " & My.Settings.CompanyName & ".frx")
            ListReport.SetParameterValue("CRMConnectionString", "Data Source=" & My.Settings.SQLServer & ";AttachDbFilename=;Initial Catalog=" & My.Settings.SQLDatabase & ";Integrated Security=False;Persist Security Info=False;User ID=CRMUser;Password=S0rtmypc!")
            ListReport.SetParameterValue("InvoicePaymentID", Val(lblInvoiceIDValue.Text))
            ListReport.Prepare()
            Dim PDFExport As FastReport.Export.Pdf.PDFExport = New FastReport.Export.Pdf.PDFExport
            ListReport.Export(PDFExport, SystemPath & "Reports\EMailValuation" & lblJobIDValue.Text & ".pdf")
        Else
            ' Create Report & Export It
            Dim ListReport = New FastReport.Report
            ListReport.Load(SystemPath & "Reports\ReportInvoice - " & My.Settings.CompanyName & ".frx")
            ListReport.SetParameterValue("CRMConnectionString", "Data Source=" & My.Settings.SQLServer & ";AttachDbFilename=;Initial Catalog=" & My.Settings.SQLDatabase & ";Integrated Security=False;Persist Security Info=False;User ID=CRMUser;Password=S0rtmypc!")
            ListReport.SetParameterValue("InvoicePaymentID", Val(lblInvoiceIDValue.Text))
            ListReport.Prepare()
            Dim PDFExport As FastReport.Export.Pdf.PDFExport = New FastReport.Export.Pdf.PDFExport
            ListReport.Export(PDFExport, SystemPath & "Reports\EMailInvoice" & lblJobIDValue.Text & ".pdf")
        End If

        ' Create EMail 
        Dim objOutlook As Object
        Dim objMailMessage As Outlook.MailItem
        objOutlook = CreateObject("Outlook.Application")
        objMailMessage = objOutlook.CreateItem(0)
        With objMailMessage
            .To = tsbEMailInvoice.Tag

            If cmbPaymentType.SelectedIndex = 1 Then
                .Body = "Enclosed is a copy of the Valuation" & Chr(13) & "Thank you"
                .Subject = "Your Job Ref: " & lblJobIDValue.Text & " - " & rtbNotes.Text
                .Attachments.Add(SystemPath & "Reports\EMailValuation" & lblJobIDValue.Text & ".pdf")
            Else
                .Body = "Enclosed is a copy of your Invoice" & Chr(13) & "Thank you"
                .Subject = "Your Job Ref: " & lblJobIDValue.Text & " - " & rtbNotes.Text
                .Attachments.Add(SystemPath & "Reports\EMailInvoice" & lblJobIDValue.Text & ".pdf")
            End If

            .Display()
            .Save()
            .Close(Outlook.OlInspectorClose.olDiscard)
        End With
        objMailMessage = Nothing
        objOutlook = Nothing

        ' Log It
        Dim sMessage As String
        sMessage = Replace(">" & Format(objJob.JobID, "00000") & " - " & rtbNotes.Text & " to " & tsbEMailInvoice.Tag, "'", "`")
        WriteAuditLogRecord(Me.Name, "tsbEMailQuote_Click", "INFO", sMessage)
        MsgBox("Email has been saved as a Draft", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Action Confirmed")

Err_cmdEMailInvoice_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdEMailInvoice_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdEMailInvoice_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If


    End Sub

    Private Sub tsbPrintInvoice_Click(sender As Object, e As EventArgs) Handles tsbPrintInvoice.Click
        ' Error Checking
        On Error GoTo Err_tsbPrintInvoice_Click

        If cmbPaymentType.SelectedIndex = 1 Then
            Dim ListReport = New FastReport.Report
            ListReport.Load(SystemPath & "Reports\ReportValuation - " & My.Settings.CompanyName & ".frx")
            ListReport.SetParameterValue("CRMConnectionString", "Data Source=" & My.Settings.SQLServer & ";AttachDbFilename=;Initial Catalog=" & My.Settings.SQLDatabase & ";Integrated Security=False;Persist Security Info=False;User ID=CRMUser;Password=S0rtmypc!")
            ListReport.SetParameterValue("InvoicePaymentID", Val(lblInvoiceIDValue.Text))
            ListReport.Show()
        Else
            Dim ListReport = New FastReport.Report
            ListReport.Load(SystemPath & "Reports\ReportInvoice - " & My.Settings.CompanyName & ".frx")
            ListReport.SetParameterValue("CRMConnectionString", "Data Source=" & My.Settings.SQLServer & ";AttachDbFilename=;Initial Catalog=" & My.Settings.SQLDatabase & ";Integrated Security=False;Persist Security Info=False;User ID=CRMUser;Password=S0rtmypc!")
            ListReport.SetParameterValue("InvoicePaymentID", Val(lblInvoiceIDValue.Text))
            ListReport.Show()
        End If

Err_tsbPrintInvoice_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tsbPrintInvoice_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tsbPrintInvoice_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub frmSalesInvoice_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'AZTECCRMLoadInvoiceItems.Payment_LoadSalesInvoiceItems' table. You can move, or remove it, as needed.
        PaymentLoadSalesInvoiceItemsBindingSource.Filter = "InvoicePaymentID = 0"
        PopulateCombo(cmbPaymentType)
        If My.Settings.CompanyName = "HallidayLighting" Then
            lblMCD.Visible = True
            txtMCD.Visible = True
            lblMCDValue.Visible = True
            lblRetention.Visible = True
            txtRetention.Visible = True
            lblRetentionValue.Visible = True
            Me.Height = 640
        Else
            Me.Height = 452
        End If
        With objPayment
            .JobID = Val(lblJobIDValue.Text)
            .GetSalesInvoiceTotals()
            If mnuInvoice.Text = "&Valuation" Then
                lblTotalAmount.Text = Format(.ValuationAmount, "£#,##0.00")
                cmbPaymentType.SelectedIndex = 1
            Else
                lblTotalAmount.Text = Format(.TotalAmount, "£#,##0.00")
                cmbPaymentType.SelectedIndex = 2
            End If
            lblTotalPaid.Text = Format(-.TotalPaid, "£#,##0.00")
        End With
        ColourCodeGrid()

    End Sub

    Private Sub tsbAddInvoice_Click(sender As Object, e As EventArgs) Handles tsbAddInvoice.Click

        lblInvoiceIDValue.Text = "0000"
        dtpInvoiceDate.Value = Now
        cmbPaymentType.SelectedIndex = 0
        txtAmount.Text = ""
        txtMCD.Text = "2.5"
        txtRetention.Text = "5.0"
        rtbNotes.Text = ""

    End Sub


    Private Sub mnuInvoiceAdd_Click(sender As Object, e As EventArgs) Handles mnuInvoiceAdd.Click
        tsbAddInvoice.PerformClick()

    End Sub

    Private Sub mnuInvoiceUpdate_Click(sender As Object, e As EventArgs) Handles mnuInvoiceSave.Click
        tsbSaveInvoice.PerformClick()

    End Sub

    Private Sub mnuInvoicePrint_Click(sender As Object, e As EventArgs) Handles mnuInvoicePrint.Click
        tsbPrintInvoice.PerformClick()

    End Sub

    Private Sub mnuInvoiceEMail_Click(sender As Object, e As EventArgs) Handles mnuInvoiceEMail.Click
        tsbEMailInvoice.PerformClick()

    End Sub

    Private Sub dgvSalesInvoices_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSalesInvoices.CellContentClick

    End Sub

    Private Sub dgvSalesInvoices_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvSalesInvoices.KeyDown

        ' Error Checking
        On Error GoTo Err_dgvSalesInvoices_KeyDown

        ' Check for Delete Key
        If e.KeyCode = Keys.Delete Then
            '            Dim CurrentRow As DataRowView = TryCast(PaymentLoadSalesInvoicesBindingSource.Current, DataRowView)
            '            If MsgBox("Do you want to delete Invoice No " & CurrentRow("InvoicePaymentID") & " ?", vbYesNo + vbQuestion, "Remove Invoice Record") = vbYes Then
            ' With objPayment
            ' .InvoicePaymentID = CurrentRow("InvoicePaymentID")
            ' .RowNo = CurrentRow("RowNo")
            ' .RemoveSalesInvoiceItem()
            ' WriteAuditLogRecord(Me.Name, "dgvSalesInvoices_KeyDown", "OK", "Remove Invoice Record: " & CurrentRow("InvoicePaymentID"))
            ' End With
            'End If
        End If
        '        Me.Payment_LoadSalesInvoicesTableAdapter.Fill(Me.AZTECCRMLoadSalesInvoices.Payment_LoadSalesInvoices)

Err_dgvSalesInvoices_KeyDown:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "dgvSalesInvoices_KeyDown", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "dgvSalesInvoices_KeyDown" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub dgvSalesInvoices_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvSalesInvoices.RowHeaderMouseClick

        ' Error Checking
        On Error GoTo Err_dgvSalesInvoices_RowHeaderMouseClick

        ' Update Button / Menu Status / Redisplay Data
        '        Dim CurrentRow As DataRowView = TryCast(PaymentLoadSalesInvoicesBindingSource.Current, DataRowView)
        '        tsbSaveInvoice.Enabled = CurrentRow("PaymentTypeID") <> 0
        '        mnuInvoiceSave.Enabled = CurrentRow("PaymentTypeID") <> 0
        '        tsbPrintInvoice.Enabled = CurrentRow("PaymentTypeID") = 1 Or CurrentRow("PaymentTypeID") = 2
        '        mnuInvoicePrint.Enabled = CurrentRow("PaymentTypeID") = 1 Or CurrentRow("PaymentTypeID") = 2
        '        tsbEMailInvoice.Enabled = CurrentRow("PaymentTypeID") = 1 Or CurrentRow("PaymentTypeID") = 2
        '        mnuInvoiceEMail.Enabled = CurrentRow("PaymentTypeID") = 1 Or CurrentRow("PaymentTypeID") = 2

        ' Load Invoice Details
        '        lblInvoiceIDValue.Text = Format(CurrentRow("InvoicePaymentID"), "0000")
        '        dtpInvoiceDate.Value = CurrentRow("InvoicePaymentDate")
        '        txtAmount.Text = CurrentRow("Amount")
        '        txtMCD.Text = CurrentRow("MCDPer")
        '        txtRetention.Text = CurrentRow("RetentionPer")
        For iRCnt = 1 To cmbPaymentType.Items.Count - 1
            '           If cmbPaymentType.Items.Item(iRCnt).ToString = CurrentRow("PaymentTypeDescription") Then
            cmbPaymentType.SelectedIndex = iRCnt
            '            End If
        Next
        '        rtbNotes.Text = CurrentRow("Notes")

        ' Load Invoice Items
        PaymentLoadSalesInvoiceItemsBindingSource.Filter = "InvoicePaymentID = " & Val(lblInvoiceIDValue.Text)
        With objPayment
            .InvoicePaymentID = Val(lblInvoiceIDValue.Text)
            .GetSalesInvoiceItemTotals()
            lblItemsAmountValue.Text = Format(.TotalPrice, "£#,##0.00")
            If lblInvoiceAmountValue.Text = lblItemsAmountValue.Text Then
                lblInvoiceAmountValue.ForeColor = Color.Green
                lblItemsAmountValue.ForeColor = Color.Green
            Else
                lblInvoiceAmountValue.ForeColor = Color.Black
                lblItemsAmountValue.ForeColor = Color.Black
            End If
        End With
        '
Err_dgvSalesInvoices_RowHeaderMouseClick:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "dgvSalesInvoices_RowHeaderMouseClick", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "dgvSalesInvoices_RowHeaderMouseClick" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub dgvSalesInvoices_RowHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvSalesInvoices.RowHeaderMouseDoubleClick


    End Sub

    Private Sub mnuFileClose_Click(sender As Object, e As EventArgs) Handles mnuFileClose.Click
        Me.Close()
    End Sub

    Sub CalculateTotals()
        ' ** Calculate Invoice Totals
        lblMCDValue.Text = Format(ValCur(txtAmount.Text) * txtMCD.Text / 100, "£#,##0.00")
        lblRetentionValue.Text = Format((ValCur(txtAmount.Text) - ValCur(txtAmount.Text) * ValCur(txtMCD.Text) / 100) * ValCur(txtRetention.Text) / 100, "£#,##0.00")
        lblInvoiceAmountValue.Text = Format(ValCur(txtAmount.Text) - ValCur(lblMCDValue.Text) - ValCur(lblRetentionValue.Text), "£#,##0.00")

        If ValCur(lblInvoiceAmountValue.Text) <> 0 Then
            mnuInvoiceSave.Enabled = True
            tsbSaveInvoice.Enabled = True
        End If

        ' Calculate Sales Item Total
        objPayment.GetSalesInvoiceItemTotals()
        lblItemsAmountValue.Text = Format(objPayment.TotalPrice, "£#,##0.00")
        If lblInvoiceAmountValue.Text = lblItemsAmountValue.Text Then
            lblInvoiceAmountValue.ForeColor = Color.Green
            lblItemsAmountValue.ForeColor = Color.Green
        Else
            lblInvoiceAmountValue.ForeColor = Color.Black
            lblItemsAmountValue.ForeColor = Color.Black
        End If

    End Sub

    Private Sub txtAmount_TextChanged(sender As Object, e As EventArgs) Handles txtAmount.TextChanged
        CalculateTotals()

    End Sub

    Private Sub txtMCD_TextChanged(sender As Object, e As EventArgs) Handles txtMCD.TextChanged
        CalculateTotals()

    End Sub

    Private Sub txtRetention_TextChanged(sender As Object, e As EventArgs) Handles txtRetention.TextChanged
        CalculateTotals()

    End Sub

    Private Sub tsbSaveInvoice_Click(sender As Object, e As EventArgs) Handles tsbSaveInvoice.Click
        ' Error Checking
        On Error GoTo Err_tsbAddInvoice_Click

        ' Check for Valid Entry
        If cmbPaymentType.SelectedIndex = 0 Or txtAmount.Text = "" Then Exit Sub

        ' Add New Payment
        With objPayment
            .InvoicePaymentID = Val(lblInvoiceIDValue.Text)
            .JobID = Val(lblJobIDValue.Text)
            .InvoicePaymentDate = dtpInvoiceDate.Value
            .Amount = txtAmount.Text
            .MCDPer = txtMCD.Text
            .RetentionPer = txtRetention.Text
            .PaymentTypeID = cmbPaymentType.SelectedValue
            .Notes = rtbNotes.Text
            If .SaveSalesInvoice() Then
                dtpInvoiceDate.Value = Now
                cmbPaymentType.SelectedIndex = 0
                txtAmount.Text = ""
                txtMCD.Text = "2.5"
                txtRetention.Text = "5.0"
                rtbNotes.Text = ""
            End If
        End With

        ' Update Form
        With objPayment
            .GetSalesInvoiceTotals()
            If mnuInvoice.Text = "&Valuation" Then
                lblTotalAmount.Text = Format(.ValuationAmount, "£#,##0.00")
            Else
                lblTotalAmount.Text = Format(.TotalAmount, "£#,##0.00")
            End If
            lblTotalPaid.Text = Format(-.TotalPaid, "£#,##0.00")
        End With
        tsbAddInvoice.Enabled = True
        ColourCodeGrid()
        grpInvoiceItems.Enabled = Val(lblInvoiceIDValue.Text) <> 0

Err_tsbAddInvoice_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tsbAddInvoice_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tsbAddInvoice_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "DreamEvent - Error Reporting")
        End If

    End Sub

    Sub ColourCodeGrid()
        ' Colour Code Titles
        For Each row As DataGridViewRow In dgvSalesInvoices.Rows
            If row.Cells.Item(10).Value = 1 Then
                row.DefaultCellStyle.BackColor = Color.LemonChiffon
            ElseIf row.Cells.Item(10).Value = 2 Then
                row.DefaultCellStyle.BackColor = Color.Gold
            ElseIf row.Cells.Item(10).Value = 3 Then
                row.DefaultCellStyle.BackColor = Color.LightGreen
            End If
        Next

    End Sub

    Private Sub dgvInvoiceItems_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvInvoiceItems.KeyDown

        ' Error Checking
        On Error GoTo Err_dgvInvoiceItems_KeyDown

        ' Check for Delete Key / Delete Record
        If e.KeyCode = Keys.Delete Then
            Dim CurrentRow As DataRowView = TryCast(PaymentLoadSalesInvoiceItemsBindingSource.Current, DataRowView)
            If MsgBox("Are You Sure", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "AztecCRM - Action Confirmation") = vbOK Then
                With objPayment
                    .InvoicePaymentID = CurrentRow("InvoicePaymentID")
                    .RowNo = Val(CurrentRow("RowNo"))
                    .RemoveSalesInvoiceItem()
                    CalculateTotals()
                    lblItemDescription.Text = ""
                    nudQuantity.Value = 1
                    txtPrice.Text = ""
                    PaymentLoadSalesInvoiceItemsBindingSource.Filter = "InvoicePaymentID = " & Val(lblInvoiceIDValue.Text)
                    WriteAuditLogRecord(Me.Name, "dgvInvoiceItems_KeyDown", "OK", "Remove Invoice Item: " & lblInvoiceIDValue.Text & " - " & .RowNo)
                    .GetSalesInvoiceItemTotals()
                    If mnuInvoice.Text = "&Valuation" Then
                        lblTotalAmount.Text = Format(.ValuationAmount, "£#,##0.00")
                    Else
                        lblTotalAmount.Text = Format(.TotalAmount, "£#,##0.00")
                    End If
                    If lblInvoiceAmountValue.Text = lblItemsAmountValue.Text Then
                        lblInvoiceAmountValue.ForeColor = Color.Green
                        lblItemsAmountValue.ForeColor = Color.Green
                    Else
                        lblInvoiceAmountValue.ForeColor = Color.Black
                        lblItemsAmountValue.ForeColor = Color.Black
                    End If
                End With
            End If

Err_dgvInvoiceItems_KeyDown:
            If Err.Number <> 0 Then
                sErrDescription = Err.Description
                WriteAuditLogRecord(Me.Name, "dgvInvoiceItems_KeyDown", "Error", sErrDescription)
                MsgBox("System Error occurred" & Chr(13) & "dgvInvoiceItems_KeyDown" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            End If


        End If

    End Sub

    Private Sub cmdAddItem_Click(sender As Object, e As EventArgs) Handles cmdAddItem.Click

        ' Error Checking
        On Error GoTo Err_cmdAddItem_Click

        '        Dim CurrentRow As DataRowView = TryCast(PaymentLoadSalesInvoicesBindingSource.Current, DataRowView)
        '        With objPayment
        ' .JobID = Val(lblJobIDValue.Text)
        ' .InvoicePaymentID = CurrentRow("InvoicePaymentID")
        ' .RowNo = 0
        ' .ItemDescription = txtItemDescription.Text
        ' .Quantity = nudQuantity.Value
        ' .Price = Val(txtPrice.Text)
        ' .SaveSalesInvoiceItem()
        CalculateTotals()
        txtItemDescription.Text = ""
        nudQuantity.Value = 1
        txtPrice.Text = ""
        PaymentLoadSalesInvoiceItemsBindingSource.Filter = "InvoicePaymentID = " & Val(lblInvoiceIDValue.Text)
        WriteAuditLogRecord(Me.Name, "cmdAddItem_Click", "OK", "Add Invoice Item: " & lblInvoiceIDValue.Text)
        ' End With

Err_cmdAddItem_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdAddItem_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdAddItem_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub


    Private Sub cmbPaymentType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPaymentType.SelectedIndexChanged
        Select Case cmbPaymentType.SelectedIndex
            Case 1
                tsbAddInvoice.Text = "Add Valuation"
                tsbPrintInvoice.Text = "Print Valuation"
                tsbEMailInvoice.Text = "EMail Valuation"
                txtMCD.Enabled = True
                txtRetention.Enabled = True
                mnuInvoice.Text = "&Valuation"

            Case 2
                tsbAddInvoice.Text = "Add Invoice"
                tsbPrintInvoice.Text = "Print Invoice"
                tsbEMailInvoice.Text = "EMail Invoice"
                txtMCD.Text = "0.0"
                txtMCD.Enabled = False
                txtRetention.Text = "0.0"
                txtRetention.Enabled = False
                mnuInvoice.Text = "&Invoice"

        End Select
    End Sub

    Private Sub dgvInvoiceItems_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInvoiceItems.CellContentClick

    End Sub
End Class