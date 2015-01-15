

Public Class frmFinance

    ' Dimension Form Variables
    Dim objPayment As New clsPayment

    Private Sub frmFinance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ColourCodeGrid()

    End Sub

    Private Sub tsbClose_Click(sender As Object, e As EventArgs) Handles tsbClose.Click
        Me.Close()

    End Sub

    Private Sub mnuFileClose_Click(sender As Object, e As EventArgs) Handles mnuFileClose.Click
        Me.Close()
    End Sub

    Private Sub dgvPurchaseInvoice_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvPurchaseInvoice.RowHeaderMouseClick
        ' Error Checking
        On Error GoTo Err_dgvPurchaseInvoice_RowHeaderMouseClick

        ' Get Invoice ID
        Dim CurrentRow As DataRowView = TryCast(PaymentLoadPurchaseInvoicesBindingSource.Current, DataRowView)
        lblPurchaseInvoiceIDValue.Text = Format(CurrentRow("InvoiceID"), "0000")
        tsbMarkAsApproved.Enabled = (My.Settings.PaymentApproved = "On")
        tsbMarkAsPaid.Enabled = (My.Settings.PaymentPaid = "On")
        tsbClearAllFlags.Enabled = True

Err_dgvPurchaseInvoice_RowHeaderMouseClick:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "dgvPurchaseInvoice_RowHeaderMouseClick", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "dgvPurchaseInvoice_RowHeaderMouseClick" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub


    Private Sub dgvSalesInvoice_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvSalesInvoice.RowHeaderMouseClick
        ' Error Checking
        On Error GoTo Err_dgvSalesInvoice_RowHeaderMouseClick

        ' Get Invoice ID
        Dim CurrentRow As DataRowView = TryCast(PaymentLoadSalesInvoicesBindingSource.Current, DataRowView)
        lblSalesInvoiceIDValue.Text = Format(CurrentRow("InvoicePaymentID"), "0000")
        lblJobIDValue.Text = Format(CurrentRow("JobID"), "0000")
        With objPayment
            .JobID = Val(lblJobIDValue.Text)
            .GetSalesInvoiceTotals()
            tsbSavePayment.Enabled = (.TotalAmount + .TotalPaid > 0)
            lblJobTotal.Text = Format(.TotalAmount + .TotalPaid, "£#,##0.00")
        End With
        tsbClearAllFlags.Enabled = True

Err_dgvSalesInvoice_RowHeaderMouseClick:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "dgvSalesInvoice_RowHeaderMouseClick", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "dgvSalesInvoice_RowHeaderMouseClick" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub tsbMarkAsPaid_Click(sender As Object, e As EventArgs) Handles tsbMarkAsPaid.Click
        '** Mark as Paid

        ' Error Checking
        On Error GoTo Err_tsbMarkAsPaid_Click

        Select Case tabPayments.Tag
            Case "P"
                With objPayment
                    .InvoiceID = Val(txtPurchaseInvoiceID.Text)
                    .PaidFlag = True
                    .PaidDate = dtpPurchaseInvoiceDatePaid.Value
                    .Notes = txtPurchaseInvoiceNotes.Text & "(Paid By " & Environment.UserName & ")"
                    .UpdatePurchaseInvoicePaidStatus()
                    lblPurchaseInvoiceIDValue.Text = "0000"
                    dtpPurchaseInvoiceDatePaid.Value = Now
                    txtPurchaseInvoiceNotes.Text = ""
                    tsbMarkAsApproved.Enabled = False
                    tsbMarkAsPaid.Enabled = False
                    tsbClearAllFlags.Enabled = False
                    tsbSavePayment.Enabled = False
                    '                    Me.Payment_LoadPurchaseInvoicesTableAdapter.Fill(Me.AZTECCRMPurchaseInvoices.Payment_LoadPurchaseInvoices)
                End With
                WriteAuditLogRecord(Me.Name, "tsbMarkAsPaid_Click", "Info", "Invoice No " & txtPurchaseInvoiceID.Text & " Paid")

            Case "S"
                With objPayment
                    .InvoiceID = Val(txtPurchaseInvoiceID.Text)
                    .PaidFlag = True
                    .PaidDate = dtpPurchaseInvoiceDatePaid.Value
                    .Notes = txtPurchaseInvoiceNotes.Text & "(Paid By " & Environment.UserName & ")"
                    .UpdatePurchaseInvoicePaidStatus()
                    lblPurchaseInvoiceIDValue.Text = "0000"
                    dtpPurchaseInvoiceDatePaid.Value = Now
                    txtPurchaseInvoiceNotes.Text = ""
                    tsbMarkAsApproved.Enabled = False
                    tsbMarkAsPaid.Enabled = False
                    tsbClearAllFlags.Enabled = False
                    tsbSavePayment.Enabled = False
                    '                    Me.Payment_LoadPurchaseInvoicesTableAdapter.Fill(Me.AZTECCRMPurchaseInvoices.Payment_LoadPurchaseInvoices)
                End With
                WriteAuditLogRecord(Me.Name, "tsbMarkAsPaid_Click", "Info", "Invoice No " & txtPurchaseInvoiceID.Text & " Paid")

        End Select

Err_tsbMarkAsPaid_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tsbMarkAsPaid_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tsbMarkAsPaid_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If
    End Sub

    Private Sub tsbClearAllFlags_Click(sender As Object, e As EventArgs) Handles tsbClearAllFlags.Click
        '** Mark as Paid

        ' Error Checking
        On Error GoTo Err_tsbClearAllFlags_Click

        Select Case tabPayments.Tag
            Case "P"
                With objPayment
                    .InvoiceID = Val(lblPurchaseInvoiceIDValue.Text)
                    .ClearPurchaseInvoiceStatusFlags()
                    lblPurchaseInvoiceIDValue.Text = "0000"
                    dtpPurchaseInvoiceDatePaid.Value = Now
                    txtPurchaseInvoiceNotes.Text = ""
                    '                   Me.Payment_LoadPurchaseInvoicesTableAdapter.Fill(Me.AZTECCRMPurchaseInvoices.Payment_LoadPurchaseInvoices)
                End With
                WriteAuditLogRecord(Me.Name, "tsbMarkAsPaid_Click", "Info", "Invoice No " & txtPurchaseInvoiceID.Text & " Cleared")

            Case "S"
                With objPayment
                    .InvoicePaymentID = Val(lblSalesInvoiceIDValue.Text)
                    .ClearSalesInvoiceStatusFlags()
                    lblSalesInvoiceIDValue.Text = "0000"
                    dtpSalesInvoiceDatePaid.Value = Now
                    txtSalesInvoiceNotes.Text = ""
                    '                    Me.Payment_LoadSalesInvoicesTableAdapter.Fill(Me.AZTECCRMSalesInvoices.Payment_LoadSalesInvoices)
                End With
                WriteAuditLogRecord(Me.Name, "tsbMarkAsPaid_Click", "Info", "Invoice No " & txtSalesInvoiceID.Text & " Cleared")
        End Select

        ' Clear Up
        tsbMarkAsApproved.Enabled = False
        tsbMarkAsPaid.Enabled = False
        tsbClearAllFlags.Enabled = False
        tsbSavePayment.Enabled = False
        ColourCodeGrid()

Err_tsbClearAllFlags_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tsbClearAllFlags_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tsbClearAllFlags_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub tabPayments_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabPayments.SelectedIndexChanged
        ' Check Tab
        Select Case tabPayments.SelectedTab.Name
            Case "tabPaymentsPurchases"
                tabPayments.Tag = "P"
                tsbSavePayment.Enabled = False
                tsbMarkAsApproved.Enabled = (My.Settings.PaymentApproved = "On")
                tsbMarkAsPaid.Enabled = (My.Settings.PaymentPaid = "On")
                tsbClearAllFlags.Enabled = False

            Case "tabPaymentsSales"
                tabPayments.Tag = "S"
                tsbSavePayment.Enabled = False
                tsbMarkAsApproved.Enabled = False
                tsbMarkAsPaid.Enabled = False
                tsbClearAllFlags.Enabled = False

        End Select
        ColourCodeGrid()

    End Sub

    Private Sub tsbMarkAsApproved_Click(sender As Object, e As EventArgs) Handles tsbMarkAsApproved.Click
        '** Mark as Paid

        ' Error Checking
        On Error GoTo Err_tsbMarkAsApproved_Click

        With objPayment
            .InvoiceID = Val(lblPurchaseInvoiceIDValue.Text)
            .ApprovedFlag = True
            .ApprovedDate = dtpPurchaseInvoiceDatePaid.Value
            .Notes = txtPurchaseInvoiceNotes.Text & "(Approved By " & Environment.UserName & ")"
            .UpdatePurchaseInvoiceApprovedStatus()
            lblPurchaseInvoiceIDValue.Text = "0000"
            dtpPurchaseInvoiceDatePaid.Value = Now
            txtPurchaseInvoiceNotes.Text = ""
            tsbMarkAsApproved.Enabled = False
            tsbMarkAsPaid.Enabled = False
            tsbClearAllFlags.Enabled = False
            tsbSavePayment.Enabled = False
            '           Me.Payment_LoadPurchaseInvoicesTableAdapter.Fill(Me.AZTECCRMPurchaseInvoices.Payment_LoadPurchaseInvoices)
            ColourCodeGrid()
        End With
        WriteAuditLogRecord(Me.Name, "tsbMarkAsApproved_Click", "Info", "Invoice No " & txtPurchaseInvoiceID.Text & " Approved")

Err_tsbMarkAsApproved_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tsbMarkAsApproved_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tsbMarkAsApproved_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub tabPaymentsPurchases_Click(sender As Object, e As EventArgs) Handles tabPaymentsPurchases.Click

    End Sub

    Private Sub dgvPurchaseInvoice_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPurchaseInvoice.CellContentClick

    End Sub

    Private Sub tsbSavePayment_Click(sender As Object, e As EventArgs) Handles tsbSavePayment.Click
        '** Mark as Paid

        ' Error Checking
        On Error GoTo Err_tsbSavePayment_Click

        ' Dimension Local Variables
        Dim iInvoicePaymentID As Integer

        ' Check for Invalid Value
        If ValCur(txtAmount.Text) = 0 Then Exit Sub

        With objPayment
            .InvoicePaymentID = 0
            .JobID = Val(lblJobIDValue.Text)
            .InvoicePaymentDate = dtpSalesInvoiceDatePaid.Value
            .Amount = -ValCur(txtAmount.Text)
            .PaymentTypeID = 3
            .Notes = txtSalesInvoiceNotes.Text
            .SaveSalesInvoice()
            iInvoicePaymentID = .InvoicePaymentID
            lblSalesInvoiceIDValue.Text = "0000"
            dtpSalesInvoiceDatePaid.Value = Now
            txtAmount.Text = ""
            txtSalesInvoiceNotes.Text = ""
            tsbClearAllFlags.Enabled = False
            '            Me.Payment_LoadSalesInvoicesTableAdapter.Fill(Me.AZTECCRMSalesInvoices.Payment_LoadSalesInvoices)
            ColourCodeGrid()
        End With
        WriteAuditLogRecord(Me.Name, "tsbSavePayment_Click", "Info", "Invoice No " & Format(iInvoicePaymentID, "0000") & " Paid")

        With objPayment
            .JobID = Val(txtJobID.Text)
            .GetSalesInvoiceTotals()
            lblJobIDValue.Text = Val(txtJobID.Text)
            tsbSavePayment.Enabled = (.TotalAmount + .TotalPaid > 0)
            lblJobTotal.Text = Format(.TotalAmount + .TotalPaid, "£#,##0.00")
        End With

Err_tsbSavePayment_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "tsbSavePayment_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "tsbSavePayment_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub txtPurchaseInvoiceID_TextChanged(sender As Object, e As EventArgs) Handles txtPurchaseInvoiceID.TextChanged
        ApplyPurchaseInvoiceFilter()

    End Sub

    Private Sub txtSalesInvoiceID_TextChanged(sender As Object, e As EventArgs) Handles txtSalesInvoiceID.TextChanged
        ApplySalesInvoiceFilter()

    End Sub

    Private Sub txtOrderID_TextChanged(sender As Object, e As EventArgs) Handles txtOrderID.TextChanged
        ApplyPurchaseInvoiceFilter()

    End Sub

    Private Sub txtJobID_TextChanged(sender As Object, e As EventArgs) Handles txtJobID.TextChanged
        ApplySalesInvoiceFilter()

    End Sub

    Private Sub tabPaymentsSales_Click(sender As Object, e As EventArgs) Handles tabPaymentsSales.Click

    End Sub

    Sub ColourCodeGrid()
        ' Colour Code Titles
        For Each row As DataGridViewRow In dgvSalesInvoice.Rows
            If row.Cells.Item("PaymentTypeID").Value = 1 Then
                row.DefaultCellStyle.BackColor = Color.LemonChiffon
            ElseIf row.Cells.Item("PaymentTypeID").Value = 2 Then
                row.DefaultCellStyle.BackColor = Color.Gold
            ElseIf row.Cells.Item("PaymentTypeID").Value = 3 Then
                row.DefaultCellStyle.BackColor = Color.LightGreen
            End If
        Next

    End Sub

    Sub ApplyPurchaseInvoiceFilter()

        ' Invoice Filter
        If Val(txtPurchaseInvoiceID.Text) = 0 Then
            If chkShowCancelledPurchaseInvoices.Checked = True Then
                If chkShowApprovedPurchaseInvoices.Checked = True Then
                    PaymentLoadPurchaseInvoicesBindingSource.Filter = "ApprovedFlag = 1"
                Else
                    PaymentLoadPurchaseInvoicesBindingSource.Filter = "ApprovedFlag = 0"
                End If
            Else
                If chkShowApprovedPurchaseInvoices.Checked = True Then
                    PaymentLoadPurchaseInvoicesBindingSource.Filter = "CancelFlag = 0 AND ApprovedFlag = 1"
                Else
                    PaymentLoadPurchaseInvoicesBindingSource.Filter = "CancelFlag = 0 AND ApprovedFlag = 0"
                End If
            End If
        Else
            If chkShowCancelledPurchaseInvoices.Checked = True Then
                PaymentLoadPurchaseInvoicesBindingSource.Filter = "InvoiceID = " & txtPurchaseInvoiceID.Text
            Else
                PaymentLoadPurchaseInvoicesBindingSource.Filter = "InvoiceID = " & txtPurchaseInvoiceID.Text
            End If
            Exit Sub
        End If

        ' Apply Order ID Filter
        If Val(txtOrderID.Text) = 0 Then
            If chkShowCancelledSalesInvoices.Checked = True Then
                If chkShowApprovedPurchaseInvoices.Checked = True Then
                    PaymentLoadPurchaseInvoicesBindingSource.Filter = "ApprovedFlag = 1"
                Else
                    PaymentLoadPurchaseInvoicesBindingSource.Filter = "ApprovedFlag = 0"
                End If
            Else
                If chkShowApprovedPurchaseInvoices.Checked = True Then
                    PaymentLoadPurchaseInvoicesBindingSource.Filter = "CancelFlag = 0 AND ApprovedFlag = 1"
                Else
                    PaymentLoadPurchaseInvoicesBindingSource.Filter = "CancelFlag = 0 AND ApprovedFlag = 0"
                End If
            End If
        Else
            If chkShowCancelledPurchaseInvoices.Checked = True Then
                PaymentLoadPurchaseInvoicesBindingSource.Filter = "OrderID = " & txtOrderID.Text
            Else
                PaymentLoadPurchaseInvoicesBindingSource.Filter = "OrderID = " & txtOrderID.Text
            End If
        End If

    End Sub

    Sub ApplySalesInvoiceFilter()

        ' Invoice Filter
        If Val(txtPurchaseInvoiceID.Text) = 0 Then
            If chkShowCancelledSalesInvoices.Checked = True Then
                PaymentLoadSalesInvoicesBindingSource.Filter = ""
            Else
                PaymentLoadSalesInvoicesBindingSource.Filter = "CancelFlag = 0"
            End If
        Else
            If chkShowCancelledPurchaseInvoices.Checked = True Then
                PaymentLoadSalesInvoicesBindingSource.Filter = "InvoicePaymentID = " & txtSalesInvoiceID.Text
            Else
                PaymentLoadSalesInvoicesBindingSource.Filter = "InvoicePaymentID = " & txtSalesInvoiceID.Text & " AND CancelFlag = 0"
            End If
            ColourCodeGrid()
            Exit Sub
        End If

        ' Job Filter
        If Val(txtJobID.Text) = 0 Then
            If chkShowCancelledSalesInvoices.Checked = True Then
                PaymentLoadSalesInvoicesBindingSource.Filter = ""
            Else
                PaymentLoadSalesInvoicesBindingSource.Filter = "CancelFlag = 0"
            End If
            lblJobTotal.Text = "£0.00"
        Else
            If chkShowCancelledPurchaseInvoices.Checked = True Then
                PaymentLoadSalesInvoicesBindingSource.Filter = "JobID = " & txtJobID.Text
            Else
                PaymentLoadSalesInvoicesBindingSource.Filter = "JobID = " & txtJobID.Text & " AND CancelFlag = 0"
            End If
            With objPayment
                .JobID = Val(txtJobID.Text)
                .GetSalesInvoiceTotals()
                lblJobIDValue.Text = Format(Val(txtJobID.Text), "0000")
                tsbSavePayment.Enabled = (.TotalAmount + .TotalPaid > 0)
                lblJobTotal.Text = Format(.TotalAmount + .TotalPaid, "£#,##0.00")
            End With
        End If
        ColourCodeGrid()

    End Sub

    Private Sub chkShowCancelledPurchaseInvoices_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowCancelledPurchaseInvoices.CheckedChanged
        ApplyPurchaseInvoiceFilter()

    End Sub

    Private Sub chkShowCancelledSalesInvoices_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowCancelledSalesInvoices.CheckedChanged
        ApplySalesInvoiceFilter()

    End Sub

    Private Sub dgvSalesInvoice_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSalesInvoice.CellContentClick

    End Sub

    Private Sub dgvPurchaseInvoice_RowHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvPurchaseInvoice.RowHeaderMouseDoubleClick

        ' Error Checking
        On Error GoTo Err_dgvPurchaseInvoice_RowHeaderMouseDoubleClick

        ' Load Existing Order Sheet
        Dim CurrentRow As DataRowView = TryCast(PaymentLoadPurchaseInvoicesBindingSource.Current, DataRowView)

Err_dgvPurchaseInvoice_RowHeaderMouseDoubleClick:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "dgvPurchaseInvoice_RowHeaderMouseDoubleClick", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "dgvPurchaseInvoice_RowHeaderMouseDoubleClick" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub chkShowApprovedPurchaseInvoices_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowApprovedPurchaseInvoices.CheckedChanged
        ApplyPurchaseInvoiceFilter()

    End Sub
End Class