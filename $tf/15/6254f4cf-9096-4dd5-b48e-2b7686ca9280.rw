Imports DevDefined.OAuth.Consumer
Imports System.Security.Cryptography.X509Certificates
Imports XeroApi.OAuth
Imports XeroApi
Imports XeroApi.Model
Imports Microsoft.Office.Interop

Public Class frmXeroCustomerCheck

    ' Dimension Form Variables
    Dim objLandlord As New clsLandlord
    Dim objPayment As New clsPayment
    Const UserAgent As String = "Xero.API.ScreenCast v1.0 (Private App Testing)"
    Private privateCertificate As X509Certificate2 = New X509Certificate2(SystemPath & "public_privatekey.pfx", "S0rtmypc")
    Private uConsumerSession As IOAuthSession = New XeroApiPrivateSession(UserAgent, My.Settings.ConsumerKey, privateCertificate)
    Private uRepository As New Repository(uConsumerSession)

    Private Sub frmXeroCustomerCheck_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdCreateInvoice_Click(sender As Object, e As EventArgs) Handles cmdCreateInvoice.Click

        ' Error Checking
        On Error GoTo Err_cmdCreateInvoice_Click

        ' Dimension Local Variables
        Dim uInvoice As New Invoice
        Dim uContact As New Contact
        Dim sReturnCode As String
        Dim uRecSnap As ADODB.Recordset

        ' Create Invoice
        uInvoice.Type = "ACCREC"
        uContact.Name = lblContactNameValue.Text
        uContact.IsCustomer = True
        uInvoice.Contact = uContact
        uInvoice.Date = Now
        uInvoice.DueDate = DateAdd(DateInterval.Day, 14, Now)
        uInvoice.Status = "AUTHORISED"
        uInvoice.LineAmountTypes = LineAmountType.Inclusive
        uInvoice.InvoiceNumber = "CRM-" & Format(Val(lblJobIDValue.Text), "00000")
        uInvoice.Reference = "CRM Job"

        uInvoice.LineItems = New LineItems
        With objPayment
            .JobID = Val(lblJobIDValue.Text)
            .GetXeroJobProducts()
            uRecSnap = .JobProduct
            Do Until uRecSnap.EOF
                Dim uLineItem = New LineItem
                uLineItem.ItemCode = uRecSnap("ItemCode").Value
                uLineItem.Description = uRecSnap("FullDescription").Value
                uLineItem.Quantity = Val(uRecSnap("Quantity").Value)
                uLineItem.UnitAmount = Val(uRecSnap("Price").Value)
                uLineItem.AccountCode = uRecSnap("CategoryID").Value
                uInvoice.LineItems.Add(uLineItem)
                uRecSnap.MoveNext()
            Loop
        End With

        sReturnCode = uRepository.Create(Of Invoice)(uInvoice).ToString
        '        For Each uError As ValidationError In uInvoice.ValidationErrors
        '        Debug.Print(uError.Message)
        '        Next


        ' Store Details in Table
        With objPayment
            .InvoicePaymentID = 0
            .JobID = Val(lblJobIDValue.Text)
            .InvoicePaymentDate = Now
            .PaymentTypeID = 2
            .Amount = ValCur(lblAmountDueValue.Text)
            .MCDPer = 0
            .RetentionPer = 0
            .Notes = ""
            If .SaveSalesInvoice() = True Then
                .PaymentTypeID = 2
                .XeroID = sReturnCode.Substring(20, 36)
                If .UpdateSalesInvoiceXeroID() = True Then
                    MsgBox("Invoice Created Sucessfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Information")
                    cmdCreateInvoice.Enabled = False
                    cmdEMailInvoice.Enabled = True
                    lblAmountInvoicedValue.Text = lblAmountDueValue.Text
                Else
                    MsgBox("Invoice Has Failed", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Information")
                End If
            Else
                MsgBox("Invoice Has Failed", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Information")
            End If
        End With

Err_cmdCreateInvoice_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdCreateInvoice_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdCreateInvoice_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "BMSSortmyPC - Error Reporting")
        End If

    End Sub

    Private Sub cmdGetCustomers_Click(sender As Object, e As EventArgs) Handles cmdGetCustomers.Click

        ' Error Checking
        Try
            ' Populate List Box
            Dim uContacts = From Contacts In uRepository.Contacts Where (Contacts.LastName = lblSurnameValue.Text Or Contacts.Name = lblCompanyValue.Text)
            lstXeroContacts.Items.Clear()
            For Each uContact As Contact In uContacts
                lstXeroContacts.Items.Add(Format(Val(uContact.ContactNumber), "0000") & " - " & uContact.Name)
            Next

            ' Check for NO Match
            cmdCreateCustomer.Enabled = True
        Catch uError As Exception
            sErrDescription = uError.Message
            WriteAuditLogRecord(Me.Name, "cmdGetCustomers_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdGetCustomers_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "BMSSortmyPC - Error Reporting")
        End Try

    End Sub

    Public Sub DisplayClient()
        '** Display Client Data

        ' Error Checking
        Try
            ' Get Personal ClientData
            With objLandlord
                .LandlordID = Val(lblLandlordIDValue.Text)
                objLandlord.LoadLandlord()
                lblTitleValue.Text = .Title1
                lblForenameValue.Text = .Forename1
                lblSurnameValue.Text = .Surname1
                lblAddress1Value.Text = .Address1
                lblAddress2Value.Text = .Address2
                lblAddress3Value.Text = .Address3
                lblAddress4Value.Text = .Town
                lblPostcodeValue.Text = .PostCode
                lblPhoneNoValue.Text = .PhoneNo
                lblMobileNoValue.Text = .MobileNo
                lblEMailAddressValue.Text = .EMailAddress

            End With

            ' Check for ContactNumber (Which is ClientID)
            Dim uContacts = From Contacts In uRepository.Contacts Where (Contacts.ContactNumber = lblLandlordIDValue.Text)
            lstXeroContacts.Items.Clear()
            For Each uContact As Contact In uContacts
                lblLandlordIDValue.ForeColor = Color.Green
                cmdCreateInvoice.Enabled = True
                cmdGetCustomers.Enabled = False
            Next
        Catch uError As Exception
            sErrDescription = uError.Message
            WriteAuditLogRecord(Me.Name, "DisplayClient", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "DisplayClient" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")

        End Try
Err_DisplayClient:
        If Err.Number <> 0 Then
        End If

    End Sub

    Private Sub cmdSyncToXero_Click(sender As Object, e As EventArgs) Handles cmdSyncToXero.Click

        ' Error Checking
        Try
            ' Match Contact Name and Update ContactNumber
            Dim uContacts = From Contacts In uRepository.Contacts Where Contacts.Name = lblContactNameValue.Text
            For Each uContact As Contact In uContacts
                uContact.ContactNumber = lblLandlordIDValue.Text
                uRepository.UpdateOrCreate(Of Contact)(uContact)
            Next
            cmdCreateInvoice.Enabled = True
            cmdSyncToXero.Enabled = False
        Catch uError As Exception
            sErrDescription = sErrDescription = uError.Message
            WriteAuditLogRecord(Me.Name, "cmdSyncToXero_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdSyncToXero_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End Try

    End Sub

    Private Sub lstXeroContacts_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstXeroContacts.SelectedIndexChanged
        lblContactNameValue.Text = lstXeroContacts.SelectedItem.substring(7, Len(lstXeroContacts.SelectedItem) - 7)
        cmdSyncToXero.Enabled = True

    End Sub

    Private Sub cmdCreateCustomer_Click(sender As Object, e As EventArgs) Handles cmdCreateCustomer.Click

        ' Error Checking
        On Error GoTo Err_cmdCreateCustomer_Click

        Dim uContact As New Contact
        uContact.IsCustomer = True
        uContact.ContactNumber = lblLandlordIDValue.Text
        If lblCompanyValue.Text = "PERSONAL" Then
            uContact.Name = lblSurnameValue.Text & ", " & lblForenameValue.Text
        Else
            uContact.Name = lblCompanyValue.Text
        End If
        uContact.FirstName = lblForenameValue.Text
        uContact.LastName = lblSurnameValue.Text
        uContact.EmailAddress = lblEMailAddressValue.Text
        uRepository.Create(Of Contact)(uContact)
        cmdCreateCustomer.Enabled = False
        cmdCreateInvoice.Enabled = True

Err_cmdCreateCustomer_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdCreateCustomer_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdCreateCustomer_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub cmdEmailInvoice_Click(sender As Object, e As EventArgs) Handles cmdEMailInvoice.Click
        ' Error Checking
        On Error GoTo Err_cmdEMailInvoice_Click

        If lblEMailAddressValue.Text = "" Then Exit Sub

        ' Create Report
        Dim ListReport = New FastReport.Report
        ListReport.Load(SystemPath & "Reports\ReportA4Invoice - " & My.Settings.CompanyName & ".frx")
        ListReport.SetParameterValue("CRMConnectionString", "Data Source=" & My.Settings.SQLServer & ";AttachDbFilename=;Initial Catalog=" & My.Settings.SQLDatabase & ";Integrated Security=False;Persist Security Info=False;User ID=CRMUser;Password=S0rtmypc!")
        ListReport.SetParameterValue("JobID", Val(lblJobIDValue.Text))
        ListReport.Prepare()

        ' Create Export File
        Dim PDFExport As FastReport.Export.Pdf.PDFExport = New FastReport.Export.Pdf.PDFExport
        ListReport.Export(PDFExport, SystemPath & "Reports\EMailA4Report" & lblJobIDValue.Text & ".pdf")

        ' Create EMail 
        Dim objOutlook As Object
        Dim objMailMessage As Outlook.MailItem
        objOutlook = CreateObject("Outlook.Application")
        objMailMessage = objOutlook.CreateItem(0)
        With objMailMessage
            .To = lblEMailAddressValue.Text
            .Body = "Hello," & Chr(13) & Chr(13)
            .Body = .Body & "We appreciate your business, thank you. We hope that your job has been completed to your satisfaction. Please find attached your invoice INV-" & lblJobIDValue.Text & " for " & FormatCurrency(ValCur(lblAmountDueValue.Text)) & "." & Chr(13) & Chr(13)
            .Body = .Body & "The amount outstanding of " & FormatCurrency(ValCur(lblAmountDueValue.Text)) & " is due on " & Format(DateAdd(DateInterval.Day, 30, Now), "dd MMM yyyy") & "." & Chr(13)
            .Body = .Body & "If you have any questions, please let us know." & Chr(13) & Chr(13)
            .Body = .Body & "Many thanks," & Chr(13) & Chr(13)
            .Body = .Body & "Finance at SortmyPC" & Chr(13)
            .Body = .Body & "Mackenzie Cottage" & Chr(13)
            .Body = .Body & "302 Colinton Road" & Chr(13)
            .Body = .Body & "Edinburgh" & Chr(13)
            .Body = .Body & "EH13 0LE" & Chr(13) & Chr(13)
            .Body = .Body & "Tel: 0131 477 2644" & Chr(13)
            .Body = .Body & "Email: finance@sortmypc.co.uk,"
            .Subject = "Your Job Ref: " & lblJobIDValue.Text & " - " & lblJobNameValue.Text
            .Attachments.Add(SystemPath & "Reports\EMailA4Report" & lblJobIDValue.Text & ".pdf")
            .Display()
            .Save()
            .Close(Outlook.OlInspectorClose.olDiscard)
        End With
        objMailMessage = Nothing
        objOutlook = Nothing

        ' Advise Action
        cmdEMailInvoice.Enabled = False
        Dim sMessage As String
        sMessage = Replace(">" & lblJobIDValue.Text & " - " & lblJobNameValue.Text & " to " & lblEMailAddress.Text, "'", "`")
        WriteAuditLogRecord(Me.Name, "cmdEMailInvoice_Click", "INFO", sMessage)
        MsgBox("The Email has been saved as a Draft", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "BMSSortmyPC - Action Confirmed")

Err_cmdEMailInvoice_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdEMailInvoice_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdEMailInvoice_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "BMSSortmyPC - Error Reporting")
        End If

    End Sub

    Sub CheckPayments()

        ' Error Checking
        On Error GoTo Err_CheckPayments

        With objPayment
            .JobID = Val(lblJobIDValue.Text)
            .GetXeroSalesInvoices()
            lblAmountInvoicedValue.Text = Format(.AmountInvoiced, "£0.00")
            If .AmountInvoiced > 0 Then
                cmdCreateInvoice.Enabled = False
                cmdEMailInvoice.Enabled = True
            End If
        End With

Err_CheckPayments:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "CheckPayments", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "CheckPayments" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Me.Close()

    End Sub

    Private Sub lblJobNameValue_Click(sender As Object, e As EventArgs) Handles lblJobNameValue.Click

    End Sub
End Class