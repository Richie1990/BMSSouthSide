<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmXeroCustomerCheck
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmXeroCustomerCheck))
        Me.cmdCreateInvoice = New System.Windows.Forms.Button()
        Me.cmdGetCustomers = New System.Windows.Forms.Button()
        Me.lstXeroContacts = New System.Windows.Forms.ListBox()
        Me.lblCompanyValue = New System.Windows.Forms.Label()
        Me.lblJobIDValue = New System.Windows.Forms.Label()
        Me.lblJobNo = New System.Windows.Forms.Label()
        Me.lblAddress3Value = New System.Windows.Forms.Label()
        Me.lblAddress2Value = New System.Windows.Forms.Label()
        Me.lblAddress1Value = New System.Windows.Forms.Label()
        Me.lblCompany = New System.Windows.Forms.Label()
        Me.lblPostcode = New System.Windows.Forms.Label()
        Me.lblAddress4 = New System.Windows.Forms.Label()
        Me.lblAddress3 = New System.Windows.Forms.Label()
        Me.lblAddress2 = New System.Windows.Forms.Label()
        Me.lblAddress1 = New System.Windows.Forms.Label()
        Me.lblLandlordIDValue = New System.Windows.Forms.Label()
        Me.lblClientID = New System.Windows.Forms.Label()
        Me.lblAddress4Value = New System.Windows.Forms.Label()
        Me.lblPostcodeValue = New System.Windows.Forms.Label()
        Me.lblEMailAddress = New System.Windows.Forms.Label()
        Me.lblMobileNo = New System.Windows.Forms.Label()
        Me.lblPhoneNo = New System.Windows.Forms.Label()
        Me.lblPhoneNoValue = New System.Windows.Forms.Label()
        Me.lblMobileNoValue = New System.Windows.Forms.Label()
        Me.lblEMailAddressValue = New System.Windows.Forms.Label()
        Me.lblFullName = New System.Windows.Forms.Label()
        Me.lblTitleValue = New System.Windows.Forms.Label()
        Me.lblForenameValue = New System.Windows.Forms.Label()
        Me.lblSurnameValue = New System.Windows.Forms.Label()
        Me.cmdSyncToXero = New System.Windows.Forms.Button()
        Me.cmdCreateCustomer = New System.Windows.Forms.Button()
        Me.lblAmountDueValue = New System.Windows.Forms.Label()
        Me.lblAmountDue = New System.Windows.Forms.Label()
        Me.lblContactNameValue = New System.Windows.Forms.Label()
        Me.pnlCustomer = New System.Windows.Forms.Panel()
        Me.pnlInvoice = New System.Windows.Forms.Panel()
        Me.pnlPayment = New System.Windows.Forms.Panel()
        Me.cmdEMailInvoice = New System.Windows.Forms.Button()
        Me.lblAmountInvoicedValue = New System.Windows.Forms.Label()
        Me.lblAmountInvoiced = New System.Windows.Forms.Label()
        Me.pnlSummary = New System.Windows.Forms.Panel()
        Me.lblJobName = New System.Windows.Forms.Label()
        Me.lblJobNameValue = New System.Windows.Forms.Label()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.pnlCustomer.SuspendLayout()
        Me.pnlInvoice.SuspendLayout()
        Me.pnlPayment.SuspendLayout()
        Me.pnlSummary.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCreateInvoice
        '
        Me.cmdCreateInvoice.Enabled = False
        Me.cmdCreateInvoice.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmdCreateInvoice.Image = CType(resources.GetObject("cmdCreateInvoice.Image"), System.Drawing.Image)
        Me.cmdCreateInvoice.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCreateInvoice.Location = New System.Drawing.Point(8, 10)
        Me.cmdCreateInvoice.Name = "cmdCreateInvoice"
        Me.cmdCreateInvoice.Size = New System.Drawing.Size(476, 45)
        Me.cmdCreateInvoice.TabIndex = 0
        Me.cmdCreateInvoice.Text = "Create Xero Invoice"
        Me.cmdCreateInvoice.UseVisualStyleBackColor = True
        '
        'cmdGetCustomers
        '
        Me.cmdGetCustomers.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmdGetCustomers.Location = New System.Drawing.Point(9, 9)
        Me.cmdGetCustomers.Name = "cmdGetCustomers"
        Me.cmdGetCustomers.Size = New System.Drawing.Size(476, 28)
        Me.cmdGetCustomers.TabIndex = 1
        Me.cmdGetCustomers.Text = "Get Matching Customers"
        Me.cmdGetCustomers.UseVisualStyleBackColor = True
        '
        'lstXeroContacts
        '
        Me.lstXeroContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lstXeroContacts.FormattingEnabled = True
        Me.lstXeroContacts.ItemHeight = 16
        Me.lstXeroContacts.Location = New System.Drawing.Point(9, 43)
        Me.lstXeroContacts.Name = "lstXeroContacts"
        Me.lstXeroContacts.Size = New System.Drawing.Size(254, 68)
        Me.lstXeroContacts.TabIndex = 2
        '
        'lblCompanyValue
        '
        Me.lblCompanyValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblCompanyValue.Location = New System.Drawing.Point(96, 9)
        Me.lblCompanyValue.Name = "lblCompanyValue"
        Me.lblCompanyValue.Size = New System.Drawing.Size(301, 16)
        Me.lblCompanyValue.TabIndex = 355
        Me.lblCompanyValue.Text = "xxx"
        '
        'lblJobIDValue
        '
        Me.lblJobIDValue.AutoSize = True
        Me.lblJobIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblJobIDValue.Location = New System.Drawing.Point(97, 19)
        Me.lblJobIDValue.Name = "lblJobIDValue"
        Me.lblJobIDValue.Size = New System.Drawing.Size(43, 16)
        Me.lblJobIDValue.TabIndex = 353
        Me.lblJobIDValue.Text = "00000"
        '
        'lblJobNo
        '
        Me.lblJobNo.AutoSize = True
        Me.lblJobNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobNo.Location = New System.Drawing.Point(32, 18)
        Me.lblJobNo.Name = "lblJobNo"
        Me.lblJobNo.Size = New System.Drawing.Size(62, 16)
        Me.lblJobNo.TabIndex = 352
        Me.lblJobNo.Text = "Job No:"
        '
        'lblAddress3Value
        '
        Me.lblAddress3Value.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblAddress3Value.Location = New System.Drawing.Point(129, 118)
        Me.lblAddress3Value.Name = "lblAddress3Value"
        Me.lblAddress3Value.Size = New System.Drawing.Size(322, 16)
        Me.lblAddress3Value.TabIndex = 358
        Me.lblAddress3Value.Text = "xxx"
        '
        'lblAddress2Value
        '
        Me.lblAddress2Value.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblAddress2Value.Location = New System.Drawing.Point(129, 90)
        Me.lblAddress2Value.Name = "lblAddress2Value"
        Me.lblAddress2Value.Size = New System.Drawing.Size(322, 16)
        Me.lblAddress2Value.TabIndex = 360
        Me.lblAddress2Value.Text = "xxx"
        '
        'lblAddress1Value
        '
        Me.lblAddress1Value.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblAddress1Value.Location = New System.Drawing.Point(129, 62)
        Me.lblAddress1Value.Name = "lblAddress1Value"
        Me.lblAddress1Value.Size = New System.Drawing.Size(322, 16)
        Me.lblAddress1Value.TabIndex = 362
        Me.lblAddress1Value.Text = "xxx"
        '
        'lblCompany
        '
        Me.lblCompany.AutoSize = True
        Me.lblCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompany.Location = New System.Drawing.Point(8, 9)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(77, 16)
        Me.lblCompany.TabIndex = 369
        Me.lblCompany.Text = "Company:"
        '
        'lblPostcode
        '
        Me.lblPostcode.AutoSize = True
        Me.lblPostcode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPostcode.Location = New System.Drawing.Point(288, 146)
        Me.lblPostcode.Name = "lblPostcode"
        Me.lblPostcode.Size = New System.Drawing.Size(78, 16)
        Me.lblPostcode.TabIndex = 366
        Me.lblPostcode.Text = "Postcode:"
        '
        'lblAddress4
        '
        Me.lblAddress4.AutoSize = True
        Me.lblAddress4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress4.Location = New System.Drawing.Point(8, 146)
        Me.lblAddress4.Name = "lblAddress4"
        Me.lblAddress4.Size = New System.Drawing.Size(49, 16)
        Me.lblAddress4.TabIndex = 367
        Me.lblAddress4.Text = "Town:"
        '
        'lblAddress3
        '
        Me.lblAddress3.AutoSize = True
        Me.lblAddress3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress3.Location = New System.Drawing.Point(8, 118)
        Me.lblAddress3.Name = "lblAddress3"
        Me.lblAddress3.Size = New System.Drawing.Size(115, 16)
        Me.lblAddress3.TabIndex = 365
        Me.lblAddress3.Text = "Address Line 3:"
        '
        'lblAddress2
        '
        Me.lblAddress2.AutoSize = True
        Me.lblAddress2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress2.Location = New System.Drawing.Point(8, 90)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(115, 16)
        Me.lblAddress2.TabIndex = 364
        Me.lblAddress2.Text = "Address Line 2:"
        '
        'lblAddress1
        '
        Me.lblAddress1.AutoSize = True
        Me.lblAddress1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress1.Location = New System.Drawing.Point(8, 62)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New System.Drawing.Size(115, 16)
        Me.lblAddress1.TabIndex = 363
        Me.lblAddress1.Text = "Address Line 1:"
        '
        'lblLandlordIDValue
        '
        Me.lblLandlordIDValue.AutoSize = True
        Me.lblLandlordIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLandlordIDValue.Location = New System.Drawing.Point(98, 44)
        Me.lblLandlordIDValue.Name = "lblLandlordIDValue"
        Me.lblLandlordIDValue.Size = New System.Drawing.Size(36, 16)
        Me.lblLandlordIDValue.TabIndex = 371
        Me.lblLandlordIDValue.Text = "0000"
        '
        'lblClientID
        '
        Me.lblClientID.AutoSize = True
        Me.lblClientID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClientID.Location = New System.Drawing.Point(24, 44)
        Me.lblClientID.Name = "lblClientID"
        Me.lblClientID.Size = New System.Drawing.Size(70, 16)
        Me.lblClientID.TabIndex = 370
        Me.lblClientID.Text = "Client ID:"
        '
        'lblAddress4Value
        '
        Me.lblAddress4Value.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblAddress4Value.Location = New System.Drawing.Point(63, 146)
        Me.lblAddress4Value.Name = "lblAddress4Value"
        Me.lblAddress4Value.Size = New System.Drawing.Size(200, 16)
        Me.lblAddress4Value.TabIndex = 372
        Me.lblAddress4Value.Text = "xxx"
        '
        'lblPostcodeValue
        '
        Me.lblPostcodeValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblPostcodeValue.Location = New System.Drawing.Point(372, 146)
        Me.lblPostcodeValue.Name = "lblPostcodeValue"
        Me.lblPostcodeValue.Size = New System.Drawing.Size(79, 16)
        Me.lblPostcodeValue.TabIndex = 373
        Me.lblPostcodeValue.Text = "xxx"
        '
        'lblEMailAddress
        '
        Me.lblEMailAddress.AutoSize = True
        Me.lblEMailAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEMailAddress.Location = New System.Drawing.Point(10, 203)
        Me.lblEMailAddress.Name = "lblEMailAddress"
        Me.lblEMailAddress.Size = New System.Drawing.Size(113, 16)
        Me.lblEMailAddress.TabIndex = 380
        Me.lblEMailAddress.Text = "EMail Address:"
        '
        'lblMobileNo
        '
        Me.lblMobileNo.AutoSize = True
        Me.lblMobileNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMobileNo.Location = New System.Drawing.Point(263, 173)
        Me.lblMobileNo.Name = "lblMobileNo"
        Me.lblMobileNo.Size = New System.Drawing.Size(83, 16)
        Me.lblMobileNo.TabIndex = 379
        Me.lblMobileNo.Text = "Mobile No:"
        '
        'lblPhoneNo
        '
        Me.lblPhoneNo.AutoSize = True
        Me.lblPhoneNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhoneNo.Location = New System.Drawing.Point(10, 173)
        Me.lblPhoneNo.Name = "lblPhoneNo"
        Me.lblPhoneNo.Size = New System.Drawing.Size(80, 16)
        Me.lblPhoneNo.TabIndex = 378
        Me.lblPhoneNo.Text = "Phone No:"
        '
        'lblPhoneNoValue
        '
        Me.lblPhoneNoValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblPhoneNoValue.Location = New System.Drawing.Point(96, 173)
        Me.lblPhoneNoValue.Name = "lblPhoneNoValue"
        Me.lblPhoneNoValue.Size = New System.Drawing.Size(133, 16)
        Me.lblPhoneNoValue.TabIndex = 381
        Me.lblPhoneNoValue.Text = "xxx"
        '
        'lblMobileNoValue
        '
        Me.lblMobileNoValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblMobileNoValue.Location = New System.Drawing.Point(352, 173)
        Me.lblMobileNoValue.Name = "lblMobileNoValue"
        Me.lblMobileNoValue.Size = New System.Drawing.Size(133, 16)
        Me.lblMobileNoValue.TabIndex = 382
        Me.lblMobileNoValue.Text = "xxx"
        '
        'lblEMailAddressValue
        '
        Me.lblEMailAddressValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblEMailAddressValue.Location = New System.Drawing.Point(129, 203)
        Me.lblEMailAddressValue.Name = "lblEMailAddressValue"
        Me.lblEMailAddressValue.Size = New System.Drawing.Size(356, 16)
        Me.lblEMailAddressValue.TabIndex = 383
        Me.lblEMailAddressValue.Text = "xxx"
        '
        'lblFullName
        '
        Me.lblFullName.AutoSize = True
        Me.lblFullName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFullName.Location = New System.Drawing.Point(8, 34)
        Me.lblFullName.Name = "lblFullName"
        Me.lblFullName.Size = New System.Drawing.Size(82, 16)
        Me.lblFullName.TabIndex = 385
        Me.lblFullName.Text = "Full Name:"
        '
        'lblTitleValue
        '
        Me.lblTitleValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblTitleValue.Location = New System.Drawing.Point(96, 34)
        Me.lblTitleValue.Name = "lblTitleValue"
        Me.lblTitleValue.Size = New System.Drawing.Size(40, 16)
        Me.lblTitleValue.TabIndex = 384
        Me.lblTitleValue.Text = "xxx"
        '
        'lblForenameValue
        '
        Me.lblForenameValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblForenameValue.Location = New System.Drawing.Point(142, 34)
        Me.lblForenameValue.Name = "lblForenameValue"
        Me.lblForenameValue.Size = New System.Drawing.Size(100, 16)
        Me.lblForenameValue.TabIndex = 386
        Me.lblForenameValue.Text = "xxx"
        '
        'lblSurnameValue
        '
        Me.lblSurnameValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblSurnameValue.Location = New System.Drawing.Point(248, 34)
        Me.lblSurnameValue.Name = "lblSurnameValue"
        Me.lblSurnameValue.Size = New System.Drawing.Size(98, 16)
        Me.lblSurnameValue.TabIndex = 387
        Me.lblSurnameValue.Text = "xxx"
        '
        'cmdSyncToXero
        '
        Me.cmdSyncToXero.Enabled = False
        Me.cmdSyncToXero.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmdSyncToXero.Location = New System.Drawing.Point(348, 84)
        Me.cmdSyncToXero.Name = "cmdSyncToXero"
        Me.cmdSyncToXero.Size = New System.Drawing.Size(137, 28)
        Me.cmdSyncToXero.TabIndex = 388
        Me.cmdSyncToXero.Text = "<<< Sync to Xero >>>"
        Me.cmdSyncToXero.UseVisualStyleBackColor = True
        '
        'cmdCreateCustomer
        '
        Me.cmdCreateCustomer.Enabled = False
        Me.cmdCreateCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmdCreateCustomer.Location = New System.Drawing.Point(348, 43)
        Me.cmdCreateCustomer.Name = "cmdCreateCustomer"
        Me.cmdCreateCustomer.Size = New System.Drawing.Size(137, 28)
        Me.cmdCreateCustomer.TabIndex = 389
        Me.cmdCreateCustomer.Text = "Create Customer"
        Me.cmdCreateCustomer.UseVisualStyleBackColor = True
        '
        'lblAmountDueValue
        '
        Me.lblAmountDueValue.AutoSize = True
        Me.lblAmountDueValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmountDueValue.Location = New System.Drawing.Point(321, 18)
        Me.lblAmountDueValue.Name = "lblAmountDueValue"
        Me.lblAmountDueValue.Size = New System.Drawing.Size(39, 16)
        Me.lblAmountDueValue.TabIndex = 391
        Me.lblAmountDueValue.Text = "£0.00"
        '
        'lblAmountDue
        '
        Me.lblAmountDue.AutoSize = True
        Me.lblAmountDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmountDue.Location = New System.Drawing.Point(217, 18)
        Me.lblAmountDue.Name = "lblAmountDue"
        Me.lblAmountDue.Size = New System.Drawing.Size(95, 16)
        Me.lblAmountDue.TabIndex = 390
        Me.lblAmountDue.Text = "Amount Due:"
        '
        'lblContactNameValue
        '
        Me.lblContactNameValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.lblContactNameValue.Location = New System.Drawing.Point(9, 115)
        Me.lblContactNameValue.Name = "lblContactNameValue"
        Me.lblContactNameValue.Size = New System.Drawing.Size(301, 16)
        Me.lblContactNameValue.TabIndex = 392
        Me.lblContactNameValue.Text = "xxx"
        '
        'pnlCustomer
        '
        Me.pnlCustomer.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.pnlCustomer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCustomer.Controls.Add(Me.lblContactNameValue)
        Me.pnlCustomer.Controls.Add(Me.cmdCreateCustomer)
        Me.pnlCustomer.Controls.Add(Me.cmdSyncToXero)
        Me.pnlCustomer.Controls.Add(Me.lstXeroContacts)
        Me.pnlCustomer.Controls.Add(Me.cmdGetCustomers)
        Me.pnlCustomer.Location = New System.Drawing.Point(3, 337)
        Me.pnlCustomer.Name = "pnlCustomer"
        Me.pnlCustomer.Size = New System.Drawing.Size(497, 142)
        Me.pnlCustomer.TabIndex = 395
        '
        'pnlInvoice
        '
        Me.pnlInvoice.BackColor = System.Drawing.Color.SandyBrown
        Me.pnlInvoice.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlInvoice.Controls.Add(Me.cmdCreateInvoice)
        Me.pnlInvoice.Location = New System.Drawing.Point(3, 484)
        Me.pnlInvoice.Name = "pnlInvoice"
        Me.pnlInvoice.Size = New System.Drawing.Size(497, 70)
        Me.pnlInvoice.TabIndex = 396
        '
        'pnlPayment
        '
        Me.pnlPayment.BackColor = System.Drawing.Color.PaleGreen
        Me.pnlPayment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPayment.Controls.Add(Me.cmdEMailInvoice)
        Me.pnlPayment.Location = New System.Drawing.Point(3, 562)
        Me.pnlPayment.Name = "pnlPayment"
        Me.pnlPayment.Size = New System.Drawing.Size(496, 70)
        Me.pnlPayment.TabIndex = 398
        '
        'cmdEMailInvoice
        '
        Me.cmdEMailInvoice.Enabled = False
        Me.cmdEMailInvoice.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmdEMailInvoice.Image = CType(resources.GetObject("cmdEMailInvoice.Image"), System.Drawing.Image)
        Me.cmdEMailInvoice.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEMailInvoice.Location = New System.Drawing.Point(8, 12)
        Me.cmdEMailInvoice.Name = "cmdEMailInvoice"
        Me.cmdEMailInvoice.Size = New System.Drawing.Size(476, 45)
        Me.cmdEMailInvoice.TabIndex = 393
        Me.cmdEMailInvoice.Text = "EMail A4 PDF Invoice to Client"
        Me.cmdEMailInvoice.UseVisualStyleBackColor = True
        '
        'lblAmountInvoicedValue
        '
        Me.lblAmountInvoicedValue.AutoSize = True
        Me.lblAmountInvoicedValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmountInvoicedValue.Location = New System.Drawing.Point(349, 44)
        Me.lblAmountInvoicedValue.Name = "lblAmountInvoicedValue"
        Me.lblAmountInvoicedValue.Size = New System.Drawing.Size(39, 16)
        Me.lblAmountInvoicedValue.TabIndex = 400
        Me.lblAmountInvoicedValue.Text = "£0.00"
        '
        'lblAmountInvoiced
        '
        Me.lblAmountInvoiced.AutoSize = True
        Me.lblAmountInvoiced.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmountInvoiced.Location = New System.Drawing.Point(217, 44)
        Me.lblAmountInvoiced.Name = "lblAmountInvoiced"
        Me.lblAmountInvoiced.Size = New System.Drawing.Size(126, 16)
        Me.lblAmountInvoiced.TabIndex = 399
        Me.lblAmountInvoiced.Text = "Amount Invoiced:"
        '
        'pnlSummary
        '
        Me.pnlSummary.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSummary.Controls.Add(Me.lblSurnameValue)
        Me.pnlSummary.Controls.Add(Me.lblForenameValue)
        Me.pnlSummary.Controls.Add(Me.lblFullName)
        Me.pnlSummary.Controls.Add(Me.lblTitleValue)
        Me.pnlSummary.Controls.Add(Me.lblEMailAddressValue)
        Me.pnlSummary.Controls.Add(Me.lblMobileNoValue)
        Me.pnlSummary.Controls.Add(Me.lblPhoneNoValue)
        Me.pnlSummary.Controls.Add(Me.lblEMailAddress)
        Me.pnlSummary.Controls.Add(Me.lblMobileNo)
        Me.pnlSummary.Controls.Add(Me.lblPhoneNo)
        Me.pnlSummary.Controls.Add(Me.lblPostcodeValue)
        Me.pnlSummary.Controls.Add(Me.lblAddress4Value)
        Me.pnlSummary.Controls.Add(Me.lblCompany)
        Me.pnlSummary.Controls.Add(Me.lblPostcode)
        Me.pnlSummary.Controls.Add(Me.lblAddress4)
        Me.pnlSummary.Controls.Add(Me.lblAddress3)
        Me.pnlSummary.Controls.Add(Me.lblAddress2)
        Me.pnlSummary.Controls.Add(Me.lblAddress1)
        Me.pnlSummary.Controls.Add(Me.lblAddress1Value)
        Me.pnlSummary.Controls.Add(Me.lblAddress2Value)
        Me.pnlSummary.Controls.Add(Me.lblAddress3Value)
        Me.pnlSummary.Controls.Add(Me.lblCompanyValue)
        Me.pnlSummary.Location = New System.Drawing.Point(3, 99)
        Me.pnlSummary.Name = "pnlSummary"
        Me.pnlSummary.Size = New System.Drawing.Size(497, 227)
        Me.pnlSummary.TabIndex = 403
        '
        'lblJobName
        '
        Me.lblJobName.AutoSize = True
        Me.lblJobName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobName.Location = New System.Drawing.Point(11, 70)
        Me.lblJobName.Name = "lblJobName"
        Me.lblJobName.Size = New System.Drawing.Size(83, 16)
        Me.lblJobName.TabIndex = 404
        Me.lblJobName.Text = "Job Name:"
        '
        'lblJobNameValue
        '
        Me.lblJobNameValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobNameValue.Location = New System.Drawing.Point(97, 70)
        Me.lblJobNameValue.Name = "lblJobNameValue"
        Me.lblJobNameValue.Size = New System.Drawing.Size(402, 16)
        Me.lblJobNameValue.TabIndex = 405
        Me.lblJobNameValue.Text = "xxx"
        '
        'cmdClose
        '
        Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmdClose.Image = CType(resources.GetObject("cmdClose.Image"), System.Drawing.Image)
        Me.cmdClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClose.Location = New System.Drawing.Point(429, 639)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(60, 60)
        Me.cmdClose.TabIndex = 406
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'frmXeroCustomerCheck
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(502, 706)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.lblJobNameValue)
        Me.Controls.Add(Me.lblJobName)
        Me.Controls.Add(Me.pnlSummary)
        Me.Controls.Add(Me.lblAmountInvoicedValue)
        Me.Controls.Add(Me.lblAmountInvoiced)
        Me.Controls.Add(Me.pnlPayment)
        Me.Controls.Add(Me.pnlInvoice)
        Me.Controls.Add(Me.pnlCustomer)
        Me.Controls.Add(Me.lblAmountDueValue)
        Me.Controls.Add(Me.lblAmountDue)
        Me.Controls.Add(Me.lblLandlordIDValue)
        Me.Controls.Add(Me.lblClientID)
        Me.Controls.Add(Me.lblJobIDValue)
        Me.Controls.Add(Me.lblJobNo)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmXeroCustomerCheck"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmXeroCustomerCheck"
        Me.pnlCustomer.ResumeLayout(False)
        Me.pnlInvoice.ResumeLayout(False)
        Me.pnlPayment.ResumeLayout(False)
        Me.pnlSummary.ResumeLayout(False)
        Me.pnlSummary.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdCreateInvoice As System.Windows.Forms.Button
    Friend WithEvents cmdGetCustomers As System.Windows.Forms.Button
    Friend WithEvents lstXeroContacts As System.Windows.Forms.ListBox
    Friend WithEvents lblCompanyValue As System.Windows.Forms.Label
    Friend WithEvents lblJobIDValue As System.Windows.Forms.Label
    Friend WithEvents lblJobNo As System.Windows.Forms.Label
    Friend WithEvents lblAddress3Value As System.Windows.Forms.Label
    Friend WithEvents lblAddress2Value As System.Windows.Forms.Label
    Friend WithEvents lblAddress1Value As System.Windows.Forms.Label
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents lblPostcode As System.Windows.Forms.Label
    Friend WithEvents lblAddress4 As System.Windows.Forms.Label
    Friend WithEvents lblAddress3 As System.Windows.Forms.Label
    Friend WithEvents lblAddress2 As System.Windows.Forms.Label
    Friend WithEvents lblAddress1 As System.Windows.Forms.Label
    Friend WithEvents lblLandlordIDValue As System.Windows.Forms.Label
    Friend WithEvents lblClientID As System.Windows.Forms.Label
    Friend WithEvents lblAddress4Value As System.Windows.Forms.Label
    Friend WithEvents lblPostcodeValue As System.Windows.Forms.Label
    Friend WithEvents lblEMailAddress As System.Windows.Forms.Label
    Friend WithEvents lblMobileNo As System.Windows.Forms.Label
    Friend WithEvents lblPhoneNo As System.Windows.Forms.Label
    Friend WithEvents lblPhoneNoValue As System.Windows.Forms.Label
    Friend WithEvents lblMobileNoValue As System.Windows.Forms.Label
    Friend WithEvents lblEMailAddressValue As System.Windows.Forms.Label
    Friend WithEvents lblFullName As System.Windows.Forms.Label
    Friend WithEvents lblTitleValue As System.Windows.Forms.Label
    Friend WithEvents lblForenameValue As System.Windows.Forms.Label
    Friend WithEvents lblSurnameValue As System.Windows.Forms.Label
    Friend WithEvents cmdSyncToXero As System.Windows.Forms.Button
    Friend WithEvents cmdCreateCustomer As System.Windows.Forms.Button
    Friend WithEvents lblAmountDueValue As System.Windows.Forms.Label
    Friend WithEvents lblAmountDue As System.Windows.Forms.Label
    Friend WithEvents lblContactNameValue As System.Windows.Forms.Label
    Friend WithEvents pnlCustomer As System.Windows.Forms.Panel
    Friend WithEvents pnlInvoice As System.Windows.Forms.Panel
    Friend WithEvents pnlPayment As System.Windows.Forms.Panel
    Friend WithEvents lblAmountInvoicedValue As System.Windows.Forms.Label
    Friend WithEvents lblAmountInvoiced As System.Windows.Forms.Label
    Friend WithEvents pnlSummary As System.Windows.Forms.Panel
    Friend WithEvents cmdEMailInvoice As System.Windows.Forms.Button
    Friend WithEvents lblJobName As System.Windows.Forms.Label
    Friend WithEvents lblJobNameValue As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
End Class
