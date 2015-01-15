<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLandlord
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLandlord))
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileClose = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLandlord = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewLandlord = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSaveLandlord = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRemoveLandlord = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuJob = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewProperty = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditProperty = New System.Windows.Forms.ToolStripMenuItem()
        Me.lblMessagePanel = New System.Windows.Forms.Label()
        Me.lblLandlordIDValue = New System.Windows.Forms.Label()
        Me.lblLandlordID = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.lblClients = New System.Windows.Forms.Label()
        Me.txtFilter = New System.Windows.Forms.TextBox()
        Me.pnlLandlord = New System.Windows.Forms.Panel()
        Me.lblLandlordName = New System.Windows.Forms.Label()
        Me.txtLandlordName = New System.Windows.Forms.TextBox()
        Me.lblPlaceofBirth2 = New System.Windows.Forms.Label()
        Me.txtPlaceofBirth2 = New System.Windows.Forms.TextBox()
        Me.lblPlaceofBirth1 = New System.Windows.Forms.Label()
        Me.txtPlaceofBirth1 = New System.Windows.Forms.TextBox()
        Me.lblDOB2 = New System.Windows.Forms.Label()
        Me.dtpDOB2 = New System.Windows.Forms.DateTimePicker()
        Me.lblDOB1 = New System.Windows.Forms.Label()
        Me.dtpDOB1 = New System.Windows.Forms.DateTimePicker()
        Me.lblDateJoined = New System.Windows.Forms.Label()
        Me.dtpDateJoined = New System.Windows.Forms.DateTimePicker()
        Me.txtAccountNo = New System.Windows.Forms.TextBox()
        Me.lblAccountNo = New System.Windows.Forms.Label()
        Me.txtSortcode = New System.Windows.Forms.TextBox()
        Me.txtBankName = New System.Windows.Forms.TextBox()
        Me.lblSortcode = New System.Windows.Forms.Label()
        Me.lblBankName = New System.Windows.Forms.Label()
        Me.lblCommsTitle = New System.Windows.Forms.Label()
        Me.txtCommsTitle = New System.Windows.Forms.TextBox()
        Me.lblSurname2 = New System.Windows.Forms.Label()
        Me.lblTitleForename2 = New System.Windows.Forms.Label()
        Me.txtSurname2 = New System.Windows.Forms.TextBox()
        Me.txtForename2 = New System.Windows.Forms.TextBox()
        Me.txtTitle2 = New System.Windows.Forms.TextBox()
        Me.lblStar4 = New System.Windows.Forms.Label()
        Me.lblStar3 = New System.Windows.Forms.Label()
        Me.lblSource = New System.Windows.Forms.Label()
        Me.cmbSource = New System.Windows.Forms.ComboBox()
        Me.rtbLandlordNotes = New System.Windows.Forms.RichTextBox()
        Me.lblLandlordNotes = New System.Windows.Forms.Label()
        Me.txtEMailAddress = New System.Windows.Forms.TextBox()
        Me.lblEMailAddress = New System.Windows.Forms.Label()
        Me.txtMobileNo = New System.Windows.Forms.TextBox()
        Me.lblMobileNo = New System.Windows.Forms.Label()
        Me.txtPhoneNo = New System.Windows.Forms.TextBox()
        Me.lblPhoneNo = New System.Windows.Forms.Label()
        Me.txtPostCode = New System.Windows.Forms.TextBox()
        Me.txtTown = New System.Windows.Forms.TextBox()
        Me.txtAddress3 = New System.Windows.Forms.TextBox()
        Me.txtAddress2 = New System.Windows.Forms.TextBox()
        Me.txtAddress1 = New System.Windows.Forms.TextBox()
        Me.lblPostCode = New System.Windows.Forms.Label()
        Me.lblTown = New System.Windows.Forms.Label()
        Me.lblAddress3 = New System.Windows.Forms.Label()
        Me.lblAddress2 = New System.Windows.Forms.Label()
        Me.lblAddress1 = New System.Windows.Forms.Label()
        Me.lblSurname1 = New System.Windows.Forms.Label()
        Me.lblForenameTitle1 = New System.Windows.Forms.Label()
        Me.txtSurname1 = New System.Windows.Forms.TextBox()
        Me.txtForename1 = New System.Windows.Forms.TextBox()
        Me.txtTitle1 = New System.Windows.Forms.TextBox()
        Me.tsbClient = New System.Windows.Forms.ToolStrip()
        Me.tsbNewLandlord = New System.Windows.Forms.ToolStripButton()
        Me.tsbSaveLandlord = New System.Windows.Forms.ToolStripButton()
        Me.tsbRemoveLandlord = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbNewProperty = New System.Windows.Forms.ToolStripButton()
        Me.tsbEditProperty = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbClose = New System.Windows.Forms.ToolStripButton()
        Me.ProgrammeListBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.lblFilter = New System.Windows.Forms.Label()
        Me.tvwLandlordList = New System.Windows.Forms.TreeView()
        Me.MenuStrip1.SuspendLayout()
        Me.pnlLandlord.SuspendLayout()
        Me.tsbClient.SuspendLayout()
        CType(Me.ProgrammeListBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileClose})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileClose
        '
        Me.mnuFileClose.Name = "mnuFileClose"
        Me.mnuFileClose.Size = New System.Drawing.Size(103, 22)
        Me.mnuFileClose.Text = "&Close"
        '
        'mnuLandlord
        '
        Me.mnuLandlord.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewLandlord, Me.mnuSaveLandlord, Me.mnuRemoveLandlord})
        Me.mnuLandlord.Name = "mnuLandlord"
        Me.mnuLandlord.Size = New System.Drawing.Size(66, 20)
        Me.mnuLandlord.Text = "&Landlord"
        '
        'mnuNewLandlord
        '
        Me.mnuNewLandlord.Name = "mnuNewLandlord"
        Me.mnuNewLandlord.Size = New System.Drawing.Size(148, 22)
        Me.mnuNewLandlord.Text = "&New Landlord"
        '
        'mnuSaveLandlord
        '
        Me.mnuSaveLandlord.Name = "mnuSaveLandlord"
        Me.mnuSaveLandlord.Size = New System.Drawing.Size(148, 22)
        Me.mnuSaveLandlord.Text = "&Save"
        '
        'mnuRemoveLandlord
        '
        Me.mnuRemoveLandlord.Name = "mnuRemoveLandlord"
        Me.mnuRemoveLandlord.Size = New System.Drawing.Size(148, 22)
        Me.mnuRemoveLandlord.Text = "&Remove"
        '
        'mnuJob
        '
        Me.mnuJob.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewProperty, Me.mnuEditProperty})
        Me.mnuJob.Enabled = False
        Me.mnuJob.Name = "mnuJob"
        Me.mnuJob.Size = New System.Drawing.Size(64, 20)
        Me.mnuJob.Text = "&Property"
        '
        'mnuNewProperty
        '
        Me.mnuNewProperty.Name = "mnuNewProperty"
        Me.mnuNewProperty.Size = New System.Drawing.Size(98, 22)
        Me.mnuNewProperty.Text = "&New"
        '
        'mnuEditProperty
        '
        Me.mnuEditProperty.Name = "mnuEditProperty"
        Me.mnuEditProperty.Size = New System.Drawing.Size(98, 22)
        Me.mnuEditProperty.Text = "&Edit"
        '
        'lblMessagePanel
        '
        Me.lblMessagePanel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessagePanel.ForeColor = System.Drawing.Color.Red
        Me.lblMessagePanel.Location = New System.Drawing.Point(12, 717)
        Me.lblMessagePanel.Name = "lblMessagePanel"
        Me.lblMessagePanel.Size = New System.Drawing.Size(539, 35)
        Me.lblMessagePanel.TabIndex = 187
        Me.lblMessagePanel.Text = "Message Panel"
        Me.lblMessagePanel.Visible = False
        '
        'lblLandlordIDValue
        '
        Me.lblLandlordIDValue.AutoSize = True
        Me.lblLandlordIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLandlordIDValue.Location = New System.Drawing.Point(512, 79)
        Me.lblLandlordIDValue.Name = "lblLandlordIDValue"
        Me.lblLandlordIDValue.Size = New System.Drawing.Size(36, 16)
        Me.lblLandlordIDValue.TabIndex = 182
        Me.lblLandlordIDValue.Text = "0000"
        '
        'lblLandlordID
        '
        Me.lblLandlordID.AutoSize = True
        Me.lblLandlordID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLandlordID.Location = New System.Drawing.Point(414, 80)
        Me.lblLandlordID.Name = "lblLandlordID"
        Me.lblLandlordID.Size = New System.Drawing.Size(92, 16)
        Me.lblLandlordID.TabIndex = 162
        Me.lblLandlordID.Text = "Landlord ID:"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuLandlord, Me.mnuJob})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1264, 24)
        Me.MenuStrip1.TabIndex = 191
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'lblClients
        '
        Me.lblClients.AutoSize = True
        Me.lblClients.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClients.Location = New System.Drawing.Point(12, 79)
        Me.lblClients.Name = "lblClients"
        Me.lblClients.Size = New System.Drawing.Size(77, 16)
        Me.lblClients.TabIndex = 266
        Me.lblClients.Text = "Landlords"
        '
        'txtFilter
        '
        Me.txtFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFilter.Location = New System.Drawing.Point(210, 77)
        Me.txtFilter.Name = "txtFilter"
        Me.txtFilter.Size = New System.Drawing.Size(178, 22)
        Me.txtFilter.TabIndex = 269
        '
        'pnlLandlord
        '
        Me.pnlLandlord.Controls.Add(Me.lblLandlordName)
        Me.pnlLandlord.Controls.Add(Me.txtLandlordName)
        Me.pnlLandlord.Controls.Add(Me.lblPlaceofBirth2)
        Me.pnlLandlord.Controls.Add(Me.txtPlaceofBirth2)
        Me.pnlLandlord.Controls.Add(Me.lblPlaceofBirth1)
        Me.pnlLandlord.Controls.Add(Me.txtPlaceofBirth1)
        Me.pnlLandlord.Controls.Add(Me.lblDOB2)
        Me.pnlLandlord.Controls.Add(Me.dtpDOB2)
        Me.pnlLandlord.Controls.Add(Me.lblDOB1)
        Me.pnlLandlord.Controls.Add(Me.dtpDOB1)
        Me.pnlLandlord.Controls.Add(Me.lblDateJoined)
        Me.pnlLandlord.Controls.Add(Me.dtpDateJoined)
        Me.pnlLandlord.Controls.Add(Me.txtAccountNo)
        Me.pnlLandlord.Controls.Add(Me.lblAccountNo)
        Me.pnlLandlord.Controls.Add(Me.txtSortcode)
        Me.pnlLandlord.Controls.Add(Me.txtBankName)
        Me.pnlLandlord.Controls.Add(Me.lblSortcode)
        Me.pnlLandlord.Controls.Add(Me.lblBankName)
        Me.pnlLandlord.Controls.Add(Me.lblCommsTitle)
        Me.pnlLandlord.Controls.Add(Me.txtCommsTitle)
        Me.pnlLandlord.Controls.Add(Me.lblSurname2)
        Me.pnlLandlord.Controls.Add(Me.lblTitleForename2)
        Me.pnlLandlord.Controls.Add(Me.txtSurname2)
        Me.pnlLandlord.Controls.Add(Me.txtForename2)
        Me.pnlLandlord.Controls.Add(Me.txtTitle2)
        Me.pnlLandlord.Controls.Add(Me.lblStar4)
        Me.pnlLandlord.Controls.Add(Me.lblStar3)
        Me.pnlLandlord.Controls.Add(Me.lblSource)
        Me.pnlLandlord.Controls.Add(Me.cmbSource)
        Me.pnlLandlord.Controls.Add(Me.rtbLandlordNotes)
        Me.pnlLandlord.Controls.Add(Me.lblLandlordNotes)
        Me.pnlLandlord.Controls.Add(Me.txtEMailAddress)
        Me.pnlLandlord.Controls.Add(Me.lblEMailAddress)
        Me.pnlLandlord.Controls.Add(Me.txtMobileNo)
        Me.pnlLandlord.Controls.Add(Me.lblMobileNo)
        Me.pnlLandlord.Controls.Add(Me.txtPhoneNo)
        Me.pnlLandlord.Controls.Add(Me.lblPhoneNo)
        Me.pnlLandlord.Controls.Add(Me.txtPostCode)
        Me.pnlLandlord.Controls.Add(Me.txtTown)
        Me.pnlLandlord.Controls.Add(Me.txtAddress3)
        Me.pnlLandlord.Controls.Add(Me.txtAddress2)
        Me.pnlLandlord.Controls.Add(Me.txtAddress1)
        Me.pnlLandlord.Controls.Add(Me.lblPostCode)
        Me.pnlLandlord.Controls.Add(Me.lblTown)
        Me.pnlLandlord.Controls.Add(Me.lblAddress3)
        Me.pnlLandlord.Controls.Add(Me.lblAddress2)
        Me.pnlLandlord.Controls.Add(Me.lblAddress1)
        Me.pnlLandlord.Controls.Add(Me.lblSurname1)
        Me.pnlLandlord.Controls.Add(Me.lblForenameTitle1)
        Me.pnlLandlord.Controls.Add(Me.txtSurname1)
        Me.pnlLandlord.Controls.Add(Me.txtForename1)
        Me.pnlLandlord.Controls.Add(Me.txtTitle1)
        Me.pnlLandlord.Enabled = False
        Me.pnlLandlord.Location = New System.Drawing.Point(405, 103)
        Me.pnlLandlord.Name = "pnlLandlord"
        Me.pnlLandlord.Size = New System.Drawing.Size(847, 494)
        Me.pnlLandlord.TabIndex = 337
        '
        'lblLandlordName
        '
        Me.lblLandlordName.AutoSize = True
        Me.lblLandlordName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLandlordName.Location = New System.Drawing.Point(7, 73)
        Me.lblLandlordName.Name = "lblLandlordName"
        Me.lblLandlordName.Size = New System.Drawing.Size(118, 16)
        Me.lblLandlordName.TabIndex = 559
        Me.lblLandlordName.Text = "Landlord Name:"
        '
        'txtLandlordName
        '
        Me.txtLandlordName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLandlordName.Location = New System.Drawing.Point(132, 70)
        Me.txtLandlordName.MaxLength = 30
        Me.txtLandlordName.Name = "txtLandlordName"
        Me.txtLandlordName.Size = New System.Drawing.Size(703, 22)
        Me.txtLandlordName.TabIndex = 558
        '
        'lblPlaceofBirth2
        '
        Me.lblPlaceofBirth2.AutoSize = True
        Me.lblPlaceofBirth2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlaceofBirth2.Location = New System.Drawing.Point(466, 126)
        Me.lblPlaceofBirth2.Name = "lblPlaceofBirth2"
        Me.lblPlaceofBirth2.Size = New System.Drawing.Size(104, 16)
        Me.lblPlaceofBirth2.TabIndex = 557
        Me.lblPlaceofBirth2.Text = "Place of Birth:"
        '
        'txtPlaceofBirth2
        '
        Me.txtPlaceofBirth2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPlaceofBirth2.Location = New System.Drawing.Point(589, 123)
        Me.txtPlaceofBirth2.MaxLength = 30
        Me.txtPlaceofBirth2.Name = "txtPlaceofBirth2"
        Me.txtPlaceofBirth2.Size = New System.Drawing.Size(246, 22)
        Me.txtPlaceofBirth2.TabIndex = 556
        '
        'lblPlaceofBirth1
        '
        Me.lblPlaceofBirth1.AutoSize = True
        Me.lblPlaceofBirth1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlaceofBirth1.Location = New System.Drawing.Point(7, 129)
        Me.lblPlaceofBirth1.Name = "lblPlaceofBirth1"
        Me.lblPlaceofBirth1.Size = New System.Drawing.Size(104, 16)
        Me.lblPlaceofBirth1.TabIndex = 555
        Me.lblPlaceofBirth1.Text = "Place of Birth:"
        '
        'txtPlaceofBirth1
        '
        Me.txtPlaceofBirth1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPlaceofBirth1.Location = New System.Drawing.Point(132, 123)
        Me.txtPlaceofBirth1.MaxLength = 30
        Me.txtPlaceofBirth1.Name = "txtPlaceofBirth1"
        Me.txtPlaceofBirth1.Size = New System.Drawing.Size(246, 22)
        Me.txtPlaceofBirth1.TabIndex = 554
        '
        'lblDOB2
        '
        Me.lblDOB2.AutoSize = True
        Me.lblDOB2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDOB2.Location = New System.Drawing.Point(466, 101)
        Me.lblDOB2.Name = "lblDOB2"
        Me.lblDOB2.Size = New System.Drawing.Size(44, 16)
        Me.lblDOB2.TabIndex = 553
        Me.lblDOB2.Text = "DOB:"
        '
        'dtpDOB2
        '
        Me.dtpDOB2.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.dtpDOB2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.dtpDOB2.Location = New System.Drawing.Point(589, 96)
        Me.dtpDOB2.Name = "dtpDOB2"
        Me.dtpDOB2.Size = New System.Drawing.Size(147, 22)
        Me.dtpDOB2.TabIndex = 552
        '
        'lblDOB1
        '
        Me.lblDOB1.AutoSize = True
        Me.lblDOB1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDOB1.Location = New System.Drawing.Point(7, 101)
        Me.lblDOB1.Name = "lblDOB1"
        Me.lblDOB1.Size = New System.Drawing.Size(44, 16)
        Me.lblDOB1.TabIndex = 551
        Me.lblDOB1.Text = "DOB:"
        '
        'dtpDOB1
        '
        Me.dtpDOB1.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.dtpDOB1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.dtpDOB1.Location = New System.Drawing.Point(132, 96)
        Me.dtpDOB1.Name = "dtpDOB1"
        Me.dtpDOB1.Size = New System.Drawing.Size(147, 22)
        Me.dtpDOB1.TabIndex = 550
        '
        'lblDateJoined
        '
        Me.lblDateJoined.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateJoined.Location = New System.Drawing.Point(413, 377)
        Me.lblDateJoined.Name = "lblDateJoined"
        Me.lblDateJoined.Size = New System.Drawing.Size(96, 18)
        Me.lblDateJoined.TabIndex = 549
        Me.lblDateJoined.Text = "Date Joined:"
        Me.lblDateJoined.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtpDateJoined
        '
        Me.dtpDateJoined.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.dtpDateJoined.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.dtpDateJoined.Location = New System.Drawing.Point(513, 377)
        Me.dtpDateJoined.Name = "dtpDateJoined"
        Me.dtpDateJoined.Size = New System.Drawing.Size(147, 22)
        Me.dtpDateJoined.TabIndex = 548
        '
        'txtAccountNo
        '
        Me.txtAccountNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAccountNo.Location = New System.Drawing.Point(728, 343)
        Me.txtAccountNo.MaxLength = 8
        Me.txtAccountNo.Name = "txtAccountNo"
        Me.txtAccountNo.Size = New System.Drawing.Size(107, 22)
        Me.txtAccountNo.TabIndex = 381
        '
        'lblAccountNo
        '
        Me.lblAccountNo.AutoSize = True
        Me.lblAccountNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAccountNo.Location = New System.Drawing.Point(631, 348)
        Me.lblAccountNo.Name = "lblAccountNo"
        Me.lblAccountNo.Size = New System.Drawing.Size(91, 16)
        Me.lblAccountNo.TabIndex = 382
        Me.lblAccountNo.Text = "Account No:"
        '
        'txtSortcode
        '
        Me.txtSortcode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSortcode.Location = New System.Drawing.Point(505, 345)
        Me.txtSortcode.MaxLength = 8
        Me.txtSortcode.Name = "txtSortcode"
        Me.txtSortcode.Size = New System.Drawing.Size(107, 22)
        Me.txtSortcode.TabIndex = 378
        '
        'txtBankName
        '
        Me.txtBankName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBankName.Location = New System.Drawing.Point(132, 346)
        Me.txtBankName.MaxLength = 45
        Me.txtBankName.Name = "txtBankName"
        Me.txtBankName.Size = New System.Drawing.Size(267, 22)
        Me.txtBankName.TabIndex = 377
        '
        'lblSortcode
        '
        Me.lblSortcode.AutoSize = True
        Me.lblSortcode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSortcode.Location = New System.Drawing.Point(421, 348)
        Me.lblSortcode.Name = "lblSortcode"
        Me.lblSortcode.Size = New System.Drawing.Size(81, 16)
        Me.lblSortcode.TabIndex = 379
        Me.lblSortcode.Text = "Sort Code:"
        '
        'lblBankName
        '
        Me.lblBankName.AutoSize = True
        Me.lblBankName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBankName.Location = New System.Drawing.Point(7, 349)
        Me.lblBankName.Name = "lblBankName"
        Me.lblBankName.Size = New System.Drawing.Size(92, 16)
        Me.lblBankName.TabIndex = 380
        Me.lblBankName.Text = "Bank Name:"
        '
        'lblCommsTitle
        '
        Me.lblCommsTitle.AutoSize = True
        Me.lblCommsTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCommsTitle.Location = New System.Drawing.Point(7, 154)
        Me.lblCommsTitle.Name = "lblCommsTitle"
        Me.lblCommsTitle.Size = New System.Drawing.Size(102, 16)
        Me.lblCommsTitle.TabIndex = 376
        Me.lblCommsTitle.Text = "Comms. Title:"
        '
        'txtCommsTitle
        '
        Me.txtCommsTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCommsTitle.Location = New System.Drawing.Point(132, 151)
        Me.txtCommsTitle.MaxLength = 30
        Me.txtCommsTitle.Name = "txtCommsTitle"
        Me.txtCommsTitle.Size = New System.Drawing.Size(346, 22)
        Me.txtCommsTitle.TabIndex = 375
        '
        'lblSurname2
        '
        Me.lblSurname2.AutoSize = True
        Me.lblSurname2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSurname2.Location = New System.Drawing.Point(466, 45)
        Me.lblSurname2.Name = "lblSurname2"
        Me.lblSurname2.Size = New System.Drawing.Size(73, 16)
        Me.lblSurname2.TabIndex = 374
        Me.lblSurname2.Text = "Surname:"
        '
        'lblTitleForename2
        '
        Me.lblTitleForename2.AutoSize = True
        Me.lblTitleForename2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitleForename2.Location = New System.Drawing.Point(466, 17)
        Me.lblTitleForename2.Name = "lblTitleForename2"
        Me.lblTitleForename2.Size = New System.Drawing.Size(118, 16)
        Me.lblTitleForename2.TabIndex = 373
        Me.lblTitleForename2.Text = "Title/Forename:"
        '
        'txtSurname2
        '
        Me.txtSurname2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSurname2.Location = New System.Drawing.Point(589, 42)
        Me.txtSurname2.MaxLength = 30
        Me.txtSurname2.Name = "txtSurname2"
        Me.txtSurname2.Size = New System.Drawing.Size(246, 22)
        Me.txtSurname2.TabIndex = 372
        '
        'txtForename2
        '
        Me.txtForename2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtForename2.Location = New System.Drawing.Point(648, 14)
        Me.txtForename2.MaxLength = 30
        Me.txtForename2.Name = "txtForename2"
        Me.txtForename2.Size = New System.Drawing.Size(187, 22)
        Me.txtForename2.TabIndex = 371
        '
        'txtTitle2
        '
        Me.txtTitle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle2.Location = New System.Drawing.Point(589, 14)
        Me.txtTitle2.MaxLength = 8
        Me.txtTitle2.Name = "txtTitle2"
        Me.txtTitle2.Size = New System.Drawing.Size(53, 22)
        Me.txtTitle2.TabIndex = 370
        '
        'lblStar4
        '
        Me.lblStar4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStar4.ForeColor = System.Drawing.Color.Red
        Me.lblStar4.Location = New System.Drawing.Point(384, 45)
        Me.lblStar4.Name = "lblStar4"
        Me.lblStar4.Size = New System.Drawing.Size(15, 16)
        Me.lblStar4.TabIndex = 366
        Me.lblStar4.Text = "*"
        Me.lblStar4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblStar3
        '
        Me.lblStar3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStar3.ForeColor = System.Drawing.Color.Red
        Me.lblStar3.Location = New System.Drawing.Point(384, 17)
        Me.lblStar3.Name = "lblStar3"
        Me.lblStar3.Size = New System.Drawing.Size(15, 16)
        Me.lblStar3.TabIndex = 365
        Me.lblStar3.Text = "*"
        Me.lblStar3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSource
        '
        Me.lblSource.AutoSize = True
        Me.lblSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSource.Location = New System.Drawing.Point(7, 377)
        Me.lblSource.Name = "lblSource"
        Me.lblSource.Size = New System.Drawing.Size(61, 16)
        Me.lblSource.TabIndex = 360
        Me.lblSource.Text = "Source:"
        '
        'cmbSource
        '
        Me.cmbSource.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbSource.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbSource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmbSource.FormattingEnabled = True
        Me.cmbSource.Items.AddRange(New Object() {"1" & Global.Microsoft.VisualBasic.ChrW(9) & "Live"})
        Me.cmbSource.Location = New System.Drawing.Point(132, 374)
        Me.cmbSource.Name = "cmbSource"
        Me.cmbSource.Size = New System.Drawing.Size(267, 24)
        Me.cmbSource.TabIndex = 354
        Me.cmbSource.Tag = "Landlord_LoadSource"
        '
        'rtbLandlordNotes
        '
        Me.rtbLandlordNotes.Location = New System.Drawing.Point(12, 424)
        Me.rtbLandlordNotes.Name = "rtbLandlordNotes"
        Me.rtbLandlordNotes.Size = New System.Drawing.Size(823, 60)
        Me.rtbLandlordNotes.TabIndex = 355
        Me.rtbLandlordNotes.Text = ""
        '
        'lblLandlordNotes
        '
        Me.lblLandlordNotes.AutoSize = True
        Me.lblLandlordNotes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLandlordNotes.Location = New System.Drawing.Point(7, 405)
        Me.lblLandlordNotes.Name = "lblLandlordNotes"
        Me.lblLandlordNotes.Size = New System.Drawing.Size(114, 16)
        Me.lblLandlordNotes.TabIndex = 359
        Me.lblLandlordNotes.Text = "Landlord Notes"
        '
        'txtEMailAddress
        '
        Me.txtEMailAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEMailAddress.Location = New System.Drawing.Point(132, 318)
        Me.txtEMailAddress.MaxLength = 60
        Me.txtEMailAddress.Name = "txtEMailAddress"
        Me.txtEMailAddress.Size = New System.Drawing.Size(346, 22)
        Me.txtEMailAddress.TabIndex = 353
        '
        'lblEMailAddress
        '
        Me.lblEMailAddress.AutoSize = True
        Me.lblEMailAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEMailAddress.Location = New System.Drawing.Point(7, 322)
        Me.lblEMailAddress.Name = "lblEMailAddress"
        Me.lblEMailAddress.Size = New System.Drawing.Size(113, 16)
        Me.lblEMailAddress.TabIndex = 358
        Me.lblEMailAddress.Text = "EMail Address:"
        '
        'txtMobileNo
        '
        Me.txtMobileNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMobileNo.Location = New System.Drawing.Point(354, 290)
        Me.txtMobileNo.MaxLength = 18
        Me.txtMobileNo.Name = "txtMobileNo"
        Me.txtMobileNo.Size = New System.Drawing.Size(124, 22)
        Me.txtMobileNo.TabIndex = 351
        '
        'lblMobileNo
        '
        Me.lblMobileNo.AutoSize = True
        Me.lblMobileNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMobileNo.Location = New System.Drawing.Point(265, 293)
        Me.lblMobileNo.Name = "lblMobileNo"
        Me.lblMobileNo.Size = New System.Drawing.Size(83, 16)
        Me.lblMobileNo.TabIndex = 357
        Me.lblMobileNo.Text = "Mobile No:"
        '
        'txtPhoneNo
        '
        Me.txtPhoneNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhoneNo.Location = New System.Drawing.Point(132, 290)
        Me.txtPhoneNo.MaxLength = 148
        Me.txtPhoneNo.Name = "txtPhoneNo"
        Me.txtPhoneNo.Size = New System.Drawing.Size(124, 22)
        Me.txtPhoneNo.TabIndex = 349
        '
        'lblPhoneNo
        '
        Me.lblPhoneNo.AutoSize = True
        Me.lblPhoneNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhoneNo.Location = New System.Drawing.Point(7, 293)
        Me.lblPhoneNo.Name = "lblPhoneNo"
        Me.lblPhoneNo.Size = New System.Drawing.Size(80, 16)
        Me.lblPhoneNo.TabIndex = 356
        Me.lblPhoneNo.Text = "Phone No:"
        '
        'txtPostCode
        '
        Me.txtPostCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostCode.Location = New System.Drawing.Point(371, 262)
        Me.txtPostCode.MaxLength = 8
        Me.txtPostCode.Name = "txtPostCode"
        Me.txtPostCode.Size = New System.Drawing.Size(107, 22)
        Me.txtPostCode.TabIndex = 348
        '
        'txtTown
        '
        Me.txtTown.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTown.Location = New System.Drawing.Point(132, 262)
        Me.txtTown.MaxLength = 45
        Me.txtTown.Name = "txtTown"
        Me.txtTown.Size = New System.Drawing.Size(143, 22)
        Me.txtTown.TabIndex = 346
        '
        'txtAddress3
        '
        Me.txtAddress3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddress3.Location = New System.Drawing.Point(132, 234)
        Me.txtAddress3.MaxLength = 45
        Me.txtAddress3.Name = "txtAddress3"
        Me.txtAddress3.Size = New System.Drawing.Size(346, 22)
        Me.txtAddress3.TabIndex = 344
        '
        'txtAddress2
        '
        Me.txtAddress2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddress2.Location = New System.Drawing.Point(132, 206)
        Me.txtAddress2.MaxLength = 45
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(346, 22)
        Me.txtAddress2.TabIndex = 342
        '
        'txtAddress1
        '
        Me.txtAddress1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddress1.Location = New System.Drawing.Point(132, 178)
        Me.txtAddress1.MaxLength = 45
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(346, 22)
        Me.txtAddress1.TabIndex = 340
        '
        'lblPostCode
        '
        Me.lblPostCode.AutoSize = True
        Me.lblPostCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPostCode.Location = New System.Drawing.Point(287, 265)
        Me.lblPostCode.Name = "lblPostCode"
        Me.lblPostCode.Size = New System.Drawing.Size(78, 16)
        Me.lblPostCode.TabIndex = 350
        Me.lblPostCode.Text = "Postcode:"
        '
        'lblTown
        '
        Me.lblTown.AutoSize = True
        Me.lblTown.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTown.Location = New System.Drawing.Point(7, 239)
        Me.lblTown.Name = "lblTown"
        Me.lblTown.Size = New System.Drawing.Size(49, 16)
        Me.lblTown.TabIndex = 352
        Me.lblTown.Text = "Town:"
        '
        'lblAddress3
        '
        Me.lblAddress3.AutoSize = True
        Me.lblAddress3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress3.Location = New System.Drawing.Point(7, 237)
        Me.lblAddress3.Name = "lblAddress3"
        Me.lblAddress3.Size = New System.Drawing.Size(115, 16)
        Me.lblAddress3.TabIndex = 347
        Me.lblAddress3.Text = "Address Line 3:"
        '
        'lblAddress2
        '
        Me.lblAddress2.AutoSize = True
        Me.lblAddress2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress2.Location = New System.Drawing.Point(7, 209)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(115, 16)
        Me.lblAddress2.TabIndex = 345
        Me.lblAddress2.Text = "Address Line 2:"
        '
        'lblAddress1
        '
        Me.lblAddress1.AutoSize = True
        Me.lblAddress1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress1.Location = New System.Drawing.Point(7, 181)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New System.Drawing.Size(115, 16)
        Me.lblAddress1.TabIndex = 343
        Me.lblAddress1.Text = "Address Line 1:"
        '
        'lblSurname1
        '
        Me.lblSurname1.AutoSize = True
        Me.lblSurname1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSurname1.Location = New System.Drawing.Point(7, 45)
        Me.lblSurname1.Name = "lblSurname1"
        Me.lblSurname1.Size = New System.Drawing.Size(73, 16)
        Me.lblSurname1.TabIndex = 341
        Me.lblSurname1.Text = "Surname:"
        '
        'lblForenameTitle1
        '
        Me.lblForenameTitle1.AutoSize = True
        Me.lblForenameTitle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblForenameTitle1.Location = New System.Drawing.Point(7, 17)
        Me.lblForenameTitle1.Name = "lblForenameTitle1"
        Me.lblForenameTitle1.Size = New System.Drawing.Size(118, 16)
        Me.lblForenameTitle1.TabIndex = 339
        Me.lblForenameTitle1.Text = "Title/Forename:"
        '
        'txtSurname1
        '
        Me.txtSurname1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSurname1.Location = New System.Drawing.Point(132, 42)
        Me.txtSurname1.MaxLength = 30
        Me.txtSurname1.Name = "txtSurname1"
        Me.txtSurname1.Size = New System.Drawing.Size(246, 22)
        Me.txtSurname1.TabIndex = 337
        '
        'txtForename1
        '
        Me.txtForename1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtForename1.Location = New System.Drawing.Point(191, 14)
        Me.txtForename1.MaxLength = 30
        Me.txtForename1.Name = "txtForename1"
        Me.txtForename1.Size = New System.Drawing.Size(187, 22)
        Me.txtForename1.TabIndex = 336
        '
        'txtTitle1
        '
        Me.txtTitle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle1.Location = New System.Drawing.Point(132, 14)
        Me.txtTitle1.MaxLength = 8
        Me.txtTitle1.Name = "txtTitle1"
        Me.txtTitle1.Size = New System.Drawing.Size(53, 22)
        Me.txtTitle1.TabIndex = 335
        '
        'tsbClient
        '
        Me.tsbClient.ImageScalingSize = New System.Drawing.Size(40, 40)
        Me.tsbClient.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbNewLandlord, Me.tsbSaveLandlord, Me.tsbRemoveLandlord, Me.ToolStripSeparator1, Me.tsbNewProperty, Me.tsbEditProperty, Me.ToolStripSeparator2, Me.tsbClose})
        Me.tsbClient.Location = New System.Drawing.Point(0, 24)
        Me.tsbClient.Name = "tsbClient"
        Me.tsbClient.Size = New System.Drawing.Size(1264, 47)
        Me.tsbClient.TabIndex = 338
        Me.tsbClient.Text = "tstClient"
        '
        'tsbNewLandlord
        '
        Me.tsbNewLandlord.Image = CType(resources.GetObject("tsbNewLandlord.Image"), System.Drawing.Image)
        Me.tsbNewLandlord.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbNewLandlord.Name = "tsbNewLandlord"
        Me.tsbNewLandlord.Size = New System.Drawing.Size(123, 44)
        Me.tsbNewLandlord.Text = "Add Landlord"
        Me.tsbNewLandlord.ToolTipText = "New Landlord"
        '
        'tsbSaveLandlord
        '
        Me.tsbSaveLandlord.Image = CType(resources.GetObject("tsbSaveLandlord.Image"), System.Drawing.Image)
        Me.tsbSaveLandlord.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbSaveLandlord.Name = "tsbSaveLandlord"
        Me.tsbSaveLandlord.Size = New System.Drawing.Size(125, 44)
        Me.tsbSaveLandlord.Text = "Save Landlord"
        Me.tsbSaveLandlord.ToolTipText = "Save Landlord"
        '
        'tsbRemoveLandlord
        '
        Me.tsbRemoveLandlord.Image = CType(resources.GetObject("tsbRemoveLandlord.Image"), System.Drawing.Image)
        Me.tsbRemoveLandlord.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbRemoveLandlord.Name = "tsbRemoveLandlord"
        Me.tsbRemoveLandlord.Size = New System.Drawing.Size(144, 44)
        Me.tsbRemoveLandlord.Text = "Remove Landlord"
        Me.tsbRemoveLandlord.ToolTipText = "Remove Landlord"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 47)
        '
        'tsbNewProperty
        '
        Me.tsbNewProperty.Enabled = False
        Me.tsbNewProperty.Image = CType(resources.GetObject("tsbNewProperty.Image"), System.Drawing.Image)
        Me.tsbNewProperty.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbNewProperty.Name = "tsbNewProperty"
        Me.tsbNewProperty.Size = New System.Drawing.Size(121, 44)
        Me.tsbNewProperty.Text = "Add Property"
        Me.tsbNewProperty.ToolTipText = "New Property"
        '
        'tsbEditProperty
        '
        Me.tsbEditProperty.Enabled = False
        Me.tsbEditProperty.Image = CType(resources.GetObject("tsbEditProperty.Image"), System.Drawing.Image)
        Me.tsbEditProperty.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbEditProperty.Name = "tsbEditProperty"
        Me.tsbEditProperty.Size = New System.Drawing.Size(119, 44)
        Me.tsbEditProperty.Text = "Edit Property"
        Me.tsbEditProperty.ToolTipText = "Edit Property"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 47)
        '
        'tsbClose
        '
        Me.tsbClose.Image = CType(resources.GetObject("tsbClose.Image"), System.Drawing.Image)
        Me.tsbClose.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbClose.Name = "tsbClose"
        Me.tsbClose.Size = New System.Drawing.Size(80, 44)
        Me.tsbClose.Text = "Close"
        '
        'ProgrammeListBindingSource
        '
        Me.ProgrammeListBindingSource.DataMember = "ProgrammeList"
        Me.ProgrammeListBindingSource.Filter = "ClientID = 0"
        '
        'lblFilter
        '
        Me.lblFilter.AutoSize = True
        Me.lblFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFilter.Location = New System.Drawing.Point(160, 80)
        Me.lblFilter.Name = "lblFilter"
        Me.lblFilter.Size = New System.Drawing.Size(47, 16)
        Me.lblFilter.TabIndex = 342
        Me.lblFilter.Text = "Filter:"
        '
        'tvwLandlordList
        '
        Me.tvwLandlordList.Location = New System.Drawing.Point(15, 106)
        Me.tvwLandlordList.Name = "tvwLandlordList"
        Me.tvwLandlordList.Size = New System.Drawing.Size(373, 604)
        Me.tvwLandlordList.TabIndex = 343
        '
        'frmLandlord
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 761)
        Me.Controls.Add(Me.tvwLandlordList)
        Me.Controls.Add(Me.lblFilter)
        Me.Controls.Add(Me.txtFilter)
        Me.Controls.Add(Me.tsbClient)
        Me.Controls.Add(Me.pnlLandlord)
        Me.Controls.Add(Me.lblClients)
        Me.Controls.Add(Me.lblMessagePanel)
        Me.Controls.Add(Me.lblLandlordIDValue)
        Me.Controls.Add(Me.lblLandlordID)
        Me.Controls.Add(Me.MenuStrip1)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLandlord"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Landlord Management"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.pnlLandlord.ResumeLayout(False)
        Me.pnlLandlord.PerformLayout()
        Me.tsbClient.ResumeLayout(False)
        Me.tsbClient.PerformLayout()
        CType(Me.ProgrammeListBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileClose As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLandlord As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNewLandlord As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSaveLandlord As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRemoveLandlord As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuJob As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProgrammeListBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents lblMessagePanel As System.Windows.Forms.Label
    Friend WithEvents lblLandlordIDValue As System.Windows.Forms.Label
    Friend WithEvents lblLandlordID As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents lblClients As System.Windows.Forms.Label
    Friend WithEvents txtFilter As System.Windows.Forms.TextBox
    Friend WithEvents pnlLandlord As System.Windows.Forms.Panel
    Friend WithEvents lblSource As System.Windows.Forms.Label
    Friend WithEvents cmbSource As System.Windows.Forms.ComboBox
    Friend WithEvents rtbLandlordNotes As System.Windows.Forms.RichTextBox
    Friend WithEvents lblLandlordNotes As System.Windows.Forms.Label
    Friend WithEvents txtEMailAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblEMailAddress As System.Windows.Forms.Label
    Friend WithEvents txtMobileNo As System.Windows.Forms.TextBox
    Friend WithEvents lblMobileNo As System.Windows.Forms.Label
    Friend WithEvents txtPhoneNo As System.Windows.Forms.TextBox
    Friend WithEvents lblPhoneNo As System.Windows.Forms.Label
    Friend WithEvents txtPostCode As System.Windows.Forms.TextBox
    Friend WithEvents txtTown As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress3 As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress2 As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress1 As System.Windows.Forms.TextBox
    Friend WithEvents lblPostCode As System.Windows.Forms.Label
    Friend WithEvents lblTown As System.Windows.Forms.Label
    Friend WithEvents lblAddress3 As System.Windows.Forms.Label
    Friend WithEvents lblAddress2 As System.Windows.Forms.Label
    Friend WithEvents lblAddress1 As System.Windows.Forms.Label
    Friend WithEvents lblSurname1 As System.Windows.Forms.Label
    Friend WithEvents lblForenameTitle1 As System.Windows.Forms.Label
    Friend WithEvents txtSurname1 As System.Windows.Forms.TextBox
    Friend WithEvents txtForename1 As System.Windows.Forms.TextBox
    Friend WithEvents tsbClient As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbNewLandlord As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbClose As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbSaveLandlord As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbRemoveLandlord As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbEditProperty As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbNewProperty As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents txtTitle1 As System.Windows.Forms.TextBox
    Friend WithEvents mnuNewProperty As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditProperty As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblStar4 As System.Windows.Forms.Label
    Friend WithEvents lblStar3 As System.Windows.Forms.Label
    Friend WithEvents lblFilter As System.Windows.Forms.Label
    Friend WithEvents lblSurname2 As System.Windows.Forms.Label
    Friend WithEvents lblTitleForename2 As System.Windows.Forms.Label
    Friend WithEvents txtSurname2 As System.Windows.Forms.TextBox
    Friend WithEvents txtForename2 As System.Windows.Forms.TextBox
    Friend WithEvents txtTitle2 As System.Windows.Forms.TextBox
    Friend WithEvents tvwLandlordList As System.Windows.Forms.TreeView
    Friend WithEvents lblPlaceofBirth2 As System.Windows.Forms.Label
    Friend WithEvents txtPlaceofBirth2 As System.Windows.Forms.TextBox
    Friend WithEvents lblPlaceofBirth1 As System.Windows.Forms.Label
    Friend WithEvents txtPlaceofBirth1 As System.Windows.Forms.TextBox
    Friend WithEvents lblDOB2 As System.Windows.Forms.Label
    Friend WithEvents dtpDOB2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDOB1 As System.Windows.Forms.Label
    Friend WithEvents dtpDOB1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateJoined As System.Windows.Forms.Label
    Friend WithEvents dtpDateJoined As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtAccountNo As System.Windows.Forms.TextBox
    Friend WithEvents lblAccountNo As System.Windows.Forms.Label
    Friend WithEvents txtSortcode As System.Windows.Forms.TextBox
    Friend WithEvents txtBankName As System.Windows.Forms.TextBox
    Friend WithEvents lblSortcode As System.Windows.Forms.Label
    Friend WithEvents lblBankName As System.Windows.Forms.Label
    Friend WithEvents lblCommsTitle As System.Windows.Forms.Label
    Friend WithEvents txtCommsTitle As System.Windows.Forms.TextBox
    Friend WithEvents lblLandlordName As System.Windows.Forms.Label
    Friend WithEvents txtLandlordName As System.Windows.Forms.TextBox
End Class
