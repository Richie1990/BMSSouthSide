<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStatisticsSortmyPC
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
        Me.tmrStatistics = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblInProgressHoursMonth = New System.Windows.Forms.Label()
        Me.lblInProgressNoMonth = New System.Windows.Forms.Label()
        Me.lblInProgress = New System.Windows.Forms.Label()
        Me.lblUnallocatedNoMonth = New System.Windows.Forms.Label()
        Me.lblUnallocated = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.nudDay = New System.Windows.Forms.NumericUpDown()
        Me.nudWeek = New System.Windows.Forms.NumericUpDown()
        Me.nudMonth = New System.Windows.Forms.NumericUpDown()
        Me.lblDay = New System.Windows.Forms.Label()
        Me.lblContractsHoursDay = New System.Windows.Forms.Label()
        Me.lblContractsNoDay = New System.Windows.Forms.Label()
        Me.lblCompletedHoursDay = New System.Windows.Forms.Label()
        Me.lblCompletedNoDay = New System.Windows.Forms.Label()
        Me.lblContractsHoursWeek = New System.Windows.Forms.Label()
        Me.lblContractsNoWeek = New System.Windows.Forms.Label()
        Me.lblCompletedHoursWeek = New System.Windows.Forms.Label()
        Me.lblCompletedNoWeek = New System.Windows.Forms.Label()
        Me.lblContractsHoursMonth = New System.Windows.Forms.Label()
        Me.lblContractsNoMonth = New System.Windows.Forms.Label()
        Me.lblCompletedHoursMonth = New System.Windows.Forms.Label()
        Me.lblCompletedNoMonth = New System.Windows.Forms.Label()
        Me.lblWeek = New System.Windows.Forms.Label()
        Me.lblContracts = New System.Windows.Forms.Label()
        Me.lblMonth = New System.Windows.Forms.Label()
        Me.lblCompleted = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.lblContractsEngineer = New System.Windows.Forms.Label()
        Me.lblCompletedEngineer = New System.Windows.Forms.Label()
        Me.lblJobsinProgressEngineer = New System.Windows.Forms.Label()
        Me.lblContractsHoursDay2 = New System.Windows.Forms.Label()
        Me.lblContractsNoDay2 = New System.Windows.Forms.Label()
        Me.lblCompletedHoursDay2 = New System.Windows.Forms.Label()
        Me.lblCompletedNoDay2 = New System.Windows.Forms.Label()
        Me.lblInProgressHoursDay2 = New System.Windows.Forms.Label()
        Me.lblInProgressNoDay2 = New System.Windows.Forms.Label()
        Me.cmbEngineer2 = New System.Windows.Forms.ComboBox()
        Me.lblEngineer2 = New System.Windows.Forms.Label()
        Me.lblContractsHoursDay1 = New System.Windows.Forms.Label()
        Me.lblContractsNoDay1 = New System.Windows.Forms.Label()
        Me.lblCompletedHoursDay1 = New System.Windows.Forms.Label()
        Me.lblCompletedNoDay1 = New System.Windows.Forms.Label()
        Me.lblInProgressHoursDay1 = New System.Windows.Forms.Label()
        Me.lblInProgressNoDay1 = New System.Windows.Forms.Label()
        Me.cmbEngineer1 = New System.Windows.Forms.ComboBox()
        Me.lblEngineer1 = New System.Windows.Forms.Label()
        Me.lblContractsHoursMonth1 = New System.Windows.Forms.Label()
        Me.lblContractsNoMonth1 = New System.Windows.Forms.Label()
        Me.lblCompletedHoursMonth1 = New System.Windows.Forms.Label()
        Me.lblCompletedNoMonth1 = New System.Windows.Forms.Label()
        Me.lblContractsHoursMonth2 = New System.Windows.Forms.Label()
        Me.lblContractsNoMonth2 = New System.Windows.Forms.Label()
        Me.lblCompletedHoursMonth2 = New System.Windows.Forms.Label()
        Me.lblCompletedNoMonth2 = New System.Windows.Forms.Label()
        Me.lblContractsTravelHoursMonth1 = New System.Windows.Forms.Label()
        Me.lblContractsTravelHoursDay1 = New System.Windows.Forms.Label()
        Me.lblContractsTravelHoursMonth2 = New System.Windows.Forms.Label()
        Me.lblContractsTravelHoursDay2 = New System.Windows.Forms.Label()
        Me.lblContractsTravelHoursMonth = New System.Windows.Forms.Label()
        Me.lblContractsTravelHoursWeek = New System.Windows.Forms.Label()
        Me.lblContractsTravelHoursDay = New System.Windows.Forms.Label()
        Me.lblMonth1 = New System.Windows.Forms.Label()
        Me.lblMonth2 = New System.Windows.Forms.Label()
        Me.lblDay1 = New System.Windows.Forms.Label()
        Me.lblDay2 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.nudDay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudWeek, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMonth, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'tmrStatistics
        '
        Me.tmrStatistics.Enabled = True
        Me.tmrStatistics.Interval = 3000
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.lblInProgressHoursMonth)
        Me.Panel1.Controls.Add(Me.lblInProgressNoMonth)
        Me.Panel1.Controls.Add(Me.lblInProgress)
        Me.Panel1.Controls.Add(Me.lblUnallocatedNoMonth)
        Me.Panel1.Controls.Add(Me.lblUnallocated)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(312, 61)
        Me.Panel1.TabIndex = 461
        '
        'lblInProgressHoursMonth
        '
        Me.lblInProgressHoursMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInProgressHoursMonth.Location = New System.Drawing.Point(234, 34)
        Me.lblInProgressHoursMonth.Name = "lblInProgressHoursMonth"
        Me.lblInProgressHoursMonth.Size = New System.Drawing.Size(40, 18)
        Me.lblInProgressHoursMonth.TabIndex = 424
        Me.lblInProgressHoursMonth.Text = "00:00"
        '
        'lblInProgressNoMonth
        '
        Me.lblInProgressNoMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInProgressNoMonth.Location = New System.Drawing.Point(194, 34)
        Me.lblInProgressNoMonth.Name = "lblInProgressNoMonth"
        Me.lblInProgressNoMonth.Size = New System.Drawing.Size(30, 18)
        Me.lblInProgressNoMonth.TabIndex = 423
        Me.lblInProgressNoMonth.Text = "000"
        '
        'lblInProgress
        '
        Me.lblInProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInProgress.Location = New System.Drawing.Point(167, 9)
        Me.lblInProgress.Name = "lblInProgress"
        Me.lblInProgress.Size = New System.Drawing.Size(130, 21)
        Me.lblInProgress.TabIndex = 422
        Me.lblInProgress.Text = "Jobs in Progress"
        Me.lblInProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblUnallocatedNoMonth
        '
        Me.lblUnallocatedNoMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUnallocatedNoMonth.Location = New System.Drawing.Point(62, 34)
        Me.lblUnallocatedNoMonth.Name = "lblUnallocatedNoMonth"
        Me.lblUnallocatedNoMonth.Size = New System.Drawing.Size(30, 18)
        Me.lblUnallocatedNoMonth.TabIndex = 421
        Me.lblUnallocatedNoMonth.Text = "000"
        '
        'lblUnallocated
        '
        Me.lblUnallocated.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUnallocated.Location = New System.Drawing.Point(13, 9)
        Me.lblUnallocated.Name = "lblUnallocated"
        Me.lblUnallocated.Size = New System.Drawing.Size(130, 19)
        Me.lblUnallocated.TabIndex = 420
        Me.lblUnallocated.Text = "Unallocated Jobs"
        Me.lblUnallocated.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.lblContractsTravelHoursDay)
        Me.Panel2.Controls.Add(Me.lblContractsTravelHoursWeek)
        Me.Panel2.Controls.Add(Me.lblContractsTravelHoursMonth)
        Me.Panel2.Controls.Add(Me.nudDay)
        Me.Panel2.Controls.Add(Me.nudWeek)
        Me.Panel2.Controls.Add(Me.nudMonth)
        Me.Panel2.Controls.Add(Me.lblDay)
        Me.Panel2.Controls.Add(Me.lblContractsHoursDay)
        Me.Panel2.Controls.Add(Me.lblContractsNoDay)
        Me.Panel2.Controls.Add(Me.lblCompletedHoursDay)
        Me.Panel2.Controls.Add(Me.lblCompletedNoDay)
        Me.Panel2.Controls.Add(Me.lblContractsHoursWeek)
        Me.Panel2.Controls.Add(Me.lblContractsNoWeek)
        Me.Panel2.Controls.Add(Me.lblCompletedHoursWeek)
        Me.Panel2.Controls.Add(Me.lblCompletedNoWeek)
        Me.Panel2.Controls.Add(Me.lblContractsHoursMonth)
        Me.Panel2.Controls.Add(Me.lblContractsNoMonth)
        Me.Panel2.Controls.Add(Me.lblCompletedHoursMonth)
        Me.Panel2.Controls.Add(Me.lblCompletedNoMonth)
        Me.Panel2.Controls.Add(Me.lblWeek)
        Me.Panel2.Controls.Add(Me.lblContracts)
        Me.Panel2.Controls.Add(Me.lblMonth)
        Me.Panel2.Controls.Add(Me.lblCompleted)
        Me.Panel2.Location = New System.Drawing.Point(330, 12)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(427, 115)
        Me.Panel2.TabIndex = 465
        '
        'nudDay
        '
        Me.nudDay.Location = New System.Drawing.Point(69, 79)
        Me.nudDay.Maximum = New Decimal(New Integer() {31, 0, 0, 0})
        Me.nudDay.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudDay.Name = "nudDay"
        Me.nudDay.Size = New System.Drawing.Size(42, 20)
        Me.nudDay.TabIndex = 480
        Me.nudDay.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'nudWeek
        '
        Me.nudWeek.Location = New System.Drawing.Point(69, 57)
        Me.nudWeek.Maximum = New Decimal(New Integer() {53, 0, 0, 0})
        Me.nudWeek.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudWeek.Name = "nudWeek"
        Me.nudWeek.Size = New System.Drawing.Size(42, 20)
        Me.nudWeek.TabIndex = 479
        Me.nudWeek.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'nudMonth
        '
        Me.nudMonth.Location = New System.Drawing.Point(69, 34)
        Me.nudMonth.Maximum = New Decimal(New Integer() {12, 0, 0, 0})
        Me.nudMonth.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudMonth.Name = "nudMonth"
        Me.nudMonth.Size = New System.Drawing.Size(42, 20)
        Me.nudMonth.TabIndex = 478
        Me.nudMonth.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'lblDay
        '
        Me.lblDay.AutoSize = True
        Me.lblDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDay.Location = New System.Drawing.Point(17, 81)
        Me.lblDay.Name = "lblDay"
        Me.lblDay.Size = New System.Drawing.Size(33, 13)
        Me.lblDay.TabIndex = 477
        Me.lblDay.Text = "Day:"
        '
        'lblContractsHoursDay
        '
        Me.lblContractsHoursDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsHoursDay.Location = New System.Drawing.Point(315, 79)
        Me.lblContractsHoursDay.Name = "lblContractsHoursDay"
        Me.lblContractsHoursDay.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsHoursDay.TabIndex = 476
        Me.lblContractsHoursDay.Text = "00:00"
        '
        'lblContractsNoDay
        '
        Me.lblContractsNoDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsNoDay.Location = New System.Drawing.Point(275, 79)
        Me.lblContractsNoDay.Name = "lblContractsNoDay"
        Me.lblContractsNoDay.Size = New System.Drawing.Size(30, 13)
        Me.lblContractsNoDay.TabIndex = 475
        Me.lblContractsNoDay.Text = "000"
        '
        'lblCompletedHoursDay
        '
        Me.lblCompletedHoursDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedHoursDay.Location = New System.Drawing.Point(185, 79)
        Me.lblCompletedHoursDay.Name = "lblCompletedHoursDay"
        Me.lblCompletedHoursDay.Size = New System.Drawing.Size(40, 13)
        Me.lblCompletedHoursDay.TabIndex = 474
        Me.lblCompletedHoursDay.Text = "00:00"
        '
        'lblCompletedNoDay
        '
        Me.lblCompletedNoDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedNoDay.Location = New System.Drawing.Point(145, 79)
        Me.lblCompletedNoDay.Name = "lblCompletedNoDay"
        Me.lblCompletedNoDay.Size = New System.Drawing.Size(30, 13)
        Me.lblCompletedNoDay.TabIndex = 473
        Me.lblCompletedNoDay.Text = "000"
        '
        'lblContractsHoursWeek
        '
        Me.lblContractsHoursWeek.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsHoursWeek.Location = New System.Drawing.Point(315, 56)
        Me.lblContractsHoursWeek.Name = "lblContractsHoursWeek"
        Me.lblContractsHoursWeek.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsHoursWeek.TabIndex = 472
        Me.lblContractsHoursWeek.Text = "00:00"
        '
        'lblContractsNoWeek
        '
        Me.lblContractsNoWeek.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsNoWeek.Location = New System.Drawing.Point(275, 56)
        Me.lblContractsNoWeek.Name = "lblContractsNoWeek"
        Me.lblContractsNoWeek.Size = New System.Drawing.Size(30, 13)
        Me.lblContractsNoWeek.TabIndex = 471
        Me.lblContractsNoWeek.Text = "000"
        '
        'lblCompletedHoursWeek
        '
        Me.lblCompletedHoursWeek.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedHoursWeek.Location = New System.Drawing.Point(185, 56)
        Me.lblCompletedHoursWeek.Name = "lblCompletedHoursWeek"
        Me.lblCompletedHoursWeek.Size = New System.Drawing.Size(40, 13)
        Me.lblCompletedHoursWeek.TabIndex = 470
        Me.lblCompletedHoursWeek.Text = "00:00"
        '
        'lblCompletedNoWeek
        '
        Me.lblCompletedNoWeek.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedNoWeek.Location = New System.Drawing.Point(145, 56)
        Me.lblCompletedNoWeek.Name = "lblCompletedNoWeek"
        Me.lblCompletedNoWeek.Size = New System.Drawing.Size(30, 13)
        Me.lblCompletedNoWeek.TabIndex = 469
        Me.lblCompletedNoWeek.Text = "000"
        '
        'lblContractsHoursMonth
        '
        Me.lblContractsHoursMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsHoursMonth.Location = New System.Drawing.Point(315, 34)
        Me.lblContractsHoursMonth.Name = "lblContractsHoursMonth"
        Me.lblContractsHoursMonth.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsHoursMonth.TabIndex = 468
        Me.lblContractsHoursMonth.Text = "00:00"
        '
        'lblContractsNoMonth
        '
        Me.lblContractsNoMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsNoMonth.Location = New System.Drawing.Point(275, 34)
        Me.lblContractsNoMonth.Name = "lblContractsNoMonth"
        Me.lblContractsNoMonth.Size = New System.Drawing.Size(30, 13)
        Me.lblContractsNoMonth.TabIndex = 467
        Me.lblContractsNoMonth.Text = "000"
        '
        'lblCompletedHoursMonth
        '
        Me.lblCompletedHoursMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedHoursMonth.Location = New System.Drawing.Point(185, 34)
        Me.lblCompletedHoursMonth.Name = "lblCompletedHoursMonth"
        Me.lblCompletedHoursMonth.Size = New System.Drawing.Size(40, 13)
        Me.lblCompletedHoursMonth.TabIndex = 466
        Me.lblCompletedHoursMonth.Text = "00:00"
        '
        'lblCompletedNoMonth
        '
        Me.lblCompletedNoMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedNoMonth.Location = New System.Drawing.Point(145, 34)
        Me.lblCompletedNoMonth.Name = "lblCompletedNoMonth"
        Me.lblCompletedNoMonth.Size = New System.Drawing.Size(30, 13)
        Me.lblCompletedNoMonth.TabIndex = 465
        Me.lblCompletedNoMonth.Text = "000"
        '
        'lblWeek
        '
        Me.lblWeek.AutoSize = True
        Me.lblWeek.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWeek.Location = New System.Drawing.Point(17, 59)
        Me.lblWeek.Name = "lblWeek"
        Me.lblWeek.Size = New System.Drawing.Size(44, 13)
        Me.lblWeek.TabIndex = 464
        Me.lblWeek.Text = "Week:"
        '
        'lblContracts
        '
        Me.lblContracts.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContracts.Location = New System.Drawing.Point(245, 11)
        Me.lblContracts.Name = "lblContracts"
        Me.lblContracts.Size = New System.Drawing.Size(130, 13)
        Me.lblContracts.TabIndex = 463
        Me.lblContracts.Text = "Contracts Completed"
        Me.lblContracts.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMonth
        '
        Me.lblMonth.AutoSize = True
        Me.lblMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonth.Location = New System.Drawing.Point(17, 36)
        Me.lblMonth.Name = "lblMonth"
        Me.lblMonth.Size = New System.Drawing.Size(46, 13)
        Me.lblMonth.TabIndex = 462
        Me.lblMonth.Text = "Month:"
        '
        'lblCompleted
        '
        Me.lblCompleted.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompleted.Location = New System.Drawing.Point(115, 11)
        Me.lblCompleted.Name = "lblCompleted"
        Me.lblCompleted.Size = New System.Drawing.Size(130, 13)
        Me.lblCompleted.TabIndex = 461
        Me.lblCompleted.Text = "Jobs Completed"
        Me.lblCompleted.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.lblDay2)
        Me.Panel3.Controls.Add(Me.lblDay1)
        Me.Panel3.Controls.Add(Me.lblMonth2)
        Me.Panel3.Controls.Add(Me.lblMonth1)
        Me.Panel3.Controls.Add(Me.lblContractsTravelHoursDay2)
        Me.Panel3.Controls.Add(Me.lblContractsTravelHoursMonth2)
        Me.Panel3.Controls.Add(Me.lblContractsTravelHoursDay1)
        Me.Panel3.Controls.Add(Me.lblContractsTravelHoursMonth1)
        Me.Panel3.Controls.Add(Me.lblContractsHoursMonth2)
        Me.Panel3.Controls.Add(Me.lblContractsNoMonth2)
        Me.Panel3.Controls.Add(Me.lblCompletedHoursMonth2)
        Me.Panel3.Controls.Add(Me.lblCompletedNoMonth2)
        Me.Panel3.Controls.Add(Me.lblContractsHoursMonth1)
        Me.Panel3.Controls.Add(Me.lblContractsNoMonth1)
        Me.Panel3.Controls.Add(Me.lblCompletedHoursMonth1)
        Me.Panel3.Controls.Add(Me.lblCompletedNoMonth1)
        Me.Panel3.Controls.Add(Me.lblContractsEngineer)
        Me.Panel3.Controls.Add(Me.lblCompletedEngineer)
        Me.Panel3.Controls.Add(Me.lblJobsinProgressEngineer)
        Me.Panel3.Controls.Add(Me.lblContractsHoursDay2)
        Me.Panel3.Controls.Add(Me.lblContractsNoDay2)
        Me.Panel3.Controls.Add(Me.lblCompletedHoursDay2)
        Me.Panel3.Controls.Add(Me.lblCompletedNoDay2)
        Me.Panel3.Controls.Add(Me.lblInProgressHoursDay2)
        Me.Panel3.Controls.Add(Me.lblInProgressNoDay2)
        Me.Panel3.Controls.Add(Me.cmbEngineer2)
        Me.Panel3.Controls.Add(Me.lblEngineer2)
        Me.Panel3.Controls.Add(Me.lblContractsHoursDay1)
        Me.Panel3.Controls.Add(Me.lblContractsNoDay1)
        Me.Panel3.Controls.Add(Me.lblCompletedHoursDay1)
        Me.Panel3.Controls.Add(Me.lblCompletedNoDay1)
        Me.Panel3.Controls.Add(Me.lblInProgressHoursDay1)
        Me.Panel3.Controls.Add(Me.lblInProgressNoDay1)
        Me.Panel3.Controls.Add(Me.cmbEngineer1)
        Me.Panel3.Controls.Add(Me.lblEngineer1)
        Me.Panel3.Location = New System.Drawing.Point(12, 149)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(745, 128)
        Me.Panel3.TabIndex = 466
        '
        'lblContractsEngineer
        '
        Me.lblContractsEngineer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsEngineer.Location = New System.Drawing.Point(583, 8)
        Me.lblContractsEngineer.Name = "lblContractsEngineer"
        Me.lblContractsEngineer.Size = New System.Drawing.Size(130, 13)
        Me.lblContractsEngineer.TabIndex = 483
        Me.lblContractsEngineer.Text = "Contracts Completed"
        Me.lblContractsEngineer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCompletedEngineer
        '
        Me.lblCompletedEngineer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedEngineer.Location = New System.Drawing.Point(433, 8)
        Me.lblCompletedEngineer.Name = "lblCompletedEngineer"
        Me.lblCompletedEngineer.Size = New System.Drawing.Size(130, 13)
        Me.lblCompletedEngineer.TabIndex = 482
        Me.lblCompletedEngineer.Text = "Jobs Completed"
        Me.lblCompletedEngineer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblJobsinProgressEngineer
        '
        Me.lblJobsinProgressEngineer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobsinProgressEngineer.Location = New System.Drawing.Point(262, 8)
        Me.lblJobsinProgressEngineer.Name = "lblJobsinProgressEngineer"
        Me.lblJobsinProgressEngineer.Size = New System.Drawing.Size(130, 13)
        Me.lblJobsinProgressEngineer.TabIndex = 481
        Me.lblJobsinProgressEngineer.Text = "Jobs in Progress"
        Me.lblJobsinProgressEngineer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblContractsHoursDay2
        '
        Me.lblContractsHoursDay2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsHoursDay2.Location = New System.Drawing.Point(628, 105)
        Me.lblContractsHoursDay2.Name = "lblContractsHoursDay2"
        Me.lblContractsHoursDay2.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsHoursDay2.TabIndex = 480
        Me.lblContractsHoursDay2.Text = "00:00"
        '
        'lblContractsNoDay2
        '
        Me.lblContractsNoDay2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsNoDay2.Location = New System.Drawing.Point(588, 105)
        Me.lblContractsNoDay2.Name = "lblContractsNoDay2"
        Me.lblContractsNoDay2.Size = New System.Drawing.Size(30, 13)
        Me.lblContractsNoDay2.TabIndex = 479
        Me.lblContractsNoDay2.Text = "000"
        '
        'lblCompletedHoursDay2
        '
        Me.lblCompletedHoursDay2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedHoursDay2.Location = New System.Drawing.Point(503, 105)
        Me.lblCompletedHoursDay2.Name = "lblCompletedHoursDay2"
        Me.lblCompletedHoursDay2.Size = New System.Drawing.Size(40, 13)
        Me.lblCompletedHoursDay2.TabIndex = 478
        Me.lblCompletedHoursDay2.Text = "00:00"
        '
        'lblCompletedNoDay2
        '
        Me.lblCompletedNoDay2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedNoDay2.Location = New System.Drawing.Point(463, 105)
        Me.lblCompletedNoDay2.Name = "lblCompletedNoDay2"
        Me.lblCompletedNoDay2.Size = New System.Drawing.Size(30, 13)
        Me.lblCompletedNoDay2.TabIndex = 477
        Me.lblCompletedNoDay2.Text = "000"
        '
        'lblInProgressHoursDay2
        '
        Me.lblInProgressHoursDay2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInProgressHoursDay2.Location = New System.Drawing.Point(329, 95)
        Me.lblInProgressHoursDay2.Name = "lblInProgressHoursDay2"
        Me.lblInProgressHoursDay2.Size = New System.Drawing.Size(40, 13)
        Me.lblInProgressHoursDay2.TabIndex = 476
        Me.lblInProgressHoursDay2.Text = "00:00"
        '
        'lblInProgressNoDay2
        '
        Me.lblInProgressNoDay2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInProgressNoDay2.Location = New System.Drawing.Point(289, 95)
        Me.lblInProgressNoDay2.Name = "lblInProgressNoDay2"
        Me.lblInProgressNoDay2.Size = New System.Drawing.Size(30, 13)
        Me.lblInProgressNoDay2.TabIndex = 475
        Me.lblInProgressNoDay2.Text = "000"
        '
        'cmbEngineer2
        '
        Me.cmbEngineer2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbEngineer2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbEngineer2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbEngineer2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbEngineer2.FormattingEnabled = True
        Me.cmbEngineer2.Location = New System.Drawing.Point(75, 82)
        Me.cmbEngineer2.Name = "cmbEngineer2"
        Me.cmbEngineer2.Size = New System.Drawing.Size(179, 21)
        Me.cmbEngineer2.TabIndex = 473
        Me.cmbEngineer2.Tag = "Job_LoadEngineers"
        '
        'lblEngineer2
        '
        Me.lblEngineer2.AutoSize = True
        Me.lblEngineer2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEngineer2.Location = New System.Drawing.Point(10, 85)
        Me.lblEngineer2.Name = "lblEngineer2"
        Me.lblEngineer2.Size = New System.Drawing.Size(61, 13)
        Me.lblEngineer2.TabIndex = 474
        Me.lblEngineer2.Text = "Engineer:"
        '
        'lblContractsHoursDay1
        '
        Me.lblContractsHoursDay1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsHoursDay1.Location = New System.Drawing.Point(628, 52)
        Me.lblContractsHoursDay1.Name = "lblContractsHoursDay1"
        Me.lblContractsHoursDay1.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsHoursDay1.TabIndex = 472
        Me.lblContractsHoursDay1.Text = "00:00"
        '
        'lblContractsNoDay1
        '
        Me.lblContractsNoDay1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsNoDay1.Location = New System.Drawing.Point(588, 52)
        Me.lblContractsNoDay1.Name = "lblContractsNoDay1"
        Me.lblContractsNoDay1.Size = New System.Drawing.Size(30, 13)
        Me.lblContractsNoDay1.TabIndex = 471
        Me.lblContractsNoDay1.Text = "000"
        '
        'lblCompletedHoursDay1
        '
        Me.lblCompletedHoursDay1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedHoursDay1.Location = New System.Drawing.Point(503, 52)
        Me.lblCompletedHoursDay1.Name = "lblCompletedHoursDay1"
        Me.lblCompletedHoursDay1.Size = New System.Drawing.Size(40, 13)
        Me.lblCompletedHoursDay1.TabIndex = 470
        Me.lblCompletedHoursDay1.Text = "00:00"
        '
        'lblCompletedNoDay1
        '
        Me.lblCompletedNoDay1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedNoDay1.Location = New System.Drawing.Point(463, 52)
        Me.lblCompletedNoDay1.Name = "lblCompletedNoDay1"
        Me.lblCompletedNoDay1.Size = New System.Drawing.Size(30, 13)
        Me.lblCompletedNoDay1.TabIndex = 469
        Me.lblCompletedNoDay1.Text = "000"
        '
        'lblInProgressHoursDay1
        '
        Me.lblInProgressHoursDay1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInProgressHoursDay1.Location = New System.Drawing.Point(328, 43)
        Me.lblInProgressHoursDay1.Name = "lblInProgressHoursDay1"
        Me.lblInProgressHoursDay1.Size = New System.Drawing.Size(40, 13)
        Me.lblInProgressHoursDay1.TabIndex = 468
        Me.lblInProgressHoursDay1.Text = "00:00"
        '
        'lblInProgressNoDay1
        '
        Me.lblInProgressNoDay1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInProgressNoDay1.Location = New System.Drawing.Point(288, 43)
        Me.lblInProgressNoDay1.Name = "lblInProgressNoDay1"
        Me.lblInProgressNoDay1.Size = New System.Drawing.Size(30, 13)
        Me.lblInProgressNoDay1.TabIndex = 467
        Me.lblInProgressNoDay1.Text = "000"
        '
        'cmbEngineer1
        '
        Me.cmbEngineer1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbEngineer1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbEngineer1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbEngineer1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbEngineer1.FormattingEnabled = True
        Me.cmbEngineer1.Location = New System.Drawing.Point(75, 29)
        Me.cmbEngineer1.Name = "cmbEngineer1"
        Me.cmbEngineer1.Size = New System.Drawing.Size(179, 21)
        Me.cmbEngineer1.TabIndex = 465
        Me.cmbEngineer1.Tag = "Job_LoadEngineers"
        '
        'lblEngineer1
        '
        Me.lblEngineer1.AutoSize = True
        Me.lblEngineer1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEngineer1.Location = New System.Drawing.Point(10, 32)
        Me.lblEngineer1.Name = "lblEngineer1"
        Me.lblEngineer1.Size = New System.Drawing.Size(61, 13)
        Me.lblEngineer1.TabIndex = 466
        Me.lblEngineer1.Text = "Engineer:"
        '
        'lblContractsHoursMonth1
        '
        Me.lblContractsHoursMonth1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsHoursMonth1.Location = New System.Drawing.Point(628, 32)
        Me.lblContractsHoursMonth1.Name = "lblContractsHoursMonth1"
        Me.lblContractsHoursMonth1.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsHoursMonth1.TabIndex = 487
        Me.lblContractsHoursMonth1.Text = "00:00"
        '
        'lblContractsNoMonth1
        '
        Me.lblContractsNoMonth1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsNoMonth1.Location = New System.Drawing.Point(588, 32)
        Me.lblContractsNoMonth1.Name = "lblContractsNoMonth1"
        Me.lblContractsNoMonth1.Size = New System.Drawing.Size(30, 13)
        Me.lblContractsNoMonth1.TabIndex = 486
        Me.lblContractsNoMonth1.Text = "000"
        '
        'lblCompletedHoursMonth1
        '
        Me.lblCompletedHoursMonth1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedHoursMonth1.Location = New System.Drawing.Point(503, 32)
        Me.lblCompletedHoursMonth1.Name = "lblCompletedHoursMonth1"
        Me.lblCompletedHoursMonth1.Size = New System.Drawing.Size(40, 13)
        Me.lblCompletedHoursMonth1.TabIndex = 485
        Me.lblCompletedHoursMonth1.Text = "00:00"
        '
        'lblCompletedNoMonth1
        '
        Me.lblCompletedNoMonth1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedNoMonth1.Location = New System.Drawing.Point(463, 32)
        Me.lblCompletedNoMonth1.Name = "lblCompletedNoMonth1"
        Me.lblCompletedNoMonth1.Size = New System.Drawing.Size(30, 13)
        Me.lblCompletedNoMonth1.TabIndex = 484
        Me.lblCompletedNoMonth1.Text = "000"
        '
        'lblContractsHoursMonth2
        '
        Me.lblContractsHoursMonth2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsHoursMonth2.Location = New System.Drawing.Point(628, 85)
        Me.lblContractsHoursMonth2.Name = "lblContractsHoursMonth2"
        Me.lblContractsHoursMonth2.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsHoursMonth2.TabIndex = 491
        Me.lblContractsHoursMonth2.Text = "00:00"
        '
        'lblContractsNoMonth2
        '
        Me.lblContractsNoMonth2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsNoMonth2.Location = New System.Drawing.Point(588, 85)
        Me.lblContractsNoMonth2.Name = "lblContractsNoMonth2"
        Me.lblContractsNoMonth2.Size = New System.Drawing.Size(30, 13)
        Me.lblContractsNoMonth2.TabIndex = 490
        Me.lblContractsNoMonth2.Text = "000"
        '
        'lblCompletedHoursMonth2
        '
        Me.lblCompletedHoursMonth2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedHoursMonth2.Location = New System.Drawing.Point(503, 85)
        Me.lblCompletedHoursMonth2.Name = "lblCompletedHoursMonth2"
        Me.lblCompletedHoursMonth2.Size = New System.Drawing.Size(40, 13)
        Me.lblCompletedHoursMonth2.TabIndex = 489
        Me.lblCompletedHoursMonth2.Text = "00:00"
        '
        'lblCompletedNoMonth2
        '
        Me.lblCompletedNoMonth2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompletedNoMonth2.Location = New System.Drawing.Point(463, 85)
        Me.lblCompletedNoMonth2.Name = "lblCompletedNoMonth2"
        Me.lblCompletedNoMonth2.Size = New System.Drawing.Size(30, 13)
        Me.lblCompletedNoMonth2.TabIndex = 488
        Me.lblCompletedNoMonth2.Text = "000"
        '
        'lblContractsTravelHoursMonth1
        '
        Me.lblContractsTravelHoursMonth1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsTravelHoursMonth1.Location = New System.Drawing.Point(669, 32)
        Me.lblContractsTravelHoursMonth1.Name = "lblContractsTravelHoursMonth1"
        Me.lblContractsTravelHoursMonth1.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsTravelHoursMonth1.TabIndex = 492
        Me.lblContractsTravelHoursMonth1.Text = "[00:00]"
        '
        'lblContractsTravelHoursDay1
        '
        Me.lblContractsTravelHoursDay1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsTravelHoursDay1.Location = New System.Drawing.Point(669, 52)
        Me.lblContractsTravelHoursDay1.Name = "lblContractsTravelHoursDay1"
        Me.lblContractsTravelHoursDay1.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsTravelHoursDay1.TabIndex = 493
        Me.lblContractsTravelHoursDay1.Text = "[00:00]"
        '
        'lblContractsTravelHoursMonth2
        '
        Me.lblContractsTravelHoursMonth2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsTravelHoursMonth2.Location = New System.Drawing.Point(669, 85)
        Me.lblContractsTravelHoursMonth2.Name = "lblContractsTravelHoursMonth2"
        Me.lblContractsTravelHoursMonth2.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsTravelHoursMonth2.TabIndex = 494
        Me.lblContractsTravelHoursMonth2.Text = "[00:00]"
        '
        'lblContractsTravelHoursDay2
        '
        Me.lblContractsTravelHoursDay2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsTravelHoursDay2.Location = New System.Drawing.Point(669, 105)
        Me.lblContractsTravelHoursDay2.Name = "lblContractsTravelHoursDay2"
        Me.lblContractsTravelHoursDay2.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsTravelHoursDay2.TabIndex = 495
        Me.lblContractsTravelHoursDay2.Text = "[00:00]"
        '
        'lblContractsTravelHoursMonth
        '
        Me.lblContractsTravelHoursMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsTravelHoursMonth.Location = New System.Drawing.Point(361, 34)
        Me.lblContractsTravelHoursMonth.Name = "lblContractsTravelHoursMonth"
        Me.lblContractsTravelHoursMonth.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsTravelHoursMonth.TabIndex = 493
        Me.lblContractsTravelHoursMonth.Text = "[00:00]"
        '
        'lblContractsTravelHoursWeek
        '
        Me.lblContractsTravelHoursWeek.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsTravelHoursWeek.Location = New System.Drawing.Point(361, 56)
        Me.lblContractsTravelHoursWeek.Name = "lblContractsTravelHoursWeek"
        Me.lblContractsTravelHoursWeek.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsTravelHoursWeek.TabIndex = 494
        Me.lblContractsTravelHoursWeek.Text = "[00:00]"
        '
        'lblContractsTravelHoursDay
        '
        Me.lblContractsTravelHoursDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractsTravelHoursDay.Location = New System.Drawing.Point(361, 79)
        Me.lblContractsTravelHoursDay.Name = "lblContractsTravelHoursDay"
        Me.lblContractsTravelHoursDay.Size = New System.Drawing.Size(40, 13)
        Me.lblContractsTravelHoursDay.TabIndex = 495
        Me.lblContractsTravelHoursDay.Text = "[00:00]"
        '
        'lblMonth1
        '
        Me.lblMonth1.AutoSize = True
        Me.lblMonth1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonth1.Location = New System.Drawing.Point(399, 32)
        Me.lblMonth1.Name = "lblMonth1"
        Me.lblMonth1.Size = New System.Drawing.Size(46, 13)
        Me.lblMonth1.TabIndex = 496
        Me.lblMonth1.Text = "Month:"
        '
        'lblMonth2
        '
        Me.lblMonth2.AutoSize = True
        Me.lblMonth2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonth2.Location = New System.Drawing.Point(399, 85)
        Me.lblMonth2.Name = "lblMonth2"
        Me.lblMonth2.Size = New System.Drawing.Size(46, 13)
        Me.lblMonth2.TabIndex = 497
        Me.lblMonth2.Text = "Month:"
        '
        'lblDay1
        '
        Me.lblDay1.AutoSize = True
        Me.lblDay1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDay1.Location = New System.Drawing.Point(399, 52)
        Me.lblDay1.Name = "lblDay1"
        Me.lblDay1.Size = New System.Drawing.Size(33, 13)
        Me.lblDay1.TabIndex = 498
        Me.lblDay1.Text = "Day:"
        '
        'lblDay2
        '
        Me.lblDay2.AutoSize = True
        Me.lblDay2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDay2.Location = New System.Drawing.Point(399, 105)
        Me.lblDay2.Name = "lblDay2"
        Me.lblDay2.Size = New System.Drawing.Size(33, 13)
        Me.lblDay2.TabIndex = 499
        Me.lblDay2.Text = "Day:"
        '
        'frmStatisticsSortmyPC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(769, 289)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmStatisticsSortmyPC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Statistics Dashboard"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.nudDay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudWeek, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMonth, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tmrStatistics As System.Windows.Forms.Timer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblInProgressHoursMonth As System.Windows.Forms.Label
    Friend WithEvents lblInProgressNoMonth As System.Windows.Forms.Label
    Friend WithEvents lblInProgress As System.Windows.Forms.Label
    Friend WithEvents lblUnallocatedNoMonth As System.Windows.Forms.Label
    Friend WithEvents lblUnallocated As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents nudDay As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudWeek As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudMonth As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblDay As System.Windows.Forms.Label
    Friend WithEvents lblContractsHoursDay As System.Windows.Forms.Label
    Friend WithEvents lblContractsNoDay As System.Windows.Forms.Label
    Friend WithEvents lblCompletedHoursDay As System.Windows.Forms.Label
    Friend WithEvents lblCompletedNoDay As System.Windows.Forms.Label
    Friend WithEvents lblContractsHoursWeek As System.Windows.Forms.Label
    Friend WithEvents lblContractsNoWeek As System.Windows.Forms.Label
    Friend WithEvents lblCompletedHoursWeek As System.Windows.Forms.Label
    Friend WithEvents lblCompletedNoWeek As System.Windows.Forms.Label
    Friend WithEvents lblContractsHoursMonth As System.Windows.Forms.Label
    Friend WithEvents lblContractsNoMonth As System.Windows.Forms.Label
    Friend WithEvents lblCompletedHoursMonth As System.Windows.Forms.Label
    Friend WithEvents lblCompletedNoMonth As System.Windows.Forms.Label
    Friend WithEvents lblWeek As System.Windows.Forms.Label
    Friend WithEvents lblContracts As System.Windows.Forms.Label
    Friend WithEvents lblMonth As System.Windows.Forms.Label
    Friend WithEvents lblCompleted As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents lblContractsEngineer As System.Windows.Forms.Label
    Friend WithEvents lblCompletedEngineer As System.Windows.Forms.Label
    Friend WithEvents lblJobsinProgressEngineer As System.Windows.Forms.Label
    Friend WithEvents lblContractsHoursDay2 As System.Windows.Forms.Label
    Friend WithEvents lblContractsNoDay2 As System.Windows.Forms.Label
    Friend WithEvents lblCompletedHoursDay2 As System.Windows.Forms.Label
    Friend WithEvents lblCompletedNoDay2 As System.Windows.Forms.Label
    Friend WithEvents lblInProgressHoursDay2 As System.Windows.Forms.Label
    Friend WithEvents lblInProgressNoDay2 As System.Windows.Forms.Label
    Friend WithEvents cmbEngineer2 As System.Windows.Forms.ComboBox
    Friend WithEvents lblEngineer2 As System.Windows.Forms.Label
    Friend WithEvents lblContractsHoursDay1 As System.Windows.Forms.Label
    Friend WithEvents lblContractsNoDay1 As System.Windows.Forms.Label
    Friend WithEvents lblCompletedHoursDay1 As System.Windows.Forms.Label
    Friend WithEvents lblCompletedNoDay1 As System.Windows.Forms.Label
    Friend WithEvents lblInProgressHoursDay1 As System.Windows.Forms.Label
    Friend WithEvents lblInProgressNoDay1 As System.Windows.Forms.Label
    Friend WithEvents cmbEngineer1 As System.Windows.Forms.ComboBox
    Friend WithEvents lblEngineer1 As System.Windows.Forms.Label
    Friend WithEvents lblContractsHoursMonth2 As System.Windows.Forms.Label
    Friend WithEvents lblContractsNoMonth2 As System.Windows.Forms.Label
    Friend WithEvents lblCompletedHoursMonth2 As System.Windows.Forms.Label
    Friend WithEvents lblCompletedNoMonth2 As System.Windows.Forms.Label
    Friend WithEvents lblContractsHoursMonth1 As System.Windows.Forms.Label
    Friend WithEvents lblContractsNoMonth1 As System.Windows.Forms.Label
    Friend WithEvents lblCompletedHoursMonth1 As System.Windows.Forms.Label
    Friend WithEvents lblCompletedNoMonth1 As System.Windows.Forms.Label
    Friend WithEvents lblContractsTravelHoursDay As System.Windows.Forms.Label
    Friend WithEvents lblContractsTravelHoursWeek As System.Windows.Forms.Label
    Friend WithEvents lblContractsTravelHoursMonth As System.Windows.Forms.Label
    Friend WithEvents lblContractsTravelHoursDay2 As System.Windows.Forms.Label
    Friend WithEvents lblContractsTravelHoursMonth2 As System.Windows.Forms.Label
    Friend WithEvents lblContractsTravelHoursDay1 As System.Windows.Forms.Label
    Friend WithEvents lblContractsTravelHoursMonth1 As System.Windows.Forms.Label
    Friend WithEvents lblDay2 As System.Windows.Forms.Label
    Friend WithEvents lblDay1 As System.Windows.Forms.Label
    Friend WithEvents lblMonth2 As System.Windows.Forms.Label
    Friend WithEvents lblMonth1 As System.Windows.Forms.Label
End Class
