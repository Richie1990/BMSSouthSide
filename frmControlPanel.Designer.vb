<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmControlPanel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmControlPanel))
        Me.lblCurrentUser = New System.Windows.Forms.Label()
        Me.lblDeveloper = New System.Windows.Forms.Label()
        Me.cmdUtilities = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdProduct = New System.Windows.Forms.Button()
        Me.cmdJobsList = New System.Windows.Forms.Button()
        Me.cmdLandlords = New System.Windows.Forms.Button()
        Me.lblCompany = New System.Windows.Forms.Label()
        Me.cmdSupplier = New System.Windows.Forms.Button()
        Me.cmdOrdersList = New System.Windows.Forms.Button()
        Me.cmdQuotes = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdFinance = New System.Windows.Forms.Button()
        Me.cmdEngineers = New System.Windows.Forms.Button()
        Me.cmdTimeSheets = New System.Windows.Forms.Button()
        Me.cmdClientRegister = New System.Windows.Forms.Button()
        Me.cmdStatistics = New System.Windows.Forms.Button()
        Me.lblBMSPath = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblCurrentUser
        '
        Me.lblCurrentUser.Location = New System.Drawing.Point(336, 438)
        Me.lblCurrentUser.Name = "lblCurrentUser"
        Me.lblCurrentUser.Size = New System.Drawing.Size(257, 18)
        Me.lblCurrentUser.TabIndex = 31
        Me.lblCurrentUser.Text = "Current User:"
        '
        'lblDeveloper
        '
        Me.lblDeveloper.Location = New System.Drawing.Point(336, 410)
        Me.lblDeveloper.Name = "lblDeveloper"
        Me.lblDeveloper.Size = New System.Drawing.Size(330, 18)
        Me.lblDeveloper.TabIndex = 30
        Me.lblDeveloper.Text = "Developed by Angus W Kerr, SortMyPC Ltd"
        '
        'cmdUtilities
        '
        Me.cmdUtilities.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUtilities.Image = CType(resources.GetObject("cmdUtilities.Image"), System.Drawing.Image)
        Me.cmdUtilities.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdUtilities.Location = New System.Drawing.Point(1020, 196)
        Me.cmdUtilities.Name = "cmdUtilities"
        Me.cmdUtilities.Size = New System.Drawing.Size(150, 166)
        Me.cmdUtilities.TabIndex = 29
        Me.cmdUtilities.Text = "Utilities"
        Me.cmdUtilities.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdUtilities.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Image = CType(resources.GetObject("cmdExit.Image"), System.Drawing.Image)
        Me.cmdExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdExit.Location = New System.Drawing.Point(1072, 395)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(98, 125)
        Me.cmdExit.TabIndex = 28
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdProduct
        '
        Me.cmdProduct.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdProduct.Image = CType(resources.GetObject("cmdProduct.Image"), System.Drawing.Image)
        Me.cmdProduct.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdProduct.Location = New System.Drawing.Point(620, 12)
        Me.cmdProduct.Name = "cmdProduct"
        Me.cmdProduct.Size = New System.Drawing.Size(150, 166)
        Me.cmdProduct.TabIndex = 27
        Me.cmdProduct.Text = "Products"
        Me.cmdProduct.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdProduct.UseVisualStyleBackColor = True
        Me.cmdProduct.Visible = False
        '
        'cmdJobsList
        '
        Me.cmdJobsList.Enabled = False
        Me.cmdJobsList.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdJobsList.Image = CType(resources.GetObject("cmdJobsList.Image"), System.Drawing.Image)
        Me.cmdJobsList.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdJobsList.Location = New System.Drawing.Point(820, 196)
        Me.cmdJobsList.Name = "cmdJobsList"
        Me.cmdJobsList.Size = New System.Drawing.Size(150, 166)
        Me.cmdJobsList.TabIndex = 26
        Me.cmdJobsList.Text = "Maintenance"
        Me.cmdJobsList.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdJobsList.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdJobsList.UseVisualStyleBackColor = True
        '
        'cmdLandlords
        '
        Me.cmdLandlords.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLandlords.Image = CType(resources.GetObject("cmdLandlords.Image"), System.Drawing.Image)
        Me.cmdLandlords.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdLandlords.Location = New System.Drawing.Point(12, 12)
        Me.cmdLandlords.Name = "cmdLandlords"
        Me.cmdLandlords.Size = New System.Drawing.Size(150, 166)
        Me.cmdLandlords.TabIndex = 25
        Me.cmdLandlords.Text = "Landlords"
        Me.cmdLandlords.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdLandlords.UseVisualStyleBackColor = True
        '
        'lblCompany
        '
        Me.lblCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompany.Location = New System.Drawing.Point(336, 382)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(300, 17)
        Me.lblCompany.TabIndex = 32
        Me.lblCompany.Text = "Registered To:"
        '
        'cmdSupplier
        '
        Me.cmdSupplier.Enabled = False
        Me.cmdSupplier.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSupplier.Image = CType(resources.GetObject("cmdSupplier.Image"), System.Drawing.Image)
        Me.cmdSupplier.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdSupplier.Location = New System.Drawing.Point(420, 12)
        Me.cmdSupplier.Name = "cmdSupplier"
        Me.cmdSupplier.Size = New System.Drawing.Size(150, 166)
        Me.cmdSupplier.TabIndex = 33
        Me.cmdSupplier.Text = "Tenants"
        Me.cmdSupplier.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdSupplier.UseVisualStyleBackColor = True
        '
        'cmdOrdersList
        '
        Me.cmdOrdersList.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOrdersList.Image = CType(resources.GetObject("cmdOrdersList.Image"), System.Drawing.Image)
        Me.cmdOrdersList.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdOrdersList.Location = New System.Drawing.Point(420, 196)
        Me.cmdOrdersList.Name = "cmdOrdersList"
        Me.cmdOrdersList.Size = New System.Drawing.Size(150, 166)
        Me.cmdOrdersList.TabIndex = 34
        Me.cmdOrdersList.Text = "Orders List"
        Me.cmdOrdersList.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdOrdersList.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdOrdersList.UseVisualStyleBackColor = True
        Me.cmdOrdersList.Visible = False
        '
        'cmdQuotes
        '
        Me.cmdQuotes.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuotes.Image = CType(resources.GetObject("cmdQuotes.Image"), System.Drawing.Image)
        Me.cmdQuotes.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdQuotes.Location = New System.Drawing.Point(12, 196)
        Me.cmdQuotes.Name = "cmdQuotes"
        Me.cmdQuotes.Size = New System.Drawing.Size(150, 166)
        Me.cmdQuotes.TabIndex = 35
        Me.cmdQuotes.Text = "Quotes"
        Me.cmdQuotes.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdQuotes.UseVisualStyleBackColor = True
        Me.cmdQuotes.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.InitialImage = CType(resources.GetObject("PictureBox1.InitialImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 373)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(303, 147)
        Me.PictureBox1.TabIndex = 36
        Me.PictureBox1.TabStop = False
        '
        'cmdFinance
        '
        Me.cmdFinance.Enabled = False
        Me.cmdFinance.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFinance.Image = CType(resources.GetObject("cmdFinance.Image"), System.Drawing.Image)
        Me.cmdFinance.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdFinance.Location = New System.Drawing.Point(820, 12)
        Me.cmdFinance.Name = "cmdFinance"
        Me.cmdFinance.Size = New System.Drawing.Size(150, 166)
        Me.cmdFinance.TabIndex = 37
        Me.cmdFinance.Text = "Finance"
        Me.cmdFinance.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdFinance.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdFinance.UseVisualStyleBackColor = True
        '
        'cmdEngineers
        '
        Me.cmdEngineers.Enabled = False
        Me.cmdEngineers.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEngineers.Image = CType(resources.GetObject("cmdEngineers.Image"), System.Drawing.Image)
        Me.cmdEngineers.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdEngineers.Location = New System.Drawing.Point(217, 12)
        Me.cmdEngineers.Name = "cmdEngineers"
        Me.cmdEngineers.Size = New System.Drawing.Size(150, 166)
        Me.cmdEngineers.TabIndex = 38
        Me.cmdEngineers.Text = "Properties"
        Me.cmdEngineers.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdEngineers.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdEngineers.UseVisualStyleBackColor = True
        '
        'cmdTimeSheets
        '
        Me.cmdTimeSheets.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTimeSheets.Image = CType(resources.GetObject("cmdTimeSheets.Image"), System.Drawing.Image)
        Me.cmdTimeSheets.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdTimeSheets.Location = New System.Drawing.Point(620, 196)
        Me.cmdTimeSheets.Name = "cmdTimeSheets"
        Me.cmdTimeSheets.Size = New System.Drawing.Size(150, 166)
        Me.cmdTimeSheets.TabIndex = 39
        Me.cmdTimeSheets.Text = "Time Sheets"
        Me.cmdTimeSheets.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdTimeSheets.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdTimeSheets.UseVisualStyleBackColor = True
        Me.cmdTimeSheets.Visible = False
        '
        'cmdClientRegister
        '
        Me.cmdClientRegister.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClientRegister.Image = CType(resources.GetObject("cmdClientRegister.Image"), System.Drawing.Image)
        Me.cmdClientRegister.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdClientRegister.Location = New System.Drawing.Point(217, 196)
        Me.cmdClientRegister.Name = "cmdClientRegister"
        Me.cmdClientRegister.Size = New System.Drawing.Size(150, 166)
        Me.cmdClientRegister.TabIndex = 40
        Me.cmdClientRegister.Text = "Maintenance"
        Me.cmdClientRegister.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdClientRegister.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdClientRegister.UseVisualStyleBackColor = True
        Me.cmdClientRegister.Visible = False
        '
        'cmdStatistics
        '
        Me.cmdStatistics.Enabled = False
        Me.cmdStatistics.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStatistics.Image = CType(resources.GetObject("cmdStatistics.Image"), System.Drawing.Image)
        Me.cmdStatistics.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdStatistics.Location = New System.Drawing.Point(1020, 12)
        Me.cmdStatistics.Name = "cmdStatistics"
        Me.cmdStatistics.Size = New System.Drawing.Size(150, 166)
        Me.cmdStatistics.TabIndex = 41
        Me.cmdStatistics.Text = "Statistics"
        Me.cmdStatistics.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdStatistics.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdStatistics.UseVisualStyleBackColor = True
        '
        'lblBMSPath
        '
        Me.lblBMSPath.Location = New System.Drawing.Point(336, 466)
        Me.lblBMSPath.Name = "lblBMSPath"
        Me.lblBMSPath.Size = New System.Drawing.Size(553, 18)
        Me.lblBMSPath.TabIndex = 42
        Me.lblBMSPath.Text = "BMS Path:"
        '
        'frmControlPanel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1195, 532)
        Me.Controls.Add(Me.lblBMSPath)
        Me.Controls.Add(Me.cmdStatistics)
        Me.Controls.Add(Me.cmdClientRegister)
        Me.Controls.Add(Me.cmdTimeSheets)
        Me.Controls.Add(Me.cmdEngineers)
        Me.Controls.Add(Me.cmdFinance)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdQuotes)
        Me.Controls.Add(Me.cmdOrdersList)
        Me.Controls.Add(Me.cmdSupplier)
        Me.Controls.Add(Me.lblCompany)
        Me.Controls.Add(Me.lblCurrentUser)
        Me.Controls.Add(Me.lblDeveloper)
        Me.Controls.Add(Me.cmdUtilities)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdProduct)
        Me.Controls.Add(Me.cmdJobsList)
        Me.Controls.Add(Me.cmdLandlords)
        Me.Name = "frmControlPanel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Business Management System - Control Panel"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdLandlords As System.Windows.Forms.Button
    Friend WithEvents cmdJobsList As System.Windows.Forms.Button
    Friend WithEvents cmdProduct As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdUtilities As System.Windows.Forms.Button
    Friend WithEvents lblCurrentUser As System.Windows.Forms.Label
    Friend WithEvents lblDeveloper As System.Windows.Forms.Label
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents cmdSupplier As System.Windows.Forms.Button
    Friend WithEvents cmdOrdersList As System.Windows.Forms.Button
    Friend WithEvents cmdQuotes As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdFinance As System.Windows.Forms.Button
    Friend WithEvents cmdEngineers As System.Windows.Forms.Button
    Friend WithEvents cmdTimeSheets As System.Windows.Forms.Button
    Friend WithEvents cmdClientRegister As System.Windows.Forms.Button
    Friend WithEvents cmdStatistics As System.Windows.Forms.Button
    Friend WithEvents lblBMSPath As System.Windows.Forms.Label
End Class
