Public Class frmControlPanel

    Private Sub frmControlPanel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Error Checking
        On Error GoTo Err_Entry_Load

        ' Dimension Local Variables
        Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)

        ' Update Company / Product Name / Version No
        lblCompany.Text = "Registered to: " & My.Settings.CompanyName
        lblCurrentUser.Text = "Current User: " & Environment.UserName
        lblBMSPath.Text = "BMS Path: " & SystemPath
        lblDeveloper.Text = lblDeveloper.Text & " - Version No: " & myBuildInfo.FileMajorPart & "." & myBuildInfo.FileMinorPart & "." & myBuildInfo.FileBuildPart

Err_Entry_Load:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "Entry_Load", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "Entry_Load" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub cmdLandlords_Click(sender As Object, e As EventArgs) Handles cmdLandlords.Click
        ' Error Checking
        '        On Error GoTo Err_cmdLandlords_Click

        ' Load Client
        Dim uClient As New frmLandlord
        With uClient
            .ShowDialog()
        End With
        cmdLandlords.Enabled = True

Err_cmdLandlords_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdLandlords_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdLandlords_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub cmdUtilities_Click(sender As Object, e As EventArgs) Handles cmdUtilities.Click
        ' Error Checking
        '        On Error GoTo Err_cmdUtilities_Click

        ' Load Product
        Dim uUtilities As New frmUtilities
        With uUtilities
            .ShowDialog()
        End With
        cmdUtilities.Enabled = True

Err_cmdUtilities_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdUtilities_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdUtilities_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        End

    End Sub

    Private Sub cmdSuppliers_Click(sender As Object, e As EventArgs) Handles cmdSupplier.Click
        ' Error Checking
        On Error GoTo Err_cmdSuppliers_Click

        ' Load Supplier
        Dim uSupplier As New frmSupplier
        With uSupplier
            .ShowDialog()
        End With
        cmdSupplier.Enabled = True

Err_cmdSuppliers_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdSuppliers_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdSuppliers_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Private Sub cmdFinance_Click(sender As Object, e As EventArgs) Handles cmdFinance.Click
        ' Error Checking
        On Error GoTo Err_cmdFinance_Click

        ' Load Product
        Dim uFinance As New frmFinance
        With uFinance
            .ShowDialog()
        End With
        cmdFinance.Enabled = True

Err_cmdFinance_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdFinance_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdFinance_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If


    End Sub

    Private Sub cmdStatistics_Click(sender As Object, e As EventArgs) Handles cmdStatistics.Click
        ' Error Checking
        On Error GoTo Err_cmdStatistics_Click

        ' Load TimeSheet
        ' Load Statistics Form
        Dim uStatistics As New frmStatisticsSortmyPC
        With uStatistics
            .Show()
        End With
        cmdStatistics.Enabled = False


Err_cmdStatistics_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdStatistics_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdStatistics_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If


    End Sub
End Class