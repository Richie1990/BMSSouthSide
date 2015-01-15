Imports System.IO

Public Class frmUtilities

    Dim objLandlord As New clsLandlord

    Private Sub cmdReportDesign_Click(sender As Object, e As EventArgs) Handles cmdDesignReport.Click
        ' Error Checking
        On Error GoTo Err_cmdReportDesign_Click

        Dim ListReport = New FastReport.Report
        ListReport.Load(cmbReportList.Text)
        ListReport.SetParameterValue("CRMConnectionString", "Data Source=" & My.Settings.SQLServer & ";AttachDbFilename=;Initial Catalog=" & My.Settings.SQLDatabase & ";Integrated Security=False;Persist Security Info=False;User ID=CRMUser;Password=S0rtmypc!")
        ListReport.Design()

Err_cmdReportDesign_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdReportDesign_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdReportDesign_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If
    End Sub

    Private Sub frmUtilities_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Read List of Files
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(SystemPath & "Reports", FileIO.SearchOption.SearchTopLevelOnly, "*.frx")
            cmbReportList.Items.Add(foundFile)

        Next

    End Sub

    Private Sub cmdPreviewReport_Click(sender As Object, e As EventArgs) Handles cmdPreviewReport.Click
        ' Error Checking
        On Error GoTo Err_cmdPreviewReport_Click

        Dim ListReport = New FastReport.Report
        ListReport.Load(cmbReportList.Text)
        ListReport.SetParameterValue("CRMConnectionString", "Data Source=" & My.Settings.SQLServer & ";AttachDbFilename=;Initial Catalog=" & My.Settings.SQLDatabase & ";Integrated Security=False;Persist Security Info=False;User ID=CRMUser;Password=S0rtmypc!")
        ListReport.SetParameterValue("JobID", 1)
        ListReport.SetParameterValue("CustomerCopy", True)
        ListReport.Show()

Err_cmdPreviewReport_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdPreviewReport_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdPreviewReport_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

 
    Private Sub cmdImportLandlords_Click(sender As Object, e As EventArgs) Handles cmdImportLandlords.Click
        '**Import Landlords Data from Access Database

        ' Error Checking
        On Error GoTo Err_cmdImportLandlords_Click

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uAccessCommand As ADODB.Command

        ' Check For Open Connections
        If uDBase Is Nothing Then
            OpenConnection()
            If uAccessDBase Is Nothing Then
                OpenAccessConnection()
                bConnection = True
            End If
        End If

        ' Run Stored Procedure - Load Landlord Records
        uAccessCommand = New ADODB.Command
        With uAccessCommand
            .ActiveConnection = uAccessDBase
            .CommandTimeout = 0
            .CommandText = "SELECT * FROM Landlords"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        Do Until uRecSnap.EOF

            ' Run Stored Procedure - Save Landlord Record
            lblRecordID.Text = Format(uRecSnap("LandlordID").Value, "00000")
            lblRecordID.Refresh()

            With objLandlord
                .LandlordID = uRecSnap("LandlordID").Value
                .Title1 = ""
                .Forename1 = ""
                .Surname1 = ""
                .DOB1 = If(IsDBNull(uRecSnap("DateofBirth").Value), "01/01/1900", uRecSnap("DateofBirth").Value)
                .PlaceofBirth1 = If(IsDBNull(uRecSnap("PlaceofBirth").Value), "", uRecSnap("PlaceofBirth").Value)
                .Title2 = ""
                .Forename2 = ""
                .Surname2 = ""
                .DOB2 = "01/01/1900"
                .PlaceofBirth2 = ""
                .CommsTitle = If(IsDBNull(uRecSnap("Dear").Value), "", uRecSnap("Dear").Value)
                .LandlordName = If(IsDBNull(uRecSnap("Landlordname").Value), "", uRecSnap("Landlordname").Value)
                .Address1 = If(IsDBNull(uRecSnap("LandlordAddr1").Value), "", uRecSnap("LandlordAddr1").Value)
                .Address2 = If(IsDBNull(uRecSnap("LandlordAddr2").Value), "", uRecSnap("LandlordAddr2").Value)
                .Address3 = ""
                .Town = If(IsDBNull(uRecSnap("Landlordcity").Value), "", uRecSnap("Landlordcity").Value)
                .PostCode = If(IsDBNull(uRecSnap("LandlordPostCode").Value), "", uRecSnap("LandlordPostCode").Value)
                If IsDBNull(uRecSnap("Phone").Value) Then
                    .PhoneNo = ""
                Else
                    If Len(uRecSnap("Phone").Value) > 14 Then
                        .PhoneNo = uRecSnap("Phone").Value.ToString.Substring(0, 14)
                    Else
                        .PhoneNo = uRecSnap("Phone").Value
                    End If
                End If
                .MobileNo = ""
                .EMailAddress = If(IsDBNull(uRecSnap("EMail").Value), "", uRecSnap("EMail").Value)
                .CommsStatus = If(IsDBNull(uRecSnap("CommsMethod").Value), "", uRecSnap("CommsMethod").Value)
                .LandlordNotes = If(IsDBNull(uRecSnap("Notes").Value), "", uRecSnap("Notes").Value)
                .SourceID = 0
                .JoinDate = If(IsDBNull(uRecSnap("LandlordGained").Value), "01/01/1900", uRecSnap("LandlordGained").Value)
                .BankName = If(IsDBNull(uRecSnap("RentTo").Value), "", uRecSnap("RentTo").Value)
                .Sortcode = If(IsDBNull(uRecSnap("Sortcode").Value), 0, Val(uRecSnap("Sortcode").Value))
                '.AccountNo = 0
                .AccountNo = If(IsDBNull(uRecSnap("AccountCode").Value), 0, Val(uRecSnap("AccountCode").Value))
                .LStatus = True
                If .SaveLandlord() = False Then ' Error Occurred
                    lstImportErrors.Items.Add(.LandlordID & " - " & .LandlordName)
                    lstImportErrors.Refresh()
                End If
            End With
            Me.Refresh()
            uRecSnap.MoveNext()
        Loop

        ' Close Connections
        If bConnection Then
            CloseAccessConnection()
            CloseConnection()
        End If

        MsgBox("Landlord Data Imported successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AztecCRM - Contact Information")

Err_cmdImportLandlords_Click:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord(Me.Name, "cmdImportLandlords_Click", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "cmdImportLandlords_Click" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If


    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
End Class