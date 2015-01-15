Imports System.Globalization

Public Class frmStatisticsSortmyPC

    ' Dimension Form Variables
    Dim objJob As New clsJob


    Private Sub frmStatisticsSortmyPC_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        frmControlPanel.cmdStatistics.Enabled = True

    End Sub

    Private Sub frmStatisticsSortmyPC_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Dimension Local Variables
        Dim uCalendar As Calendar = New GregorianCalendar

        nudMonth.Text = Format(Month(Now), "00")
        nudWeek.Text = Format(uCalendar.GetWeekOfYear(Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday), "00")
        nudDay.Text = Format(Now, "dd")

        PopulateCombo(cmbEngineer1)
        PopulateCombo(cmbEngineer2)
        LoadGeneralStatistics()

    End Sub


    Private Sub tmrStatistics_Tick(sender As Object, e As EventArgs) Handles tmrStatistics.Tick

        LoadGeneralStatistics()
        LoadGeneralStatistics()
        LoadEngineer1Statistics()
        LoadEngineer2Statistics()

    End Sub


    Private Sub nudMonth_ValueChanged_1(sender As Object, e As EventArgs) Handles nudMonth.ValueChanged
        LoadGeneralStatistics()
        LoadGeneralStatistics()
        LoadEngineer1Statistics()
        LoadEngineer2Statistics()

    End Sub

    Private Sub nudWeek_ValueChanged_1(sender As Object, e As EventArgs) Handles nudWeek.ValueChanged
        LoadGeneralStatistics()

    End Sub

    Private Sub nudDay_ValueChanged_1(sender As Object, e As EventArgs) Handles nudDay.ValueChanged
        LoadGeneralStatistics()
        LoadEngineer1Statistics()
        LoadEngineer2Statistics()

    End Sub

    Private Sub cmbEngineer1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cmbEngineer1.SelectedIndexChanged
        LoadEngineer1Statistics()

    End Sub

    Private Sub cmbEngineer2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cmbEngineer2.SelectedIndexChanged
        LoadEngineer2Statistics()

    End Sub

    Sub LoadGeneralStatistics()

        With objJob
            .EngineerID = 0

            .Month = nudMonth.Value
            .LoadStatisticsByMonth()
            lblUnallocatedNoMonth.Text = Format(.UnallocatedNo, "000")
            lblInProgressNoMonth.Text = Format(.InProgressNo, "000")
            lblInProgressHoursMonth.Text = Format(Int(.InProgressHours), "00") & ":" & Format((.InProgressHours - Int(.InProgressHours)) * 100, "00")

            lblCompletedNoMonth.Text = Format(.CompletedNo, "000")
            lblCompletedHoursMonth.Text = Format(Int(.CompletedHours), "00") & ":" & Format((.CompletedHours - Int(.CompletedHours)) * 100, "00")
            lblContractsNoMonth.Text = Format(.ContractsNo, "000")
            lblContractsHoursMonth.Text = Format(Int(.ContractsHours), "00") & ":" & Format((.ContractsHours - Int(.ContractsHours)) * 100, "00")
            lblContractsTravelHoursMonth.Text = "[" & Format(Int(.ContractsTravelHours), "00") & ":" & Format((.ContractsTravelHours - Int(.ContractsTravelHours)) * 100, "00") & "]"

            .Week = nudWeek.Value
            .LoadStatisticsByWeek()
            lblCompletedNoWeek.Text = Format(.CompletedNo, "000")
            lblCompletedHoursWeek.Text = Format(Int(.CompletedHours), "00") & ":" & Format((.CompletedHours - Int(.CompletedHours)) * 100, "00")
            lblContractsNoWeek.Text = Format(.ContractsNo, "000")
            lblContractsHoursWeek.Text = Format(Int(.ContractsHours), "00") & ":" & Format((.ContractsHours - Int(.ContractsHours)) * 100, "00")
            lblContractsTravelHoursWeek.Text = "[" & Format(Int(.ContractsTravelHours), "00") & ":" & Format((.ContractsTravelHours - Int(.ContractsTravelHours)) * 100, "00") & "]"

            .Day = nudDay.Value
            .LoadStatisticsByDay()
            lblCompletedNoDay.Text = Format(.CompletedNo, "000")
            lblCompletedHoursDay.Text = Format(Int(.CompletedHours), "00") & ":" & Format((.CompletedHours - Int(.CompletedHours)) * 100, "00")
            lblContractsNoDay.Text = Format(.ContractsNo, "000")
            lblContractsHoursDay.Text = Format(Int(.ContractsHours), "00") & ":" & Format((.ContractsHours - Int(.ContractsHours)) * 100, "00")
            lblContractsTravelHoursDay.Text = "[" & Format(Int(.ContractsTravelHours), "00") & ":" & Format((.ContractsTravelHours - Int(.ContractsTravelHours)) * 100, "00") & "]"

        End With
    End Sub

    Sub LoadEngineer1Statistics()
        With objJob
            .EngineerID = cmbEngineer1.SelectedValue
            .Month = nudMonth.Value
            .LoadStatisticsByMonth()
            lblCompletedNoMonth1.Text = Format(.CompletedNo, "000")
            lblCompletedHoursMonth1.Text = Format(Int(.CompletedHours), "00") & ":" & Format((.CompletedHours - Int(.CompletedHours)) * 100, "00")
            lblContractsNoMonth1.Text = Format(.ContractsNo, "000")
            lblContractsHoursMonth1.Text = Format(Int(.ContractsHours), "00") & ":" & Format((.ContractsHours - Int(.ContractsHours)) * 100, "00")
            lblContractsTravelHoursMonth1.Text = "[" & Format(Int(.ContractsTravelHours), "00") & ":" & Format((.ContractsTravelHours - Int(.ContractsTravelHours)) * 100, "00") & "]"

            .Day = nudDay.Value
            .LoadStatisticsByDay()
            lblInProgressNoDay1.Text = Format(.InProgressNo, "000")
            lblInProgressHoursDay1.Text = Format(Int(.InProgressHours), "00") & ":" & Format((.InProgressHours - Int(.InProgressHours)) * 100, "00")
            lblCompletedNoDay1.Text = Format(.CompletedNo, "000")
            lblCompletedHoursDay1.Text = Format(Int(.CompletedHours), "00") & ":" & Format((.CompletedHours - Int(.CompletedHours)) * 100, "00")
            lblContractsNoDay1.Text = Format(.ContractsNo, "000")
            lblContractsHoursDay1.Text = Format(Int(.ContractsHours), "00") & ":" & Format((.ContractsHours - Int(.ContractsHours)) * 100, "00")
            lblContractsTravelHoursDay1.Text = "[" & Format(Int(.ContractsTravelHours), "00") & ":" & Format((.ContractsTravelHours - Int(.ContractsTravelHours)) * 100, "00") & "]"
        End With

    End Sub

    Sub LoadEngineer2Statistics()
        With objJob
            .EngineerID = cmbEngineer2.SelectedValue
            .Month = nudMonth.Value
            .LoadStatisticsByMonth()
            lblCompletedNoMonth2.Text = Format(.CompletedNo, "000")
            lblCompletedHoursMonth2.Text = Format(Int(.CompletedHours), "00") & ":" & Format((.CompletedHours - Int(.CompletedHours)) * 100, "00")
            lblContractsNoMonth2.Text = Format(.ContractsNo, "000")
            lblContractsHoursMonth2.Text = Format(Int(.ContractsHours), "00") & ":" & Format((.ContractsHours - Int(.ContractsHours)) * 100, "00")
            lblContractsTravelHoursMonth2.Text = "[" & Format(Int(.ContractsTravelHours), "00") & ":" & Format((.ContractsTravelHours - Int(.ContractsTravelHours)) * 100, "00") & "]"

            .Day = nudDay.Value
            .LoadStatisticsByDay()
            lblInProgressNoDay2.Text = Format(.InProgressNo, "000")
            lblInProgressHoursDay2.Text = Format(Int(.InProgressHours), "00") & ":" & Format((.InProgressHours - Int(.InProgressHours)) * 100, "00")
            lblCompletedNoDay2.Text = Format(.CompletedNo, "000")
            lblCompletedHoursDay2.Text = Format(Int(.CompletedHours), "00") & ":" & Format((.CompletedHours - Int(.CompletedHours)) * 100, "00")
            lblContractsNoDay2.Text = Format(.ContractsNo, "000")
            lblContractsHoursDay2.Text = Format(Int(.ContractsHours), "00") & ":" & Format((.ContractsHours - Int(.ContractsHours)) * 100, "00")
            lblContractsTravelHoursDay2.Text = "[" & Format(Int(.ContractsTravelHours), "00") & ":" & Format((.ContractsTravelHours - Int(.ContractsTravelHours)) * 100, "00") & "]"
        End With

    End Sub
End Class