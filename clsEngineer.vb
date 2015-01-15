Public Class clsEngineer
    ' Dimension Class Variables
    Private iEngineerID As Integer
    Private sTitle As String
    Private sForename As String
    Private sSurname As String
    Private sUserName As String
    Private sMobileNo As String
    Private sEMailAddress As String
    Private dRateNormal As Double
    Private dRateOvertime As Double
    Private dRateDouble As Double
    Private dRateTravel As Double
    Private bELive As Boolean



    Public Property EngineerID() As Integer
        '** Expose Engineer ID
        Get
            EngineerID = iEngineerID
        End Get
        Set(ByVal Value As Integer)
            iEngineerID = Value
        End Set

    End Property

    Public Property Title() As String
        '** Expose Title
        Get
            Title = sTitle
        End Get
        Set(ByVal Value As String)
            sTitle = Value
        End Set

    End Property

    Public Property Forename() As String
        '** Expose Forename
        Get
            Forename = sForename
        End Get
        Set(ByVal Value As String)
            sForename = Value
        End Set

    End Property

    Public Property Surname() As String
        '** Expose Surname
        Get
            Surname = sSurname
        End Get
        Set(ByVal Value As String)
            sSurname = Value
        End Set

    End Property

    Public Property UserName() As String
        '** Expose UserName
        Get
            UserName = sUserName
        End Get
        Set(ByVal Value As String)
            sUserName = Value
        End Set

    End Property

    Public Property MobileNo() As String
        '** Expose Mobile No
        Get
            MobileNo = sMobileNo
        End Get
        Set(ByVal Value As String)
            sMobileNo = Value
        End Set

    End Property

    Public Property EMailAddress() As String
        '** Expose EMail Address
        Get
            EMailAddress = sEMailAddress
        End Get
        Set(ByVal Value As String)
            sEMailAddress = Value
        End Set

    End Property

    Public Property RateNormal() As String
        '** Expose Rate Normal
        Get
            RateNormal = dRateNormal
        End Get
        Set(ByVal Value As String)
            dRateNormal = Value
        End Set

    End Property

    Public Property RateOvertime() As String
        '** Expose Rate Overtime
        Get
            RateOvertime = dRateOvertime
        End Get
        Set(ByVal Value As String)
            dRateOvertime = Value
        End Set

    End Property

    Public Property RateDouble() As String
        '** Expose Rate Double
        Get
            RateDouble = dRateDouble
        End Get
        Set(ByVal Value As String)
            dRateDouble = Value
        End Set

    End Property

    Public Property RateTravel() As String
        '** Expose Rate Travel
        Get
            RateTravel = dRateTravel
        End Get
        Set(ByVal Value As String)
            dRateTravel = Value
        End Set

    End Property

    Public Property ELive() As Boolean
        '** Expose Live Status
        Get
            ELive = bELive
        End Get
        Set(ByVal Value As Boolean)
            bELive = Value
        End Set

    End Property

    Public Function LoadEngineer() As Boolean
        '**Finds the Current Engineer Data
        '**Raises an Error if Engineer not found

        ' Error Checking
        On Error GoTo Err_LoadEngineer

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Client Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            .CommandText = "Engineer_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Check for Valid Client - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Engineer Class Object", sErrDescription)
            LoadEngineer = False
            Exit Function
        End If

        ' Store Data Values
        sTitle = If(IsDBNull(uRecSnap("Title").Value), "", uRecSnap("Title").Value)
        sForename = If(IsDBNull(uRecSnap("Forename").Value), "", uRecSnap("Forename").Value)
        sSurname = If(IsDBNull(uRecSnap("Surname").Value), "", uRecSnap("Surname").Value)
        sUserName = If(IsDBNull(uRecSnap("UserName").Value), "", uRecSnap("UserName").Value)
        sMobileNo = If(IsDBNull(uRecSnap("MobileNo").Value), "", uRecSnap("MobileNo").Value)
        sEMailAddress = If(IsDBNull(uRecSnap("EMailAddress").Value), "", uRecSnap("EMailAddress").Value)
        dRateNormal = If(IsDBNull(uRecSnap("RateNormal").Value), "", uRecSnap("RateNormal").Value)
        dRateOvertime = If(IsDBNull(uRecSnap("RateOvertime").Value), "", uRecSnap("RateOvertime").Value)
        dRateDouble = If(IsDBNull(uRecSnap("RateDouble").Value), "", uRecSnap("RateDouble").Value)
        dRateTravel = If(IsDBNull(uRecSnap("RateTravel").Value), "", uRecSnap("RateTravel").Value)
        bELive = uRecSnap("ELive").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadEngineer = True

Err_LoadEngineer:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsEngineer", "LoadEngineer", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadEngineer" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadEngineer = False
        End If

    End Function

    Public Function SaveEngineer() As Boolean
        '** Save Engineer Data Record

        ' Error Checking
        On Error GoTo Err_SaveEngineer

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iEngineerID = 0 Then iEngineerID = GetEngineerID()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Client Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            uPar = .CreateParameter("@Title", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20)
            .Parameters.Append(uPar)
            .Parameters("@Title").Value = sTitle
            uPar = .CreateParameter("@Forename", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Forename").Value = sForename
            uPar = .CreateParameter("@Surname", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Surname").Value = sSurname
            uPar = .CreateParameter("@UserName", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20)
            .Parameters.Append(uPar)
            .Parameters("@UserName").Value = UserName
            uPar = .CreateParameter("@MobileNo", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 14)
            .Parameters.Append(uPar)
            .Parameters("@MobileNo").Value = sMobileNo
            uPar = .CreateParameter("@EMailAddress", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@EMailAddress").Value = sEMailAddress
            uPar = .CreateParameter("@RateNormal", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RateNormal").Value = dRateNormal
            uPar = .CreateParameter("@RateOvertime", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RateOvertime").Value = dRateOvertime
            uPar = .CreateParameter("@RateDouble", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RateDouble").Value = dRateDouble
            uPar = .CreateParameter("@RateTravel", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RateTravel").Value = dRateTravel
            uPar = .CreateParameter("@ELive", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ELive").Value = -bELive
            .CommandText = "Engineer_SaveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveEngineer = True

Err_SaveEngineer:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsEngineer", "SaveEngineer", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveEngineer" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveEngineer = False
        End If

    End Function


    Public Function RemoveItem() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveItem

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Client Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            .CommandText = "Engineer_RemoveItem"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveItem = True

Err_RemoveItem:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsEngineer", "RemoveItem", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveItem" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveItem = False
        End If

    End Function

    Public Function GetEngineerID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetEngineerID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Client Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            .CommandText = "Engineer_GetID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Engineer Class Object", sErrDescription)
            GetEngineerID = 0
            Exit Function
        End If

        ' Store Data Values
        GetEngineerID = uRecSnap("NewEngineerID").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetEngineerID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsEngineer", "GetEngineerID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetEngineerID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetEngineerID = 0
        End If

    End Function


End Class
