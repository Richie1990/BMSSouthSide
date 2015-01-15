Public Class clsLandlord

    ' Dimension Class Variables - Company
    Private iLandlordID As Integer
    Private sTitle1 As String
    Private sForename1 As String
    Private sSurname1 As String
    Private dDOB1 As Date
    Private sPlaceofBirth1 As String
    Private sTitle2 As String
    Private sForename2 As String
    Private sSurname2 As String
    Private dDOB2 As Date
    Private sPlaceofBirth2 As String
    Private sCommsTitle As String
    Private sLandlordName As String
    Private sAddress1 As String
    Private sAddress2 As String
    Private sAddress3 As String
    Private sTown As String
    Private sPostCode As String
    Private sPhoneNo As String
    Private sMobileNo As String
    Private sEMailAddress As String
    Private sCommsStatus As String
    Private sLandlordNotes As String
    Private iSourceID As Integer
    Private sSourceDescription As String
    Private dJoinDate As Date
    Private sBankName As String
    Private iSortcode As Integer
    Private iAccountNo As Integer
    Private bLStatus As Boolean




    dfbdfvb

    Public Property LandlordID() As Integer
        '** Expose Landlord ID
        Get
            LandlordID = iLandlordID
        End Get
        Set(ByVal Value As Integer)
            iLandlordID = Value
        End Set

    End Property

    Public Property Title1() As String
        '** Expose Title1
        Get
            Title1 = sTitle1
        End Get
        Set(ByVal Value As String)
            sTitle1 = Value
        End Set

    End Property

    Public Property Forename1() As String
        '** Expose Forename1
        Get
            Forename1 = sForename1
        End Get
        Set(ByVal Value As String)
            sForename1 = Value
        End Set

    End Property

    Public Property Surname1() As String
        '** Expose Surname1
        Get
            Surname1 = sSurname1
        End Get
        Set(ByVal Value As String)
            sSurname1 = Value
        End Set

    End Property

    Public Property DOB1() As Date
        '** Expose DOB1
        Get
            DOB1 = dDOB1
        End Get
        Set(ByVal Value As Date)
            dDOB1 = Value
        End Set

    End Property

    Public Property PlaceofBirth1() As String
        '** Expose PlaceofBirth1
        Get
            PlaceofBirth1 = sPlaceofBirth1
        End Get
        Set(ByVal Value As String)
            sPlaceofBirth1 = Value
        End Set

    End Property


    Public Property Title2() As String
        '** Expose Title2
        Get
            Title2 = sTitle2
        End Get
        Set(ByVal Value As String)
            sTitle2 = Value
        End Set

    End Property

    Public Property Forename2() As String
        '** Expose Forename2
        Get
            Forename2 = sForename2
        End Get
        Set(ByVal Value As String)
            sForename2 = Value
        End Set

    End Property

    Public Property Surname2() As String
        '** Expose Surname2
        Get
            Surname2 = sSurname2
        End Get
        Set(ByVal Value As String)
            sSurname2 = Value
        End Set

    End Property

    Public Property DOB2() As Date
        '** Expose DOB2
        Get
            DOB2 = dDOB2
        End Get
        Set(ByVal Value As Date)
            dDOB2 = Value
        End Set

    End Property

    Public Property PlaceofBirth2() As String
        '** Expose PlaceofBirth2
        Get
            PlaceofBirth2 = sPlaceofBirth2
        End Get
        Set(ByVal Value As String)
            sPlaceofBirth2 = Value
        End Set

    End Property

    Public Property CommsTitle() As String
        '** Expose CommsTitle
        Get
            CommsTitle = sCommsTitle
        End Get
        Set(ByVal Value As String)
            sCommsTitle = Value
        End Set

    End Property

    Public Property LandlordName() As String
        '** Expose LandlordName
        Get
            LandlordName = sLandlordName
        End Get
        Set(ByVal Value As String)
            sLandlordName = Value
        End Set

    End Property

    Public Property Address1() As String
        '** Expose Address Details
        Get
            Address1 = sAddress1
        End Get
        Set(ByVal Value As String)
            sAddress1 = Value
        End Set

    End Property

    Public Property Address2() As String
        '** Expose Address Details
        Get
            Address2 = sAddress2
        End Get
        Set(ByVal Value As String)
            sAddress2 = Value
        End Set

    End Property

    Public Property Address3() As String
        '** Expose Address Details
        Get
            Address3 = sAddress3
        End Get
        Set(ByVal Value As String)
            sAddress3 = Value
        End Set

    End Property

    Public Property Town() As String
        '** Expose Address Details
        Get
            Town = sTown
        End Get
        Set(ByVal Value As String)
            sTown = Value
        End Set

    End Property

    Public Property PostCode() As String
        '** Expose Address Details
        Get
            PostCode = sPostCode
        End Get
        Set(ByVal Value As String)
            sPostCode = Value
        End Set

    End Property

    Public Property PhoneNo() As String
        '** Expose Phone No
        Get
            PhoneNo = sPhoneNo
        End Get
        Set(ByVal Value As String)
            sPhoneNo = Value
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

    Public Property CommsStatus() As String
        '** Expose CommsStatus
        Get
            CommsStatus = sCommsStatus
        End Get
        Set(ByVal Value As String)
            sCommsStatus = Value
        End Set

    End Property

    Public Property LandlordNotes() As String
        '** Expose Landlord Notes
        Get
            LandlordNotes = sLandlordNotes
        End Get
        Set(ByVal Value As String)
            sLandlordNotes = Value
        End Set

    End Property

    Public Property SourceID() As Integer
        '** Expose Source ID
        Get
            SourceID = iSourceID
        End Get
        Set(ByVal Value As Integer)
            iSourceID = Value
        End Set

    End Property

    Public ReadOnly Property SourceDescription() As String
        '** Expose SourceDescription
        Get
            SourceDescription = sSourceDescription
        End Get

    End Property


    Public Property JoinDate() As Date
        '** Expose Source Description
        Get
            JoinDate = dJoinDate
        End Get
        Set(ByVal Value As Date)
            dJoinDate = Value
        End Set

    End Property

    Public Property BankName() As String
        '** Expose BankName
        Get
            BankName = sBankName
        End Get
        Set(ByVal Value As String)
            sBankName = Value
        End Set

    End Property

    Public Property Sortcode() As Integer
        '** Expose Sortcode
        Get
            Sortcode = iSortcode
        End Get
        Set(ByVal Value As Integer)
            iSortcode = Value
        End Set

    End Property

    Public Property AccountNo() As Integer
        '** Expose AccountNo
        Get
            AccountNo = iAccountNo
        End Get
        Set(ByVal Value As Integer)
            iAccountNo = Value
        End Set

    End Property


    Public Property LStatus() As Boolean
        '** Expose Live Status
        Get
            LStatus = bLStatus
        End Get
        Set(ByVal Value As Boolean)
            bLStatus = Value
        End Set

    End Property


    Public Function LoadLandlord() As Boolean
        '**Finds the Current Landlord Data

        ' Error Checking
        On Error GoTo Err_LoadLandlord

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Landlord Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@LandlordID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LandlordID").Value = iLandlordID
            .CommandText = "Landlord_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        Do Until uRecSnap.EOF

            sTitle1 = If(IsDBNull(uRecSnap("Title1").Value), "", uRecSnap("Title1").Value)
            sForename1 = If(IsDBNull(uRecSnap("Forename1").Value), "", uRecSnap("Forename1").Value)
            sSurname1 = If(IsDBNull(uRecSnap("Surname1").Value), "", uRecSnap("Surname1").Value)
            dDOB1 = If(IsDBNull(uRecSnap("DOB1").Value), "01/01/1900", uRecSnap("DOB1").Value)
            sPlaceofBirth1 = If(IsDBNull(uRecSnap("PlaceofBirth1").Value), "", uRecSnap("PlaceofBirth1").Value)
            sTitle2 = If(IsDBNull(uRecSnap("Title2").Value), "", uRecSnap("Title2").Value)
            sForename2 = If(IsDBNull(uRecSnap("Forename2").Value), "", uRecSnap("Forename2").Value)
            sSurname2 = If(IsDBNull(uRecSnap("Surname2").Value), "", uRecSnap("Surname2").Value)
            dDOB2 = If(IsDBNull(uRecSnap("DOB2").Value), "02/02/2900", uRecSnap("DOB2").Value)
            sPlaceofBirth2 = If(IsDBNull(uRecSnap("PlaceofBirth2").Value), "", uRecSnap("PlaceofBirth2").Value)
            sCommsTitle = If(IsDBNull(uRecSnap("CommsTitle").Value), "", uRecSnap("CommsTitle").Value)
            sLandlordName = If(IsDBNull(uRecSnap("LandlordName").Value), "", uRecSnap("LandlordName").Value)
            sAddress1 = If(IsDBNull(uRecSnap("Address1").Value), "", uRecSnap("Address1").Value)
            sAddress2 = If(IsDBNull(uRecSnap("Address2").Value), "", uRecSnap("Address2").Value)
            sAddress3 = If(IsDBNull(uRecSnap("Address3").Value), "", uRecSnap("Address3").Value)
            sTown = If(IsDBNull(uRecSnap("Town").Value), "", uRecSnap("Town").Value)
            sPostCode = If(IsDBNull(uRecSnap("PostCode").Value), "", uRecSnap("PostCode").Value)
            sPhoneNo = If(IsDBNull(uRecSnap("PhoneNo").Value), "", uRecSnap("PhoneNo").Value)
            sMobileNo = If(IsDBNull(uRecSnap("MobileNo").Value), "", uRecSnap("MobileNo").Value)
            sEMailAddress = If(IsDBNull(uRecSnap("EMailAddress").Value), "", uRecSnap("EMailAddress").Value)
            sCommsStatus = If(IsDBNull(uRecSnap("CommsStatus").Value), "", uRecSnap("CommsStatus").Value)
            sLandlordNotes = If(IsDBNull(uRecSnap("LandlordNotes").Value), "", uRecSnap("LandlordNotes").Value)
            iSourceID = If(IsDBNull(uRecSnap("SourceID").Value), 0, uRecSnap("SourceID").Value)
            dJoinDate = If(IsDBNull(uRecSnap("JoinDate").Value), "", uRecSnap("JoinDate").Value)
            sBankName = If(IsDBNull(uRecSnap("BankName").Value), "", uRecSnap("BankName").Value)
            iSortcode = If(IsDBNull(uRecSnap("Sortcode").Value), 0, uRecSnap("Sortcode").Value)
            iAccountNo = If(IsDBNull(uRecSnap("AccountNo").Value), 0, uRecSnap("AccountNo").Value)
            bLStatus = If(IsDBNull(uRecSnap("LStatus").Value), 0, uRecSnap("LStatus").Value)
            uRecSnap.MoveNext()
        Loop

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_LoadLandlord:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsLandlord", "LoadLandlord", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadLandlord" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Function

 

    Public Function SaveLandlord() As Boolean
        '** Save Current Personal Data Record

        ' Error Checking
        On Error GoTo Err_SaveLandlord

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Landlord Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@LandlordID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LandlordID").Value = iLandlordID
            uPar = .CreateParameter("@Title1", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20)
            .Parameters.Append(uPar)
            .Parameters("@Title1").Value = sTitle1
            uPar = .CreateParameter("@Forename1", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Forename1").Value = sForename1
            uPar = .CreateParameter("@Surname1", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Surname1").Value = sSurname1
            uPar = .CreateParameter("@DOB1", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@DOB1").Value = dDOB1
            uPar = .CreateParameter("@PlaceofBirth1", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@PlaceofBirth1").Value = PlaceofBirth1
            uPar = .CreateParameter("@Title2", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20)
            .Parameters.Append(uPar)
            .Parameters("@Title2").Value = sTitle2
            uPar = .CreateParameter("@Forename2", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Forename2").Value = sForename2
            uPar = .CreateParameter("@Surname2", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Surname2").Value = sSurname2
            uPar = .CreateParameter("@DOB2", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@DOB2").Value = dDOB2
            uPar = .CreateParameter("@PlaceofBirth2", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@PlaceofBirth2").Value = PlaceofBirth2
            uPar = .CreateParameter("@CommsTitle", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@CommsTitle").Value = sCommsTitle
            uPar = .CreateParameter("@LandlordName", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@LandlordName").Value = sLandlordName
            uPar = .CreateParameter("@Address1", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Address1").Value = sAddress1
            uPar = .CreateParameter("@Address2", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Address2").Value = sAddress2
            uPar = .CreateParameter("@Address3", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Address3").Value = sAddress3
            uPar = .CreateParameter("@Town", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Town").Value = sTown
            uPar = .CreateParameter("@PostCode", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8)
            .Parameters.Append(uPar)
            .Parameters("@PostCode").Value = sPostCode
            uPar = .CreateParameter("@PhoneNo", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 18)
            .Parameters.Append(uPar)
            .Parameters("@PhoneNo").Value = sPhoneNo
            uPar = .CreateParameter("@MobileNo", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 18)
            .Parameters.Append(uPar)
            .Parameters("@MobileNo").Value = sMobileNo
            uPar = .CreateParameter("@EMailAddress", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@EMailAddress").Value = sEMailAddress
            uPar = .CreateParameter("@CommsStatus", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20)
            .Parameters.Append(uPar)
            .Parameters("@CommsStatus").Value = sCommsStatus
            uPar = .CreateParameter("@LandlordNotes", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000)
            .Parameters.Append(uPar)
            .Parameters("@LandlordNotes").Value = sLandlordNotes
            uPar = .CreateParameter("@SourceID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SourceID").Value = iSourceID
            uPar = .CreateParameter("@JoinDate", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JoinDate").Value = dJoinDate
            uPar = .CreateParameter("@BankName", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40)
            .Parameters.Append(uPar)
            .Parameters("@BankName").Value = sBankName
            uPar = .CreateParameter("@Sortcode", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Sortcode").Value = iSortcode
            uPar = .CreateParameter("@AccountNo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@AccountNo").Value = iAccountNo
            uPar = .CreateParameter("@LStatus", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LStatus").Value = -bLStatus
            .CommandText = "Landlord_ImportRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveLandlord = True

Err_SaveLandlord:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsLandlord", "SaveLandlord", "Error", sErrDescription)
            SaveLandlord = False
        End If

    End Function


    Public Function RemoveLandlord() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveLandlord

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Landlord Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@LandlordID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LandlordID").Value = iLandlordID
            .CommandText = "Landlord_RemoveItem"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveLandlord = True

Err_RemoveLandlord:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsLandlord", "RemoveLandlord", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveLandlord" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveLandlord = False
        End If

    End Function

    Public Function GetLandlordID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetLandlordID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Landlord Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            .CommandText = "Landlord_GetID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Landlord Class Object", sErrDescription)
            GetLandlordID = 0
            Exit Function
        End If

        ' Store Data Values
        GetLandlordID = uRecSnap("NewLandlordID").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetLandlordID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsLandlord", "GetLandlordID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetLandlordID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetLandlordID = 0
        End If

    End Function


    Public Function LoadLandlordProperties() As ADODB.Recordset

        '** Load Landlord Jobs
        ' Error Checking
        On Error GoTo Err_LoadLandlordJobs

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Landlord Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@LandlordID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LandlordID").Value = iLandlordID
            .CommandText = "Landlord_LoadJobList"
            uRecSnap = .Execute
        End With


        ' Store Data Values

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_LoadLandlordJobs:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsLandlord", "LoadLandlordJobs", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadLandlordJobs" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Function

End Class
