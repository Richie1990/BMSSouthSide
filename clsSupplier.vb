Public Class clsSupplier

    ' Dimension Class Variables
    Private iSupplierID As Integer
    Private iContactID As Integer
    Private sTitle As String
    Private sForename As String
    Private sSurname As String
    Private sCompany As String
    Private sAddress1 As String
    Private sAddress2 As String
    Private sAddress3 As String
    Private sAddress4 As String
    Private sPostCode As String
    Private sCompanyPhoneNo As String
    Private sPhoneNo As String
    Private sMobileNo As String
    Private sEMailAddress As String
    Private sWebsite As String
    Private sSupplierNotes As String
    Private iSourceID As Integer
    Private sSourceDescription As String
    Private bCLive As Boolean


    fgvbbbt

    Public Property SupplierID() As Integer
        '** Expose Supplier ID
        Get
            SupplierID = iSupplierID
        End Get
        Set(ByVal Value As Integer)
            iSupplierID = Value
        End Set

    End Property

    Public Property ContactID() As Integer
        '** Expose Contact ID
        Get
            ContactID = iContactID
        End Get
        Set(ByVal Value As Integer)
            iContactID = Value
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

    Public Property Company() As String
        '** Expose Company
        Get
            Company = sCompany
        End Get
        Set(ByVal Value As String)
            sCompany = Value
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

    Public Property Address4() As String
        '** Expose Address Details
        Get
            Address4 = sAddress4
        End Get
        Set(ByVal Value As String)
            sAddress4 = Value
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

    Public Property CompanyPhoneNo() As String
        '** Expose Company Phone No
        Get
            CompanyPhoneNo = sCompanyPhoneNo
        End Get
        Set(ByVal Value As String)
            sCompanyPhoneNo = Value
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

    Public Property Website() As String
        '** Expose Website
        Get
            Website = sWebsite
        End Get
        Set(ByVal Value As String)
            sWebsite = Value
        End Set

    End Property

    Public Property SupplierNotes() As String
        '** Expose Supplier Notes
        Get
            SupplierNotes = sSupplierNotes
        End Get
        Set(ByVal Value As String)
            sSupplierNotes = Value
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

    Public Property SourceDescription() As String
        '** Expose Source Description
        Get
            SourceDescription = sSourceDescription
        End Get
        Set(ByVal Value As String)
            sSourceDescription = Value
        End Set

    End Property

    Public Property CLive() As Boolean
        '** Expose Live Status
        Get
            CLive = bCLive
        End Get
        Set(ByVal Value As Boolean)
            bCLive = Value
        End Set

    End Property



    Public Function LoadSupplier() As Boolean
        '**Finds the Current Supplier Data
        '**Raises an Error if Supplier not found

        ' Error Checking
        On Error GoTo Err_LoadSupplier

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Supplier Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            .CommandText = "Supplier_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Check for Valid Supplier - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Supplier Class Object", sErrDescription)
            LoadSupplier = False
            Exit Function
        End If

        ' Store Data Values
        sCompany = If(IsDBNull(uRecSnap("Company").Value), "", uRecSnap("Company").Value)
        sAddress1 = If(IsDBNull(uRecSnap("Address1").Value), "", uRecSnap("Address1").Value)
        sAddress2 = If(IsDBNull(uRecSnap("Address2").Value), "", uRecSnap("Address2").Value)
        sAddress3 = If(IsDBNull(uRecSnap("Address3").Value), "", uRecSnap("Address3").Value)
        sAddress4 = If(IsDBNull(uRecSnap("Address4").Value), "", uRecSnap("Address4").Value)
        sPostCode = If(IsDBNull(uRecSnap("PostCode").Value), "", uRecSnap("PostCode").Value)
        sWebsite = If(IsDBNull(uRecSnap("Website").Value), "", uRecSnap("Website").Value)
        sWebsite = If(IsDBNull(uRecSnap("Website").Value), "", uRecSnap("Website").Value)
        sSupplierNotes = If(IsDBNull(uRecSnap("SupplierNotes").Value), "", uRecSnap("SupplierNotes").Value)
        bCLive = uRecSnap("CLive").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadSupplier = True

Err_LoadSupplier:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "LoadSupplier", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadSupplier" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadSupplier = False
        End If

    End Function

    Public Function LoadSupplierContact() As Boolean
        '**Finds the Current Supplier ContactData
        '**Raises an Error if Supplier not found

        ' Error Checking
        On Error GoTo Err_LoadSupplierContact

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Supplier Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            uPar = .CreateParameter("@ContactID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ContactID").Value = iContactID
            .CommandText = "SupplierContact_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Check for Valid Supplier - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Supplier Class Object", sErrDescription)
            LoadSupplierContact = False
            Exit Function
        End If

        ' Store Data Values
        sTitle = If(IsDBNull(uRecSnap("Title").Value), "", uRecSnap("Title").Value)
        sForename = If(IsDBNull(uRecSnap("Forename").Value), "", uRecSnap("Forename").Value)
        sSurname = If(IsDBNull(uRecSnap("Surname").Value), "", uRecSnap("Surname").Value)
        sPhoneNo = If(IsDBNull(uRecSnap("PhoneNo").Value), "", uRecSnap("PhoneNo").Value)
        sMobileNo = If(IsDBNull(uRecSnap("MobileNo").Value), "", uRecSnap("MobileNo").Value)
        sEMailAddress = If(IsDBNull(uRecSnap("EMailAddress").Value), "", uRecSnap("EMailAddress").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadSupplierContact = True

Err_LoadSupplierContact:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "LoadSupplierContact", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadSupplierContact" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadSupplierContact = False
        End If

    End Function

    Public Function SaveSupplier() As Boolean
        '** Save Current  Data Record

        ' Error Checking
        On Error GoTo Err_SaveSupplier

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iSupplierID = 0 Then iSupplierID = GetSupplierID()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Supplier Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            uPar = .CreateParameter("@Company", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40)
            .Parameters.Append(uPar)
            .Parameters("@Company").Value = sCompany
            uPar = .CreateParameter("@Address1", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Address1").Value = sAddress1
            uPar = .CreateParameter("@Address2", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Address2").Value = sAddress2
            uPar = .CreateParameter("@Address3", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Address3").Value = sAddress3
            uPar = .CreateParameter("@Address4", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Address4").Value = sAddress4
            uPar = .CreateParameter("@PostCode", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8)
            .Parameters.Append(uPar)
            .Parameters("@PostCode").Value = sPostCode
            uPar = .CreateParameter("@Website", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@Website").Value = sWebsite
            uPar = .CreateParameter("@SupplierNotes", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000)
            .Parameters.Append(uPar)
            .Parameters("@SupplierNotes").Value = sSupplierNotes
            uPar = .CreateParameter("@CLive", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@CLive").Value = -bCLive
            .CommandText = "Supplier_SaveRecord"
            .Execute()
        End With
        iSupplierID = uCommand.Parameters(0).Value

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveSupplier = True

Err_SaveSupplier:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "SaveSupplier", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveSupplier" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveSupplier = False
        End If

    End Function

    Public Function SaveSupplierContact() As Boolean
        '** Save Current  Data Record

        ' Error Checking
        On Error GoTo Err_SaveSupplierContact

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iContactID = 0 Then iContactID = GetSupplierContactID()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Supplier Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            uPar = .CreateParameter("@ContactID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ContactID").Value = iContactID
            uPar = .CreateParameter("@Title", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20)
            .Parameters.Append(uPar)
            .Parameters("@Title").Value = sTitle
            uPar = .CreateParameter("@Forename", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Forename").Value = sForename
            uPar = .CreateParameter("@Surname", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Surname").Value = sSurname
            uPar = .CreateParameter("@PhoneNo", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 14)
            .Parameters.Append(uPar)
            .Parameters("@PhoneNo").Value = sPhoneNo
            uPar = .CreateParameter("@MobileNo", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 14)
            .Parameters.Append(uPar)
            .Parameters("@MobileNo").Value = sMobileNo
            uPar = .CreateParameter("@EMailAddress", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@EMailAddress").Value = sEMailAddress
            .CommandText = "SupplierContact_SaveRecord"
            .Execute()
        End With
        iContactID = uCommand.Parameters(0).Value

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveSupplierContact = True

Err_SaveSupplierContact:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "SaveSupplierContact", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveSupplierContact" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveSupplierContact = False
        End If

    End Function

    Public Function SupplierRemoveItem() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_SupplierRemoveItem

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Supplier Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            .CommandText = "Supplier_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SupplierRemoveItem = True

Err_SupplierRemoveItem:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "SupplierRemoveItem", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SupplierRemoveItem" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SupplierRemoveItem = False
        End If

    End Function

    Public Function SupplierContactRemoveItem() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_SupplierContactRemoveItem

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Supplier Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            uPar = .CreateParameter("@ContactID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ContactID").Value = iContactID
            .CommandText = "SupplierContact_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SupplierContactRemoveItem = True

Err_SupplierContactRemoveItem:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "SupplierContactRemoveItem", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SupplierContactRemoveItem" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SupplierContactRemoveItem = False
        End If

    End Function

    Public Function GetSupplierID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetSupplierID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Supplier Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            .CommandText = "Supplier_GetID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Supplier Class Object", sErrDescription)
            GetSupplierID = 0
            Exit Function
        End If

        ' Store Data Values
        GetSupplierID = uRecSnap("NewSupplierID").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetSupplierID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "GetSupplierID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetSupplierID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetSupplierID = 0
        End If

    End Function

    Public Function GetSupplierContactID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetSupplierContactID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Supplier Contact ID
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            .CommandText = "SupplierContact_GetID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Supplier Class Object", sErrDescription)
            GetSupplierContactID = 0
            Exit Function
        End If

        ' Store Data Values
        GetSupplierContactID = uRecSnap("NewSupplierContactID").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetSupplierContactID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "GetSupplierContactID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetSupplierContactID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetSupplierContactID = 0
        End If

    End Function


    Public Function LoadSupplierJobs() As ADODB.Recordset

        '** Load Supplier Jobs
        ' Error Checking
        On Error GoTo Err_LoadSupplierJobs

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Supplier Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            .CommandText = "Supplier_LoadJobList"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Supplier Class Object", sErrDescription)
            LoadSupplierJobs = Nothing
            Exit Function
        End If

        ' Store Data Values
        LoadSupplierJobs = uRecSnap

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_LoadSupplierJobs:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsSupplier", "LoadSupplierJobs", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadSupplierJobs" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadSupplierJobs = Nothing
        End If

    End Function

End Class
