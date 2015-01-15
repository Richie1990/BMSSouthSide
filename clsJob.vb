Public Class clsJob

    ' Dimension Class Variables - Job
    Private iJobID As Integer
    Private iClientID As Integer
    Private iContactID As Integer
    Private sClientFullName As String
    Private sClientEMailAddress As String
    Private sClientTelephone As String
    Private sDefaultAddress As String
    Private sCAddress1 As String
    Private sCAddress2 As String
    Private sCAddress3 As String
    Private sCAddress4 As String
    Private sCPostCode As String
    Private sDAddress1 As String
    Private sDAddress2 As String
    Private sDAddress3 As String
    Private sDAddress4 As String
    Private sDPostCode As String
    Private sJobDescription As String
    Private dValue As Decimal
    Private iEngineerID As Integer
    Private sEngineerDescription As String
    Private iItemTypeID As Integer
    Private sItemTypeDescription As String
    Private iNoItems As Integer
    Private dCreateDate As Date
    Private dSaveDate As Date
    Private sNotes As String
    Private iStatusID As Integer
    Private sStatusDescription As String
    Private bPrepaid As Boolean
    Private bCash As Boolean
    Private bXeroInvoiceCreated As Boolean
    Private dValuationAmount As Double
    Private dInvoicedAmount As Double
    Private dReceivedAmount As Double
    Private dLabourCost As Double
    Private dLabourExpense As Double
    Private dMaterialCost As Double

    ' Dimension Class Variables - Job Item 
    Private iItemNo As Integer
    Private sMake As String
    Private sModel As String
    Private sSerialNo As String
    Private sServiceTag As String
    Private sPassword As String
    Private sDescription As String

    ' Dimension Class Variables - Job Item 
    Private iOfferID As Integer
    Private sOfferDescription As String
    Private dExpiryDate As Date

    ' Dimension Class Variables - Job Product
    Private iProductID As Integer
    Private sProductDescription As String
    Private iRowNo As Integer
    Private dQuantity As Integer
    Private dPrice As Double
    Private dActualPrice As Double
    Private dMarkupPer As Double
    Private dLabourHours As Double
    Private dLabourRate As Double
    Private sFullDescription As String
    Private bShowOnQuote As Boolean
    Private iOrderID As Integer

    ' Dimension Job Engineer Time
    Private iWeekNo As Integer
    Private iTaskNo As Integer
    Private dTaskDateTime As Date
    Private dHoursNormal As Double
    Private dHoursOvertime As Double
    Private dHoursDouble As Double
    Private dHoursTravel As Double
    Private dExpenses As Double
    Private dSubsistence As Double
    Private iMileage As Integer
    Private sBriefNotes As String

    ' Dimension Local Variables - Job Totals
    Private dJobTotalMarkup As Double
    Private dJobTotalMaterials As Double
    Private dJobTotalHours As Double
    Private dJobTotalCost As Double
    Private dJobTotalActualPrice As Double

    ' Dimension Local Variables - General Statistics
    Private iMonth As Integer
    Private iWeek As Integer
    Private iDay As Integer
    Private iUnallocatedNo As Integer
    Private iInProgressNo As Integer
    Private dInProgressHours As Double
    Private iCompletedNo As Integer
    Private dCompletedHours As Double
    Private iContractsNo As Integer
    Private dContractsHours As Double
    Private dContractsTravelHours As Double

    Public Property JobID() As Integer
        '** Expose Job ID
        Get
            JobID = iJobID
        End Get
        Set(ByVal Value As Integer)
            iJobID = Value
        End Set

    End Property

    Public Property ClientID() As Integer
        '** Expose Client ID
        Get
            ClientID = iClientID
        End Get
        Set(ByVal Value As Integer)
            iClientID = Value
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

    Public ReadOnly Property ClientFullName() As String
        '** Expose Client FullName
        Get
            ClientFullName = sClientFullName
        End Get

    End Property

    Public ReadOnly Property ClientEMailAddress() As String
        '** Expose Client EMailAddress
        Get
            ClientEMailAddress = sClientEMailAddress
        End Get

    End Property

    Public ReadOnly Property ClientTelephone() As String
        '** Expose Client Telephone
        Get
            ClientTelephone = sClientTelephone
        End Get

    End Property

    Public Property DefaultAddress() As String
        '** Expose DefaultAddress
        Get
            DefaultAddress = sDefaultAddress
        End Get
        Set(ByVal Value As String)
            sDefaultAddress = Value
        End Set

    End Property

    Public Property CAddress1() As String
        '** Expose Address Details
        Get
            CAddress1 = sCAddress1
        End Get
        Set(ByVal Value As String)
            sCAddress1 = Value
        End Set

    End Property

    Public Property CAddress2() As String
        '** Expose Address Details
        Get
            CAddress2 = sCAddress2
        End Get
        Set(ByVal Value As String)
            sCAddress2 = Value
        End Set

    End Property

    Public Property CAddress3() As String
        '** Expose Address Details
        Get
            CAddress3 = sCAddress3
        End Get
        Set(ByVal Value As String)
            sCAddress3 = Value
        End Set

    End Property

    Public Property CAddress4() As String
        '** Expose Address Details
        Get
            CAddress4 = sCAddress4
        End Get
        Set(ByVal Value As String)
            sCAddress4 = Value
        End Set

    End Property

    Public Property CPostCode() As String
        '** Expose Address Details
        Get
            CPostCode = sCPostCode
        End Get
        Set(ByVal Value As String)
            sCPostCode = Value
        End Set

    End Property

    Public Property DAddress1() As String
        '** Expose Address Details
        Get
            DAddress1 = sDAddress1
        End Get
        Set(ByVal Value As String)
            sDAddress1 = Value
        End Set

    End Property

    Public Property DAddress2() As String
        '** Expose Address Details
        Get
            DAddress2 = sDAddress2
        End Get
        Set(ByVal Value As String)
            sDAddress2 = Value
        End Set

    End Property

    Public Property DAddress3() As String
        '** Expose Address Details
        Get
            DAddress3 = sDAddress3
        End Get
        Set(ByVal Value As String)
            sDAddress3 = Value
        End Set

    End Property

    Public Property DAddress4() As String
        '** Expose Address Details
        Get
            DAddress4 = sDAddress4
        End Get
        Set(ByVal Value As String)
            sDAddress4 = Value
        End Set

    End Property

    Public Property DPostCode() As String
        '** Expose Address Details
        Get
            DPostCode = sDPostCode
        End Get
        Set(ByVal Value As String)
            sDPostCode = Value
        End Set

    End Property

    Public Property JobDescription() As String
        '** Expose Job Description
        Get
            JobDescription = sJobDescription
        End Get
        Set(ByVal Value As String)
            sJobDescription = Value
        End Set

    End Property

    Public Property Value() As Decimal
        '** Expose Value
        Get
            Value = dValue
        End Get
        Set(ByVal Value As Decimal)
            dValue = Value
        End Set

    End Property

    Public Property EngineerID() As Integer
        '** Expose Engineer ID
        Get
            EngineerID = iEngineerID
        End Get
        Set(ByVal Value As Integer)
            iEngineerID = Value
        End Set

    End Property

    Public ReadOnly Property EngineerDescription() As String
        '** Expose Engineer Description
        Get
            EngineerDescription = sEngineerDescription
        End Get

    End Property

    Public Property NoItems() As Integer
        '** Expose NoItems 
        Get
            NoItems = iNoItems
        End Get
        Set(ByVal Value As Integer)
            iNoItems = Value
        End Set

    End Property

    Public Property CreateDate() As Date
        '** Expose CreateDate
        Get
            CreateDate = dCreateDate
        End Get
        Set(ByVal Value As Date)
            dCreateDate = Value
        End Set

    End Property

    Public Property SaveDate() As Date
        '** Expose Save Date
        Get
            SaveDate = dSaveDate
        End Get
        Set(ByVal Value As Date)
            dSaveDate = Value
        End Set

    End Property

    Public Property Notes() As String
        '** Expose Notes
        Get
            Notes = sNotes
        End Get
        Set(ByVal Value As String)
            sNotes = Value
        End Set

    End Property

    Public Property StatusID() As Integer
        '** Expose StatusID
        Get
            StatusID = iStatusID
        End Get
        Set(ByVal Value As Integer)
            iStatusID = Value
        End Set

    End Property

    Public ReadOnly Property StatusDescription() As String
        '** Expose Status Description
        Get
            StatusDescription = sStatusDescription
        End Get

    End Property

    Public Property Prepaid() As Boolean
        '** Expose Prepaid
        Get
            Prepaid = bPrepaid
        End Get
        Set(ByVal Value As Boolean)
            bPrepaid = Value
        End Set

    End Property

    Public Property Cash() As Boolean
        '** Expose Cash
        Get
            Cash = bCash
        End Get
        Set(ByVal Value As Boolean)
            bCash = Value
        End Set

    End Property

    Public ReadOnly Property XeroInvoiceCreated() As Boolean
        '** Expose XeroInvoiceCreated
        Get
            XeroInvoiceCreated = bXeroInvoiceCreated
        End Get

    End Property

    Public ReadOnly Property ValuationAmount() As Double
        '** Expose Valuation Amount
        Get
            ValuationAmount = dValuationAmount
        End Get

    End Property

    Public ReadOnly Property InvoicedAmount() As Double
        '** Expose Invoiced Amount
        Get
            InvoicedAmount = dInvoicedAmount
        End Get

    End Property

    Public ReadOnly Property ReceivedAmount() As Double
        '** Expose Received Amount
        Get
            ReceivedAmount = dReceivedAmount
        End Get

    End Property

    Public ReadOnly Property LabourCost() As Double
        '** Expose Labour Cost
        Get
            LabourCost = dLabourCost
        End Get

    End Property

    Public ReadOnly Property LabourExpense() As Double
        '** Expose Labour Expense
        Get
            LabourExpense = dLabourExpense
        End Get

    End Property

    Public ReadOnly Property MaterialCost() As Double
        '** Expose MaterialCost
        Get
            MaterialCost = dMaterialCost
        End Get

    End Property

    Public Property ItemNo() As Integer
        '** Expose ItemNo
        Get
            ItemNo = iItemNo
        End Get
        Set(ByVal Value As Integer)
            iItemNo = Value
        End Set

    End Property

    Public Property ItemTypeID() As Integer
        '** Expose ItemTypeID
        Get
            ItemTypeID = iItemTypeID
        End Get
        Set(ByVal Value As Integer)
            iItemTypeID = Value
        End Set

    End Property

    Public ReadOnly Property ItemTypeDescription() As String
        '** Expose ItemType Description
        Get
            ItemTypeDescription = sItemTypeDescription
        End Get

    End Property

    Public Property Make() As String
        '** Expose Make
        Get
            Make = sMake
        End Get
        Set(ByVal Value As String)
            sMake = Value
        End Set

    End Property

    Public Property Model() As String
        '** Expose Model
        Get
            Model = sModel
        End Get
        Set(ByVal Value As String)
            sModel = Value
        End Set

    End Property

    Public Property SerialNo() As String
        '** Expose SerialNo
        Get
            SerialNo = sSerialNo
        End Get
        Set(ByVal Value As String)
            sSerialNo = Value
        End Set

    End Property

    Public Property ServiceTag() As String
        '** Expose Service Tag
        Get
            ServiceTag = sServiceTag
        End Get
        Set(ByVal Value As String)
            sServiceTag = Value
        End Set

    End Property

    Public Property Password() As String
        '** Expose Password
        Get
            Password = sPassword
        End Get
        Set(ByVal Value As String)
            sPassword = Value
        End Set

    End Property

    Public Property Description() As String
        '** Expose Description
        Get
            Description = sDescription
        End Get
        Set(ByVal Value As String)
            sDescription = Value
        End Set

    End Property

    Public Property OfferID() As Integer
        '** Expose OfferID
        Get
            OfferID = iOfferID
        End Get
        Set(ByVal Value As Integer)
            iOfferID = Value
        End Set

    End Property

    Public Property OfferDescription() As String
        '** Expose Offer Description
        Get
            OfferDescription = sOfferDescription
        End Get
        Set(ByVal Value As String)
            sOfferDescription = Value
        End Set

    End Property

    Public Property ExpiryDate() As Date
        '** Expose Expiry Date
        Get
            ExpiryDate = dExpiryDate
        End Get
        Set(ByVal Value As Date)
            dExpiryDate = Value
        End Set

    End Property


    Public Property ProductID() As Integer
        '** Expose ProductID
        Get
            ProductID = iProductID
        End Get
        Set(ByVal Value As Integer)
            iProductID = Value
        End Set

    End Property

    Public ReadOnly Property ProductDescription() As String
        '** Expose Product Description
        Get
            ProductDescription = sProductDescription
        End Get

    End Property

    Public Property RowNo() As Integer
        '** Expose RowNo
        Get
            RowNo = iRowNo
        End Get
        Set(ByVal Value As Integer)
            iRowNo = Value
        End Set

    End Property

    Public Property Quantity() As Double
        '** Expose Quantity
        Get
            Quantity = dQuantity
        End Get
        Set(ByVal Value As Double)
            dQuantity = Value
        End Set

    End Property

    Public Property Price() As Double
        '** Expose Price
        Get
            Price = dPrice
        End Get
        Set(ByVal Value As Double)
            dPrice = Value
        End Set

    End Property

    Public Property ActualPrice() As Double
        '** Expose Actual Price
        Get
            ActualPrice = dActualPrice
        End Get
        Set(ByVal Value As Double)
            dActualPrice = Value
        End Set

    End Property

    Public Property MarkupPer() As Double
        '** Expose MarkupPer
        Get
            MarkupPer = dMarkupPer
        End Get
        Set(ByVal Value As Double)
            dMarkupPer = Value
        End Set

    End Property

    Public Property LabourHours() As Double
        '** Expose Labour Hours
        Get
            LabourHours = dLabourHours
        End Get
        Set(ByVal Value As Double)
            dLabourHours = Value
        End Set

    End Property

    Public Property LabourRate() As Double
        '** Expose Labour Cost
        Get
            LabourRate = dLabourRate
        End Get
        Set(ByVal Value As Double)
            dLabourRate = Value
        End Set

    End Property

    Public Property FullDescription() As String
        '** Expose Full Description
        Get
            FullDescription = sFullDescription
        End Get
        Set(ByVal Value As String)
            sFullDescription = Value
        End Set

    End Property

    Public Property ShowOnQuote() As Boolean
        '** Expose ShowOnQuote
        Get
            ShowOnQuote = bShowOnQuote
        End Get
        Set(ByVal Value As Boolean)
            bShowOnQuote = Value
        End Set

    End Property

    Public Property OrderID() As Integer
        '** Expose OrderID
        Get
            OrderID = iOrderID
        End Get
        Set(ByVal Value As Integer)
            iOrderID = Value
        End Set

    End Property

    Public Property WeekNo() As Integer
        '** Expose Week No
        Get
            WeekNo = iWeekNo
        End Get
        Set(ByVal Value As Integer)
            iWeekNo = Value
        End Set

    End Property

    Public Property TaskNo() As Integer
        '** Expose Task No
        Get
            TaskNo = iTaskNo
        End Get
        Set(ByVal Value As Integer)
            iTaskNo = Value
        End Set

    End Property

    Public Property TaskDateTime() As Date
        '** Expose Task Date Time
        Get
            TaskDateTime = dTaskDateTime
        End Get
        Set(ByVal Value As Date)
            dTaskDateTime = Value
        End Set

    End Property

    Public Property HoursNormal() As Double
        '** Expose Hours Normal
        Get
            HoursNormal = dHoursNormal
        End Get
        Set(ByVal Value As Double)
            dHoursNormal = Value
        End Set

    End Property

    Public Property HoursOvertime() As Double
        '** Expose Hours Overtime
        Get
            HoursOvertime = dHoursOvertime
        End Get
        Set(ByVal Value As Double)
            dHoursOvertime = Value
        End Set

    End Property

    Public Property HoursDouble() As Double
        '** Expose Hours Double
        Get
            HoursDouble = dHoursDouble
        End Get
        Set(ByVal Value As Double)
            dHoursDouble = Value
        End Set

    End Property

    Public Property HoursTravel() As Double
        '** Expose Hours Travel
        Get
            HoursTravel = dHoursTravel
        End Get
        Set(ByVal Value As Double)
            dHoursTravel = Value
        End Set

    End Property

    Public Property Expenses() As Double
        '** Expose Expenses
        Get
            Expenses = dExpenses
        End Get
        Set(ByVal Value As Double)
            dExpenses = Value
        End Set

    End Property

    Public Property Subsistence() As Double
        '** Expose Subsistence
        Get
            Subsistence = dSubsistence
        End Get
        Set(ByVal Value As Double)
            dSubsistence = Value
        End Set

    End Property

    Public Property Mileage() As Double
        '** Expose Mileage
        Get
            Mileage = iMileage
        End Get
        Set(ByVal Value As Double)
            iMileage = Value
        End Set

    End Property

    Public Property BriefNotes() As String
        '** Expose Brief Notes
        Get
            BriefNotes = sBriefNotes
        End Get
        Set(ByVal Value As String)
            sBriefNotes = Value
        End Set

    End Property

    Public ReadOnly Property JobTotalMarkup() As Double
        '** Expose Job Total Markup
        Get
            JobTotalMarkup = dJobTotalMarkup
        End Get

    End Property

    Public ReadOnly Property JobTotalMaterials() As Double
        '** Expose Job Total Materials
        Get
            JobTotalMaterials = dJobTotalMaterials
        End Get

    End Property

    Public ReadOnly Property JobTotalHours() As Double
        '** Expose Job Total Hours
        Get
            JobTotalHours = dJobTotalHours
        End Get

    End Property

    Public ReadOnly Property JobTotalCost() As Double
        '** Expose Job Total Cost
        Get
            JobTotalCost = dJobTotalCost
        End Get

    End Property

    Public ReadOnly Property JobTotalActualPrice() As Double
        '** Expose Job Total Actual Price
        Get
            JobTotalActualPrice = dJobTotalActualPrice
        End Get

    End Property

    Public WriteOnly Property Month() As Integer
        '** Expose Month

        Set(ByVal Value As Integer)
            iMonth = Value
        End Set

    End Property

    Public WriteOnly Property Week() As Integer
        '** Expose Week

        Set(ByVal Value As Integer)
            iWeek = Value
        End Set

    End Property

    Public WriteOnly Property Day() As Integer
        '** Expose Day

        Set(ByVal Value As Integer)
            iDay = Value
        End Set

    End Property

    Public ReadOnly Property UnallocatedNo() As Integer
        '** Expose Unallocated No
        Get
            UnallocatedNo = iUnallocatedNo
        End Get

    End Property

    Public ReadOnly Property InProgressNo() As Integer
        '** Expose InProgress No
        Get
            InProgressNo = iInProgressNo
        End Get

    End Property

    Public ReadOnly Property InProgressHours() As Double
        '** Expose InProgress Hours
        Get
            InProgressHours = dInProgressHours
        End Get

    End Property

    Public ReadOnly Property CompletedNo() As Integer
        '** Expose Completed No
        Get
            CompletedNo = iCompletedNo
        End Get

    End Property

    Public ReadOnly Property CompletedHours() As Double
        '** Expose Completed Hours
        Get
            CompletedHours = dCompletedHours
        End Get

    End Property

    Public ReadOnly Property ContractsNo() As Integer
        '** Expose Contracts No
        Get
            ContractsNo = iContractsNo
        End Get

    End Property

    Public ReadOnly Property ContractsHours() As Double
        '** Expose InProgress Hours
        Get
            ContractsHours = dContractsHours
        End Get

    End Property

    Public ReadOnly Property ContractsTravelHours() As Double
        '** Expose Contracts TravelHours
        Get
            ContractsTravelHours = dContractsTravelHours
        End Get

    End Property

    Public Function LoadJob() As Boolean
        '**Finds the Current Job Data
        '**Raises an Error if Job not found

        ' Error Checking
        On Error GoTo Err_LoadJob

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            .CommandText = "Job_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            LoadJob = False
            Exit Function
        End If

        ' Store Data Values
        iClientID = If(IsDBNull(uRecSnap("ClientID").Value), 0, uRecSnap("ClientID").Value)
        iContactID = If(IsDBNull(uRecSnap("ContactID").Value), 0, uRecSnap("ContactID").Value)
        If iClientID = 0 Then
            sClientFullName = "Unallocated"
            sClientEMailAddress = ""
            sClientTelephone = ""
        Else
            sClientFullName = If(IsDBNull(uRecSnap("ClientName").Value), "", uRecSnap("ClientName").Value)
            sClientEMailAddress = If(IsDBNull(uRecSnap("ClientEMailAddress").Value), "", uRecSnap("ClientEMailAddress").Value)
            sClientTelephone = If(IsDBNull(uRecSnap("ClientTelephone").Value), "", uRecSnap("ClientTelephone").Value)
        End If
        sDefaultAddress = If(IsDBNull(uRecSnap("DefaultAddress").Value), "C", uRecSnap("DefaultAddress").Value)
        sCAddress1 = If(IsDBNull(uRecSnap("CAddress1").Value), "", uRecSnap("CAddress1").Value)
        sCAddress2 = If(IsDBNull(uRecSnap("CAddress2").Value), "", uRecSnap("CAddress2").Value)
        sCAddress3 = If(IsDBNull(uRecSnap("CAddress3").Value), "", uRecSnap("CAddress3").Value)
        sCAddress4 = If(IsDBNull(uRecSnap("CAddress4").Value), "", uRecSnap("CAddress4").Value)
        sCPostCode = If(IsDBNull(uRecSnap("CPostCode").Value), "", uRecSnap("CPostCode").Value)
        sDAddress1 = If(IsDBNull(uRecSnap("DAddress1").Value), "", uRecSnap("DAddress1").Value)
        sDAddress2 = If(IsDBNull(uRecSnap("DAddress2").Value), "", uRecSnap("DAddress2").Value)
        sDAddress3 = If(IsDBNull(uRecSnap("DAddress3").Value), "", uRecSnap("DAddress3").Value)
        sDAddress4 = If(IsDBNull(uRecSnap("DAddress4").Value), "", uRecSnap("DAddress4").Value)
        sDPostCode = If(IsDBNull(uRecSnap("DPostCode").Value), "", uRecSnap("DPostCode").Value)
        sJobDescription = If(IsDBNull(uRecSnap("JobDescription").Value), "", uRecSnap("JobDescription").Value)
        dValue = If(IsDBNull(uRecSnap("Value").Value), 0, uRecSnap("Value").Value)
        iEngineerID = If(IsDBNull(uRecSnap("EngineerID").Value), 0, uRecSnap("EngineerID").Value)
        sEngineerDescription = If(IsDBNull(uRecSnap("EngineerDescription").Value), 0, uRecSnap("EngineerDescription").Value)
        iItemTypeID = If(IsDBNull(uRecSnap("ItemTypeID").Value), 0, uRecSnap("ItemTypeID").Value)
        sItemTypeDescription = If(IsDBNull(uRecSnap("JobItemTypeDescription").Value), 0, uRecSnap("JobItemTypeDescription").Value)
        dCreateDate = If(IsDBNull(uRecSnap("CreateDate").Value), "01/01/1900", uRecSnap("CreateDate").Value)
        dSaveDate = If(IsDBNull(uRecSnap("SaveDate").Value), "01/01/1900", uRecSnap("SaveDate").Value)
        sNotes = If(IsDBNull(uRecSnap("Notes").Value), "", uRecSnap("Notes").Value)
        iStatusID = If(IsDBNull(uRecSnap("StatusID").Value), 0, uRecSnap("StatusID").Value)
        sStatusDescription = If(IsDBNull(uRecSnap("StatusDescription").Value), "", uRecSnap("StatusDescription").Value)
        bPrepaid = If(IsDBNull(uRecSnap("Prepaid").Value), 0, uRecSnap("Prepaid").Value)
        bCash = If(IsDBNull(uRecSnap("Cash").Value), 0, uRecSnap("Cash").Value)
        bXeroInvoiceCreated = If(IsDBNull(uRecSnap("XeroInvoiceCreated").Value), 0, uRecSnap("XeroInvoiceCreated").Value)
        iOfferID = If(IsDBNull(uRecSnap("OfferID").Value), 0, uRecSnap("OfferID").Value)
        sOfferDescription = If(IsDBNull(uRecSnap("OfferDescription").Value), "", uRecSnap("OfferDescription").Value)
        dValuationAmount = If(IsDBNull(uRecSnap("ValuationAmount").Value), 0, uRecSnap("ValuationAmount").Value)
        dInvoicedAmount = If(IsDBNull(uRecSnap("InvoicedAmount").Value), 0, uRecSnap("InvoicedAmount").Value)
        dReceivedAmount = If(IsDBNull(uRecSnap("ReceivedAmount").Value), 0, -uRecSnap("ReceivedAmount").Value)
        dLabourCost = If(IsDBNull(uRecSnap("LabourCost").Value), 0, uRecSnap("LabourCost").Value)
        dLabourExpense = If(IsDBNull(uRecSnap("LabourExpense").Value), 0, uRecSnap("LabourExpense").Value)
        dMaterialCost = If(IsDBNull(uRecSnap("MaterialCost").Value), 0, uRecSnap("MaterialCost").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadJob = True

Err_LoadJob:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "LoadJob", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadJob" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadJob = False
        End If

    End Function

    Public Function SaveJob() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveJob

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iJobID = 0 Then iJobID = GetJobID()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Job Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ClientID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ClientID").Value = iClientID
            uPar = .CreateParameter("@ContactID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ContactID").Value = iContactID
            uPar = .CreateParameter("@JobDescription", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@JobDescription").Value = sJobDescription
            uPar = .CreateParameter("@Value", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Value").Value = dValue
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            uPar = .CreateParameter("@ItemTypeID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ItemTypeID").Value = iItemTypeID
            uPar = .CreateParameter("@Notes", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 2000)
            .Parameters.Append(uPar)
            .Parameters("@Notes").Value = sNotes
            uPar = .CreateParameter("@StatusID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@StatusID").Value = iStatusID
            uPar = .CreateParameter("@Prepaid", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Prepaid").Value = -bPrepaid
            uPar = .CreateParameter("@Cash", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Cash").Value = -bCash
            uPar = .CreateParameter("@OfferID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OfferID").Value = iOfferID
            .CommandText = "Job_SaveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveJob = True

Err_SaveJob:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "SaveJob", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveJob" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveJob = False
        End If

    End Function

    Public Function RemoveJob() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveJob

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Job Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            .CommandText = "Job_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveJob = True

Err_RemoveJob:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "RemoveJob", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveJob" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveJob = False
        End If

    End Function

    Public Function UpdateJobStatus() As Boolean
        '** Update Status

        ' Error Checking
        On Error GoTo Err_UpdateJobStatus

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ClientID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ClientID").Value = iClientID
            uPar = .CreateParameter("@ContactID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ContactID").Value = iContactID
            uPar = .CreateParameter("@StatusID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@StatusID").Value = iStatusID
            .CommandText = "Job_UpdateStatus"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        UpdateJobStatus = True

Err_UpdateJobStatus:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "UpdateJobStatus", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "UpdateJobStatus" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            UpdateJobStatus = False
        End If

    End Function

    Public Function GetJobID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetJobID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            .CommandText = "Job_GetID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Job Class Object", sErrDescription)
            GetJobID = 0
            Exit Function
        End If

        ' Store Data Values
        GetJobID = uRecSnap("NewJobID").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetJobID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "GetJobID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetJobID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetJobID = 0
        End If

    End Function

    Public Sub GetJobTotals()

        '** Get Job Totals
        ' Error Checking
        On Error GoTo Err_GetJobTotals

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Item Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            .CommandText = "Job_GetTotals"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job Item - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Job Class Object", sErrDescription)
            Exit Sub
        End If

        ' Store Data Values
        dJobTotalMaterials = uRecSnap("JobTotalMaterials").Value
        dJobTotalMarkup = uRecSnap("JobTotalMarkup").Value
        dJobTotalHours = uRecSnap("JobTotalHours").Value
        dJobTotalCost = uRecSnap("JobTotalCost").Value
        dJobTotalActualPrice = uRecSnap("JobTotalActualPrice").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetJobTotals:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "GetJobTotals", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetJobTotals" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub


    Public Function LoadJobItem() As Boolean
        '**Finds the Current Job Item Data
        '**Raises an Error if Job Item not found

        ' Error Checking
        On Error GoTo Err_LoadJobItem

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ItemNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ItemNo").Value = iItemNo
            .CommandText = "JobItem_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            LoadJobItem = False
            Exit Function
        End If

        ' Store Data Values
        iItemNo = If(IsDBNull(uRecSnap("ItemNo").Value), 0, uRecSnap("ItemNo").Value)
        sDescription = If(IsDBNull(uRecSnap("Description").Value), "", uRecSnap("Description").Value)
        sMake = If(IsDBNull(uRecSnap("Make").Value), 0, uRecSnap("Make").Value)
        sModel = If(IsDBNull(uRecSnap("Model").Value), "", uRecSnap("Model").Value)
        sSerialNo = If(IsDBNull(uRecSnap("SerialNo").Value), "", uRecSnap("SerialNo").Value)
        sServiceTag = If(IsDBNull(uRecSnap("ServiceTag").Value), "", uRecSnap("ServiceTag").Value)
        sPassword = If(IsDBNull(uRecSnap("Password").Value), "", uRecSnap("Password").Value)
        sDescription = If(IsDBNull(uRecSnap("Description").Value), "", uRecSnap("Description").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadJobItem = True

Err_LoadJobItem:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "LoadJobItem", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadJobItem" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadJobItem = False
        End If

    End Function

    Public Function SaveJobItem() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveJobItem

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Job Item Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ItemNo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ItemNo").Value = iItemNo
            uPar = .CreateParameter("@Make", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40)
            .Parameters.Append(uPar)
            .Parameters("@Make").Value = sMake
            uPar = .CreateParameter("@Model", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40)
            .Parameters.Append(uPar)
            .Parameters("@Model").Value = sModel
            uPar = .CreateParameter("@SerialNo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@SerialNo").Value = sSerialNo
            uPar = .CreateParameter("@ServiceTag", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10)
            .Parameters.Append(uPar)
            .Parameters("@ServiceTag").Value = sServiceTag
            uPar = .CreateParameter("@Password", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20)
            .Parameters.Append(uPar)
            .Parameters("@Password").Value = sPassword
            uPar = .CreateParameter("@Description", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 100)
            .Parameters.Append(uPar)
            .Parameters("@Description").Value = sDescription
            .CommandText = "JobItem_SaveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveJobItem = True

Err_SaveJobItem:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "SaveJobItem", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveJobItem" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveJobItem = False
        End If

    End Function

    Public Function RemoveJobItem() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveJobItem

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Job Item Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ItemNo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ItemNo").Value = iItemNo
            .CommandText = "JobItem_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveJobItem = True

Err_RemoveJobItem:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "RemoveJobItem", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveJobItem" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveJobItem = False
        End If

    End Function

    Public Function GetJobItemIDNo()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetJobItemIDNo

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Item Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            .CommandText = "JobItem_GetID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job Item - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Job Class Object", sErrDescription)
            GetJobItemIDNo = 0
            Exit Function
        End If

        ' Store Data Values
        GetJobItemIDNo = uRecSnap("NewItemID").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetJobItemIDNo:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "GetJobItemIDNo", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetJobItemIDNo" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetJobItemIDNo = 0
        End If

    End Function



    Public Function LoadJobProduct() As Boolean
        '**Finds the Current Job Product Data
        '**Raises an Error if Job not found

        ' Error Checking
        On Error GoTo Err_LoadJobProduct

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("ProductID").Value = iProductID
            uPar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("RowNo").Value = iRowNo
            .CommandText = "JobProduct_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Check for Valid Job Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Job Class Object", sErrDescription)
            LoadJobProduct = False
            Exit Function
        End If

        ' Store Data Values
        dQuantity = If(IsDBNull(uRecSnap("Quantity").Value), 0, uRecSnap("Quantity").Value)
        dPrice = If(IsDBNull(uRecSnap("Price").Value), 0, uRecSnap("Price").Value)
        dMarkupPer = If(IsDBNull(uRecSnap("MarkupPer").Value), 0, uRecSnap("MarkupPer").Value)
        dLabourHours = If(IsDBNull(uRecSnap("LabourHours").Value), 0, uRecSnap("LabourHours").Value)
        dLabourRate = If(IsDBNull(uRecSnap("LabourRate").Value), 0, uRecSnap("LabourRate").Value)
        sFullDescription = If(IsDBNull(uRecSnap("FullDescription").Value), "", uRecSnap("FullDescription").Value)
        bShowOnQuote = If(IsDBNull(uRecSnap("ShowOnQuote").Value), 0, uRecSnap("ShowOnQuote").Value)


        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadJobProduct = True

Err_LoadJobProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "LoadJobProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadJobProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadJobProduct = False
        End If

    End Function

    Public Function SaveJobProduct() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveJobProduct

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            uPar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RowNo").Value = iRowNo
            uPar = .CreateParameter("@Quantity", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Quantity").Value = dQuantity
            uPar = .CreateParameter("@Price", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Price").Value = dPrice
            uPar = .CreateParameter("@ActualPrice", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ActualPrice").Value = dActualPrice
            uPar = .CreateParameter("@MarkupPer", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@MarkupPer").Value = dMarkupPer
            uPar = .CreateParameter("@LabourHours", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LabourHours").Value = dLabourHours
            uPar = .CreateParameter("@LabourRate", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LabourRate").Value = dLabourRate
            uPar = .CreateParameter("@FullDescription", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1000)
            .Parameters.Append(uPar)
            .Parameters("@FullDescription").Value = sFullDescription
            uPar = .CreateParameter("@ShowOnQuote", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ShowOnQuote").Value = -bShowOnQuote
            .CommandText = "JobProduct_SaveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveJobProduct = True

Err_SaveJobProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "SaveJobProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveJobProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveJobProduct = False
        End If

    End Function



    Public Function RemoveJobProduct() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveJobProduct

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Remove Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            uPar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RowNo").Value = iRowNo
            .CommandText = "JobProduct_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveJobProduct = True

Err_RemoveJobProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "RemoveJobProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveJobProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveJobProduct = False
        End If

    End Function

    Public Function UpdateJobProductOrderID() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_UpdateJobProductOrderID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            uPar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RowNo").Value = iRowNo
            .CommandText = "JobProduct_UpdateOrderID"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        UpdateJobProductOrderID = True

Err_UpdateJobProductOrderID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "UpdateJobProductOrderID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "UpdateJobProductOrderID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            UpdateJobProductOrderID = False
        End If

    End Function

    Public Function MoveRowUp()

        ' Error Checking
        On Error GoTo Err_MoveRowUp

        ' Dimension Local Variables
        Dim Downar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Move Row Down
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            Downar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(Downar)
            .Parameters("@JobID").Value = iJobID
            Downar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(Downar)
            .Parameters("@RowNo").Value = iRowNo
            .CommandText = "JobProduct_MoveRowUp"
            .Execute()
        End With

        ' Close Connection
        If bConnection Then CloseConnection()

Err_MoveRowUp:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "MoveRowUp", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "MoveRowUp" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Function

    Public Function MoveRowDown()

        ' Error Checking
        On Error GoTo Err_MoveRowDown

        ' Dimension Local Variables
        Dim Downar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Move Row Down
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            Downar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(Downar)
            .Parameters("@JobID").Value = iJobID
            Downar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(Downar)
            .Parameters("@RowNo").Value = iRowNo
            .CommandText = "JobProduct_MoveRowDown"
            .Execute()
        End With

        ' Close Connection
        If bConnection Then CloseConnection()

Err_MoveRowDown:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "MoveRowDown", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "MoveRowDown" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Function


    Public Function SaveJobEngineerTime() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveJobEngineerTime

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            uPar = .CreateParameter("@WeekNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@WeekNo").Value = iWeekNo
            uPar = .CreateParameter("@TaskNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@TaskNo").Value = iTaskNo
            uPar = .CreateParameter("@HoursNormal", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@HoursNormal").Value = dHoursNormal
            uPar = .CreateParameter("@HoursOvertime", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@HoursOvertime").Value = dHoursOvertime
            uPar = .CreateParameter("@HoursDouble", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@HoursDouble").Value = dHoursDouble
            uPar = .CreateParameter("@HoursTravel", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@HoursTravel").Value = dHoursTravel
            uPar = .CreateParameter("@Expenses", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Expenses").Value = dExpenses
            uPar = .CreateParameter("@Subsistence", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Subsistence").Value = dSubsistence
            uPar = .CreateParameter("@Mileage", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Mileage").Value = iMileage
            uPar = .CreateParameter("@BriefNotes", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 600)
            .Parameters.Append(uPar)
            .Parameters("@BriefNotes").Value = sBriefNotes
            .CommandText = "JobEngineerTime_SaveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveJobEngineerTime = True

Err_SaveJobEngineerTime:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "SaveJobEngineerTime", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveJobEngineerTime" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveJobEngineerTime = False
        End If

    End Function



    Public Function RemoveJobEngineerTime() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveJobEngineerTime

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Remove Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            uPar = .CreateParameter("@WeekNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@WeekNo").Value = iWeekNo
            .CommandText = "JobEngineerTime_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveJobEngineerTime = True

Err_RemoveJobEngineerTime:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "RemoveJobEngineerTime", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveJobEngineerTime" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveJobEngineerTime = False
        End If

    End Function

    Public Function LoadStatisticsByMonth() As Boolean
        '**Load Statistics on Current Data

        ' Error Checking
        On Error GoTo Err_LoadStatisticsByMonth

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@Month", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Month").Value = iMonth
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            .CommandText = "Job_StatisticsByMonth"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        iUnallocatedNo = If(IsDBNull(uRecSnap("UnallocatedNo").Value), 0, uRecSnap("UnallocatedNo").Value)
        iInProgressNo = If(IsDBNull(uRecSnap("InProgressNo").Value), 0, uRecSnap("InProgressNo").Value)
        dInProgressHours = If(IsDBNull(uRecSnap("InProgressHours").Value), 0, uRecSnap("InProgressHours").Value)
        iCompletedNo = If(IsDBNull(uRecSnap("CompletedNo").Value), 0, uRecSnap("CompletedNo").Value)
        dCompletedHours = If(IsDBNull(uRecSnap("CompletedHours").Value), 0, uRecSnap("CompletedHours").Value)
        iContractsNo = If(IsDBNull(uRecSnap("ContractsNo").Value), 0, uRecSnap("ContractsNo").Value)
        dContractsHours = If(IsDBNull(uRecSnap("ContractsHours").Value), 0, uRecSnap("ContractsHours").Value)
        dContractsTravelHours = If(IsDBNull(uRecSnap("ContractsTravelHours").Value), 0, uRecSnap("ContractsTravelHours").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadStatisticsByMonth = True

Err_LoadStatisticsByMonth:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "LoadStatisticsByMonth", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadStatisticsByMonth" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadStatisticsByMonth = False
        End If

    End Function

    Public Function LoadStatisticsByWeek() As Boolean
        '**Load Statistics on Current Data

        ' Error Checking
        On Error GoTo Err_LoadStatisticsByWeek

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@Week", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Week").Value = iWeek
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            .CommandText = "Job_StatisticsByWeek"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        iInProgressNo = If(IsDBNull(uRecSnap("InProgressNo").Value), 0, uRecSnap("InProgressNo").Value)
        dInProgressHours = If(IsDBNull(uRecSnap("InProgressHours").Value), 0, uRecSnap("InProgressHours").Value)
        iCompletedNo = If(IsDBNull(uRecSnap("CompletedNo").Value), 0, uRecSnap("CompletedNo").Value)
        dCompletedHours = If(IsDBNull(uRecSnap("CompletedHours").Value), 0, uRecSnap("CompletedHours").Value)
        iContractsNo = If(IsDBNull(uRecSnap("ContractsNo").Value), 0, uRecSnap("ContractsNo").Value)
        dContractsHours = If(IsDBNull(uRecSnap("ContractsHours").Value), 0, uRecSnap("ContractsHours").Value)
        dContractsTravelHours = If(IsDBNull(uRecSnap("ContractsTravelHours").Value), 0, uRecSnap("ContractsTravelHours").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadStatisticsByWeek = True

Err_LoadStatisticsByWeek:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "LoadStatisticsByWeek", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadStatisticsByWeek" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadStatisticsByWeek = False
        End If

    End Function

    Public Function LoadStatisticsByDay() As Boolean
        '**Load Statistics on Current Data

        ' Error Checking
        On Error GoTo Err_LoadStatisticsByDay

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Job Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@Day", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Day").Value = iDay
            uPar = .CreateParameter("@EngineerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@EngineerID").Value = iEngineerID
            .CommandText = "Job_StatisticsByDay"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        iInProgressNo = If(IsDBNull(uRecSnap("InProgressNo").Value), 0, uRecSnap("InProgressNo").Value)
        dInProgressHours = If(IsDBNull(uRecSnap("InProgressHours").Value), 0, uRecSnap("InProgressHours").Value)
        iCompletedNo = If(IsDBNull(uRecSnap("CompletedNo").Value), 0, uRecSnap("CompletedNo").Value)
        dCompletedHours = If(IsDBNull(uRecSnap("CompletedHours").Value), 0, uRecSnap("CompletedHours").Value)
        iContractsNo = If(IsDBNull(uRecSnap("ContractsNo").Value), 0, uRecSnap("ContractsNo").Value)
        dContractsHours = If(IsDBNull(uRecSnap("ContractsHours").Value), 0, uRecSnap("ContractsHours").Value)
        dContractsTravelHours = If(IsDBNull(uRecSnap("ContractsTravelHours").Value), 0, uRecSnap("ContractsTravelHours").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadStatisticsByDay = True

Err_LoadStatisticsByDay:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "LoadStatisticsByDay", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadStatisticsByDay" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadStatisticsByDay = False
        End If

    End Function

    Public Function UpdateDeliveryAddress() As Boolean
        '** Save Update Delivery Address

        ' Error Checking
        On Error GoTo Err_UpdateDeliveryAddress

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
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@DefaultAddress", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1)
            .Parameters.Append(uPar)
            .Parameters("@DefaultAddress").Value = sDefaultAddress
            uPar = .CreateParameter("@DAddress1", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@DAddress1").Value = sDAddress1
            uPar = .CreateParameter("@DAddress2", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@DAddress2").Value = sDAddress2
            uPar = .CreateParameter("@DAddress3", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@DAddress3").Value = sDAddress3
            uPar = .CreateParameter("@DAddress4", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@DAddress4").Value = sDAddress4
            uPar = .CreateParameter("@DPostCode", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8)
            .Parameters.Append(uPar)
            .Parameters("@DPostCode").Value = sDPostCode
            .CommandText = "Job_UpdateDeliveryAddress"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        UpdateDeliveryAddress = True

Err_UpdateDeliveryAddress:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsJob", "UpdateDeliveryAddress", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "UpdateDeliveryAddress" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            UpdateDeliveryAddress = False
        End If

    End Function
End Class
