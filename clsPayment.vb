Public Class clsPayment

    ' Dimension Class Variables
    ' General
    Private iOrderID As Integer
    Private iJobID As Integer
    Private iInvoiceID As Integer
    Private dInvoiceDate As Date
    Private dAmount As Double
    Private dValuationAmount As Double
    Private dMCDPer As Double
    Private dRetentionPer As Double
    Private sNotes As String
    Private bApprovedFlag As Boolean
    Private dApprovedDate As Date
    Private bPaidFlag As Boolean
    Private dPaidDate As Date
    Private dApprovedPaidAmount As Double

    ' Purchase Invoice
    Private iSupplierID As Integer
    Private sCompanyName As String

    ' Sales Invoice
    Private iInvoicePaymentID As Integer
    Private dInvoicePaymentDate As Date
    Private iPaymentTypeID As Integer
    Private sPaymentTypeDescription As String
    Private dTotalAmount As Double
    Private dTotalPaid As Double
    Private sXeroID As String
    Private dAmountInvoiced As Double
    Private dAmountPaid As Double

    ' Sales Invoice Item
    Private iRowNo As Integer
    Private sItemDescription As String
    Private iQuantity As Integer
    Private dPrice As Double
    Private dTotalPrice As Double

    ' Dimension Class Variables - JobProducts Recordset
    Dim uJobProduct As ADODB.Recordset


    Public Property OrderID() As Integer
        '** Expose OrderID
        Get
            OrderID = iOrderID
        End Get
        Set(ByVal Value As Integer)
            iOrderID = Value
        End Set

    End Property

    Public Property JobID() As Integer
        '** Expose JobID
        Get
            JobID = iJobID
        End Get
        Set(ByVal Value As Integer)
            iJobID = Value
        End Set

    End Property

    Public Property InvoiceID() As Integer
        '** Expose Invoice ID
        Get
            InvoiceID = iInvoiceID
        End Get
        Set(ByVal Value As Integer)
            iInvoiceID = Value
        End Set

    End Property

    Public Property InvoiceDate() As Date
        '** Expose Invoice Date
        Get
            InvoiceDate = dInvoiceDate
        End Get
        Set(ByVal Value As Date)
            dInvoiceDate = Value
        End Set

    End Property

    Public Property Amount() As Double
        '** Expose Amount
        Get
            Amount = dAmount
        End Get
        Set(ByVal Value As Double)
            dAmount = Value
        End Set

    End Property

    Public Property ValuationAmount() As Double
        '** Expose Valuation Amount
        Get
            ValuationAmount = dValuationAmount
        End Get
        Set(ByVal Value As Double)
            dValuationAmount = Value
        End Set

    End Property

    Public Property MCDPer() As Double
        '** Expose MCD Per
        Get
            MCDPer = dMCDPer
        End Get
        Set(ByVal Value As Double)
            dMCDPer = Value
        End Set

    End Property

    Public Property RetentionPer() As Double
        '** Expose Retention Per
        Get
            RetentionPer = dRetentionPer
        End Get
        Set(ByVal Value As Double)
            dRetentionPer = Value
        End Set

    End Property

    Public Property SupplierID() As Integer
        '** Expose Supplier ID
        Get
            SupplierID = iSupplierID
        End Get
        Set(ByVal Value As Integer)
            iSupplierID = Value
        End Set

    End Property

    Public Property CompanyName() As String
        '** Expose Company Name
        Get
            CompanyName = sCompanyName
        End Get
        Set(ByVal Value As String)
            sCompanyName = Value
        End Set

    End Property

    Public Property ApprovedFlag() As Boolean
        '** Expose Approved Flag
        Get
            ApprovedFlag = bApprovedFlag
        End Get
        Set(ByVal Value As Boolean)
            bApprovedFlag = Value
        End Set

    End Property

    Public Property ApprovedDate() As Date
        '** Expose PI Approved Date
        Get
            ApprovedDate = dApprovedDate
        End Get
        Set(ByVal Value As Date)
            dApprovedDate = Value
        End Set

    End Property

    Public Property PaidFlag() As Boolean
        '** Expose Paid Flag
        Get
            PaidFlag = bPaidFlag
        End Get
        Set(ByVal Value As Boolean)
            bPaidFlag = Value
        End Set

    End Property

    Public Property PaidDate() As Date
        '** Expose PI Paid Date
        Get
            PaidDate = dPaidDate
        End Get
        Set(ByVal Value As Date)
            dPaidDate = Value
        End Set

    End Property

    Public Property ApprovedPaidAmount() As Double
        '** Expose Approved / Paid Amount
        Get
            ApprovedPaidAmount = dApprovedPaidAmount
        End Get
        Set(ByVal Value As Double)
            dApprovedPaidAmount = Value
        End Set

    End Property

    Public Property Notes() As String
        '** Expose PI Notes
        Get
            Notes = sNotes
        End Get
        Set(ByVal Value As String)
            sNotes = Value
        End Set

    End Property

    Public Property InvoicePaymentID() As Integer
        '** Expose InvoicePaymentID
        Get
            InvoicePaymentID = iInvoicePaymentID
        End Get
        Set(ByVal Value As Integer)
            iInvoicePaymentID = Value
        End Set

    End Property

    Public Property InvoicePaymentDate() As Date
        '** Expose Invoice Payment Date
        Get
            InvoicePaymentDate = dInvoicePaymentDate
        End Get
        Set(ByVal Value As Date)
            dInvoicePaymentDate = Value
        End Set

    End Property

    Public Property PaymentTypeID() As Integer
        '** Expose Payment Type ID
        Get
            PaymentTypeID = iPaymentTypeID
        End Get
        Set(ByVal Value As Integer)
            iPaymentTypeID = Value
        End Set

    End Property

    Public Property PaymentTypeDescription() As String
        '** Expose PaymentType Description
        Get
            PaymentTypeDescription = sPaymentTypeDescription
        End Get
        Set(ByVal Value As String)
            sPaymentTypeDescription = Value
        End Set

    End Property

    Public ReadOnly Property TotalAmount() As Double
        '** Expose Total Amount
        Get
            TotalAmount = dTotalAmount
        End Get

    End Property

    Public ReadOnly Property TotalPaid() As Double
        '** Expose Total Paid
        Get
            TotalPaid = dTotalPaid
        End Get

    End Property

    Public Property XeroID() As String
        '** Expose XeroID
        Get
            XeroID = sXeroID
        End Get
        Set(ByVal Value As String)
            sXeroID = Value
        End Set

    End Property

    Public ReadOnly Property AmountInvoiced() As Double
        '** Expose Amount Invoiced
        Get
            AmountInvoiced = dAmountInvoiced
        End Get

    End Property

    Public ReadOnly Property AmountPaid() As Double
        '** Expose Amount Paid
        Get
            AmountPaid = dAmountPaid
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

    Public Property ItemDescription() As String
        '** Expose Item Description
        Get
            ItemDescription = sItemDescription
        End Get
        Set(ByVal Value As String)
            sItemDescription = Value
        End Set

    End Property

    Public Property Quantity() As Integer
        '** Expose Quantity
        Get
            Quantity = iQuantity
        End Get
        Set(ByVal Value As Integer)
            iQuantity = Value
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

    Public ReadOnly Property TotalPrice() As Double
        '** Expose Total Price
        Get
            TotalPrice = dTotalPrice
        End Get

    End Property

    Public ReadOnly Property JobProduct() As ADODB.Recordset
        '** Expose Job Products RecordSet
        Get
            JobProduct = uJobProduct
        End Get

    End Property

    Public Function LoadPurchaseInvoice() As Boolean
        '**Finds the Current Purchase Invoice Data
        '**Raises an Error if Purchase Invoice not found

        ' Error Checking
        On Error GoTo Err_LoadPurchaseInvoice

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@PurchaseInvoiceID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@PurchaseInvoiceID").Value = iInvoiceID
            .CommandText = "Payment_LoadPurchaseInvoice"
            uRecSnap = .Execute
        End With

        ' Check for Valid Purchase Invoice - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "PurchaseInvoice Class Object", sErrDescription)
            LoadPurchaseInvoice = False
            Exit Function
        End If

        ' Retrieve Data Values
        iOrderID = uRecSnap("OrderID").Value
        iSupplierID = uRecSnap("SupplierID").Value
        sCompanyName = uRecSnap("SupplierName").Value
        dAmount = uRecSnap("Amount").Value
        bApprovedFlag = uRecSnap("ApprovedFlag").Value
        dApprovedDate = uRecSnap("ApprovedDate").Value
        bPaidFlag = uRecSnap("PaidFlag").Value
        dPaidDate = uRecSnap("PaidDate").Value
        sNotes = IIf(IsDBNull(uRecSnap("Notes").Value), "", uRecSnap("Notes").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadPurchaseInvoice = True

Err_LoadPurchaseInvoice:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "LoadPurchaseInvoice", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadPurchaseInvoice" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadPurchaseInvoice = False
        End If

    End Function

    Public Function SavePurchaseInvoice() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SavePurchaseInvoice

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iInvoiceID = 0 Then iInvoiceID = GetPurchaseInvoiceID()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoiceID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoiceID").Value = iInvoiceID
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            uPar = .CreateParameter("@InvoiceDate", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoiceDate").Value = dInvoiceDate
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            uPar = .CreateParameter("@Amount", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Amount").Value = dAmount
            uPar = .CreateParameter("@Notes", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 600)
            .Parameters.Append(uPar)
            .Parameters("@Notes").Value = sNotes
            .CommandText = "Payment_SavePurchaseInvoice"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SavePurchaseInvoice = True

Err_SavePurchaseInvoice:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "SavePurchaseInvoice", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SavePurchaseInvoice" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SavePurchaseInvoice = False
        End If

    End Function

    Public Function UpdatePurchaseInvoiceApprovedStatus() As Boolean
        '** Update Approved Status - Purchase Invoice Record

        ' Error Checking
        On Error GoTo Err_UpdatePurchaseInvoiceApprovedStatus

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoiceID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoiceID").Value = iInvoiceID
            uPar = .CreateParameter("@ApprovedFlag", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ApprovedFlag").Value = -bApprovedFlag
            uPar = .CreateParameter("@ApprovedDate", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ApprovedDate").Value = dApprovedDate
            uPar = .CreateParameter("@Notes", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 600)
            .Parameters.Append(uPar)
            .Parameters("@Notes").Value = sNotes
            .CommandText = "Payment_UpdatePurchaseInvoiceApprovedStatus"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        UpdatePurchaseInvoiceApprovedStatus = True

Err_UpdatePurchaseInvoiceApprovedStatus:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "UpdatePurchaseInvoiceApprovedStatus", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "UpdatePurchaseInvoiceApprovedStatus" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            UpdatePurchaseInvoiceApprovedStatus = False
        End If

    End Function

    Public Function UpdatePurchaseInvoicePaidStatus() As Boolean
        '** Update Paid Status - Purchase Invoice Record

        ' Error Checking
        On Error GoTo Err_UpdatePurchaseInvoicePaidStatus

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoiceID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoiceID").Value = iInvoiceID
            uPar = .CreateParameter("@PaidFlag", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@PaidFlag").Value = -bPaidFlag
            uPar = .CreateParameter("@PaidDate", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@PaidDate").Value = dPaidDate
            uPar = .CreateParameter("@Notes", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 600)
            .Parameters.Append(uPar)
            .Parameters("@Notes").Value = sNotes
            .CommandText = "Payment_UpdatePurchaseInvoicePaidStatus"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        UpdatePurchaseInvoicePaidStatus = True

Err_UpdatePurchaseInvoicePaidStatus:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "UpdatePurchaseInvoicePaidStatus", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "UpdatePurchaseInvoicePaidStatus" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            UpdatePurchaseInvoicePaidStatus = False
        End If

    End Function

    Public Function ClearPurchaseInvoiceStatusFlags() As Boolean
        '** Clear Purchase Invoice Status Flags

        ' Error Checking
        On Error GoTo Err_ClearPurchaseInvoiceStatusFlags

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoiceID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoiceID").Value = iInvoiceID
            .CommandText = "Payment_ClearPurchaseInvoiceStatusFlags"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        ClearPurchaseInvoiceStatusFlags = True

Err_ClearPurchaseInvoiceStatusFlags:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "ClearPurchaseInvoiceStatusFlags", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "ClearPurchaseInvoiceStatusFlags" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            ClearPurchaseInvoiceStatusFlags = False
        End If

    End Function

    Public Function GetPurchaseInvoiceID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetPurchaseInvoiceID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            .CommandText = "Payment_GetPurchaseInvoiceID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Order - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Order Class Object", sErrDescription)
            GetPurchaseInvoiceID = 0
            Exit Function
        End If

        ' Store Data Values
        GetPurchaseInvoiceID = uRecSnap("NewInvoiceID").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetPurchaseInvoiceID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayment", "GetPurchaseInvoiceID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetPurchaseInvoiceID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetPurchaseInvoiceID = 0
        End If

    End Function

    Public Function LoadSalesInvoice() As Boolean
        '**Finds the Current Sales Invoice Data
        '**Raises an Error if Sales Invoice not found

        ' Error Checking
        On Error GoTo Err_LoadSalesInvoice

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoicePaymentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentID").Value = iInvoicePaymentID
            .CommandText = "Payment_LoadSalesInvoice"
            uRecSnap = .Execute
        End With

        ' Check for Valid Sales Invoice - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Sales Invoice Class Object", sErrDescription)
            LoadSalesInvoice = False
            Exit Function
        End If

        ' Retrieve Data Values
        iJobID = uRecSnap("JobID").Value
        dInvoiceDate = uRecSnap("InvoicePaymentDate").Value
        dAmount = uRecSnap("Amount").Value
        dMCDPer = IIf(IsDBNull(uRecSnap("MCDPer").Value), 0, uRecSnap("MCDPer").Value)
        dRetentionPer = IIf(IsDBNull(uRecSnap("RetentionPer").Value), 0, uRecSnap("RetentionPer").Value)
        iPaymentTypeID = uRecSnap("PaymentTypeID").Value
        sPaymentTypeDescription = uRecSnap("PaymentTypeDescription").Value
        sNotes = IIf(IsDBNull(uRecSnap("Notes").Value), "", uRecSnap("Notes").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadSalesInvoice = True

Err_LoadSalesInvoice:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "LoadSalesInvoice", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadSalesInvoice" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadSalesInvoice = False
        End If

    End Function

    Public Function SaveSalesInvoice() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveSalesInvoice

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iInvoiceID = 0 And iInvoicePaymentID = 0 Then iInvoicePaymentID = GetSalesInvoiceID()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save SalesInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoicePaymentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentID").Value = iInvoicePaymentID
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@InvoicePaymentDate", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentDate").Value = dInvoicePaymentDate
            uPar = .CreateParameter("@PaymentTypeID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@PaymentTypeID").Value = iPaymentTypeID
            uPar = .CreateParameter("@Amount", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Amount").Value = dAmount
            uPar = .CreateParameter("@MCDPer", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@MCDPer").Value = dMCDPer
            uPar = .CreateParameter("@RetentionPer", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RetentionPer").Value = dRetentionPer
            uPar = .CreateParameter("@Notes", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 2000)
            .Parameters.Append(uPar)
            .Parameters("@Notes").Value = sNotes
            .CommandText = "Payment_SaveSalesInvoice"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveSalesInvoice = True

Err_SaveSalesInvoice:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "SaveSalesInvoice", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveSalesInvoice" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveSalesInvoice = False
        End If

    End Function



    Public Function ClearSalesInvoiceStatusFlags() As Boolean
        '** Update Approved Status - Sales Invoice Record

        ' Error Checking
        On Error GoTo Err_ClearSalesInvoiceStatusFlags

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoicePaymentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentID").Value = iInvoicePaymentID
            .CommandText = "Payment_ClearSalesInvoiceStatusFlags"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        ClearSalesInvoiceStatusFlags = True

Err_ClearSalesInvoiceStatusFlags:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "ClearSalesInvoiceStatusFlags", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "ClearSalesInvoiceStatusFlags" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            ClearSalesInvoiceStatusFlags = False
        End If

    End Function

    Public Sub GetSalesInvoiceTotals()

        '** Get Total of Sales Invoices
        ' Error Checking
        On Error GoTo Err_GetSalesInvoiceTotals

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            .CommandText = "Payment_GetSalesInvoiceTotals"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        dTotalAmount = uRecSnap("JobTotalAmount").Value
        dValuationAmount = uRecSnap("JobValuationAmount").Value
        dTotalPaid = uRecSnap("JobTotalPaid").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetSalesInvoiceTotals:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayment", "GetSalesInvoiceTotals", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetSalesInvoiceTotals" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Public Sub GetXeroSalesInvoices()

        '** Get Details of Xero Sales Invoices
        ' Error Checking
        On Error GoTo Err_GetXeroSalesInvoices

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            .CommandText = "Payment_LoadXeroSalesInvoices"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        dAmountInvoiced = uRecSnap("AmountInvoiced").Value
        dAmountPaid = uRecSnap("AmountPaid").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetXeroSalesInvoices:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayment", "GetXeroSalesInvoices", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetXeroSalesInvoices" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Public Sub GetXeroJobProducts()

        '** Get Details of Job Products
        ' Error Checking
        On Error GoTo Err_GetXeroJobProducts

        ' Dimension Local Variables
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            .CommandText = "Payment_LoadJobProducts"
            uJobProduct = .Execute
        End With


        ' Close Connection
        If bConnection Then CloseConnection()

Err_GetXeroJobProducts:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayment", "GetXeroJobProducts", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetXeroJobProducts" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Public Function GetSalesInvoiceID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetSalesInvoiceID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            .CommandText = "Payment_GetSalesInvoiceID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Order - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Order Class Object", sErrDescription)
            GetSalesInvoiceID = 0
            Exit Function
        End If

        ' Store Data Values
        GetSalesInvoiceID = uRecSnap("NewInvoiceID").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetSalesInvoiceID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayment", "GetSalesInvoiceID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetSalesInvoiceID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetSalesInvoiceID = 0
        End If

    End Function

    Public Function SaveSalesInvoiceItem() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveSalesInvoiceItem

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iRowNo = 0 Then iRowNo = GetSalesInvoiceRowNo()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save PurchaseInvoiceItem Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoicePaymentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentID").Value = iInvoicePaymentID
            uPar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RowNo").Value = iRowNo
            uPar = .CreateParameter("@Description", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 120)
            .Parameters.Append(uPar)
            .Parameters("@Description").Value = sItemDescription
            uPar = .CreateParameter("@Quantity", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Quantity").Value = iQuantity
            uPar = .CreateParameter("@Price", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Price").Value = dPrice
            .CommandText = "Payment_SaveSalesInvoiceItem"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveSalesInvoiceItem = True

Err_SaveSalesInvoiceItem:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "SaveSalesInvoiceItem", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveSalesInvoiceItem" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveSalesInvoiceItem = False
        End If

    End Function


    Public Function RemoveSalesInvoiceItem() As Boolean
        '** Remove Sales Invocie Item Record

        ' Error Checking
        On Error GoTo Err_RemoveSalesInvoiceItem

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoicePaymentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentID").Value = iInvoicePaymentID
            uPar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RowNo").Value = iRowNo
            .CommandText = "Payment_RemoveSalesInvoiceItem"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveSalesInvoiceItem = True

Err_RemoveSalesInvoiceItem:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "RemoveSalesInvoiceItem", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveSalesInvoiceItem" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveSalesInvoiceItem = False
        End If

    End Function


    Public Function GetSalesInvoiceRowNo()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetSalesInvoiceRowNo

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoicePaymentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentID").Value = iInvoicePaymentID
            .CommandText = "Payment_GetSalesInvoiceItemRowNo"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        GetSalesInvoiceRowNo = uRecSnap("NewRowNo").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetSalesInvoiceRowNo:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayment", "GetSalesInvoiceRowNo", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetSalesInvoiceRowNo" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetSalesInvoiceRowNo = 0
        End If

    End Function

    Public Sub GetSalesInvoiceItemTotals()

        '** Get Total of Sales Invoices Items
        ' Error Checking
        On Error GoTo Err_GetSalesInvoiceItemTotals

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@InvoicePaymentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentID").Value = iInvoicePaymentID
            .CommandText = "Payment_GetSalesInvoiceItemTotals"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        dTotalPrice = uRecSnap("InvoiceTotalPrice").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetSalesInvoiceItemTotals:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayment", "GetSalesInvoiceItemTotals", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetSalesInvoiceItemTotals" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub


    Public Function UpdateSalesInvoiceXeroID() As Boolean
        '** Update XeroID on Sales Invoice

        ' Error Checking
        On Error GoTo Err_UpdateSalesInvoiceXeroID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save PurchaseInvoice Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@InvoicePaymentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@InvoicePaymentID").Value = iInvoicePaymentID
            uPar = .CreateParameter("@PaymentTypeID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@PaymentTypeID").Value = iPaymentTypeID
            uPar = .CreateParameter("@XeroID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40)
            .Parameters.Append(uPar)
            .Parameters("@XeroID").Value = sXeroID
            .CommandText = "Payment_UpdateSalesInvoiceXeroID"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        UpdateSalesInvoiceXeroID = True

Err_UpdateSalesInvoiceXeroID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsPayments", "UpdateSalesInvoiceXeroID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "UpdateSalesInvoiceXeroID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            UpdateSalesInvoiceXeroID = False
        End If

    End Function
End Class
