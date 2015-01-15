Public Class clsOrder

    ' Dimension Class Variables - Order
    Private iOrderID As Integer
    Private iSupplierID As Integer
    Private iContactID As Integer
    Private iJobID As Integer
    Private sDeliverTo As String
    Private sAddress1 As String
    Private sAddress2 As String
    Private sAddress3 As String
    Private sAddress4 As String
    Private sPostCode As String
    Private sEMailAddress As String
    Private sCompany As String
    Private sSupplierFullName As String
    Private bMultiJobOrder As Integer
    Private sOrderDescription As String
    Private dCreateDate As Date
    Private dSaveDate As Date
    Private dDeliveryDate As Date
    Private sNotes As String
    Private iStatusID As Integer
    Private sStatusDescription As String
    Private sUserName As String
    Private iCheckForJobOrder As Integer
    Private iPurchaseOrderNo As Integer

    ' Dimension Class Variables - OrderProduct
    Private iProductID As Integer
    Private sProductDescription As String
    Private iRowNo As Integer
    Private iQuantity As Integer
    Private dPrice As Decimal
    Private sOrderProductDescription As String


    Public Property OrderID() As Integer
        '** Expose Order ID
        Get
            OrderID = iOrderID
        End Get
        Set(ByVal Value As Integer)
            iOrderID = Value
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

    Public ReadOnly Property SupplierFullName() As String
        '** Expose Supplier FullName
        Get
            SupplierFullName = sSupplierFullName
        End Get

    End Property

    Public Property OrderDescription() As String
        '** Expose Order Description
        Get
            OrderDescription = sOrderDescription
        End Get
        Set(ByVal Value As String)
            sOrderDescription = Value
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

    Public Property UserName() As String
        '** Expose User Name
        Get
            UserName = sUserName
        End Get
        Set(ByVal Value As String)
            sUserName = Value
        End Set

    End Property

    Public ReadOnly Property CheckForJobOrder() As Integer
        '** Expose CheckForJobOrder
        Get
            CheckForJobOrder = iCheckForJobOrder
        End Get

    End Property

    Public Property PurchaseOrderNo() As Integer
        '** Expose PurchaseOrderNo
        Get
            PurchaseOrderNo = iPurchaseOrderNo
        End Get
        Set(ByVal Value As Integer)
            iPurchaseOrderNo = Value
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

    Public Property DeliverTo() As String
        '** Expose DeliverTo
        Get
            DeliverTo = sDeliverTo
        End Get
        Set(ByVal Value As String)
            sDeliverTo = Value
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

    Public Property ContactID() As Integer
        '** Expose Contact ID
        Get
            ContactID = iContactID
        End Get
        Set(ByVal Value As Integer)
            iContactID = Value
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

    Public Property DeliveryDate() As Date
        '** Expose Delivery Date
        Get
            DeliveryDate = dDeliveryDate
        End Get
        Set(ByVal Value As Date)
            dDeliveryDate = Value
        End Set

    End Property

    Public Property MultiJobOrder() As Boolean
        '** Expose Multi Job Order
        Get
            MultiJobOrder = bMultiJobOrder
        End Get
        Set(ByVal Value As Boolean)
            bMultiJobOrder = Value
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

    Public Property Quantity() As Integer
        '** Expose Quantity
        Get
            Quantity = iQuantity
        End Get
        Set(ByVal Value As Integer)
            iQuantity = Value
        End Set

    End Property

    Public Property Price() As Decimal
        '** Expose Price
        Get
            Price = dPrice
        End Get
        Set(ByVal Value As Decimal)
            dPrice = Value
        End Set

    End Property

    Public Property OrderProductDescription() As String
        '** Expose Order Product Description
        Get
            OrderProductDescription = sOrderProductDescription
        End Get
        Set(ByVal Value As String)
            sOrderProductDescription = Value
        End Set

    End Property

    Public Function LoadOrder() As Boolean
        '**Finds the Current Order Data
        '**Raises an Error if Order not found

        ' Error Checking
        'On Error GoTo Err_LoadOrder

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
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            .CommandText = "Order_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Check for Valid Order - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            LoadOrder = False
            Exit Function
        End If

        ' Store Data Values
        iSupplierID = If(IsDBNull(uRecSnap("SupplierID").Value), 0, uRecSnap("SupplierID").Value)
        iContactID = If(IsDBNull(uRecSnap("ContactID").Value), 0, uRecSnap("ContactID").Value)
        iJobID = If(IsDBNull(uRecSnap("JobID").Value), 0, uRecSnap("JobID").Value)
        sDeliverTo = If(IsDBNull(uRecSnap("DeliverTo").Value), "", uRecSnap("DeliverTo").Value)
        sAddress1 = If(IsDBNull(uRecSnap("Address1").Value), "", uRecSnap("Address1").Value)
        sAddress2 = If(IsDBNull(uRecSnap("Address2").Value), "", uRecSnap("Address2").Value)
        sAddress3 = If(IsDBNull(uRecSnap("Address3").Value), "", uRecSnap("Address3").Value)
        sAddress4 = If(IsDBNull(uRecSnap("Address4").Value), "", uRecSnap("Address4").Value)
        sPostCode = If(IsDBNull(uRecSnap("PostCode").Value), "", uRecSnap("PostCode").Value)
        sEMailAddress = If(IsDBNull(uRecSnap("EMailAddress").Value), 0, uRecSnap("EMailAddress").Value)
        sCompany = If(IsDBNull(uRecSnap("Company").Value), 0, uRecSnap("Company").Value)
        sSupplierFullName = If(IsDBNull(uRecSnap("SupplierFullName").Value), 0, uRecSnap("SupplierFullName").Value)
        sOrderDescription = If(IsDBNull(uRecSnap("OrderDescription").Value), "", uRecSnap("OrderDescription").Value)
        dDeliveryDate = If(IsDBNull(uRecSnap("DeliveryDate").Value), "01/01/1900", uRecSnap("DeliveryDate").Value)
        bMultiJobOrder = If(IsDBNull(uRecSnap("MultiJobOrder").Value), 0, uRecSnap("MultiJobOrder").Value)
        sNotes = If(IsDBNull(uRecSnap("Notes").Value), "", uRecSnap("Notes").Value)
        iStatusID = If(IsDBNull(uRecSnap("StatusID").Value), 0, uRecSnap("StatusID").Value)
        sUserName = If(IsDBNull(uRecSnap("UserName").Value), "", uRecSnap("UserName").Value)
        sStatusDescription = If(IsDBNull(uRecSnap("StatusDescription").Value), "", uRecSnap("StatusDescription").Value)
        iCheckForJobOrder = If(IsDBNull(uRecSnap("CheckForJobOrder").Value), 0, uRecSnap("CheckForJobOrder").Value)
        iPurchaseOrderNo = If(IsDBNull(uRecSnap("PurchaseOrderNo").Value), 0, uRecSnap("PurchaseOrderNo").Value)

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadOrder = True

Err_LoadOrder:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsOrder", "LoadOrder", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadOrder" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadOrder = False
        End If

    End Function

    Public Function SaveOrder() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveOrder

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iOrderID = 0 Then iOrderID = GetOrderID()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If


        ' Run Stored Procedure - Save Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            uPar = .CreateParameter("@SupplierID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@SupplierID").Value = iSupplierID
            uPar = .CreateParameter("@ContactID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ContactID").Value = iContactID
            uPar = .CreateParameter("@JobID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@JobID").Value = iJobID
            uPar = .CreateParameter("@DeliverTo", ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1)
            .Parameters.Append(uPar)
            .Parameters("@DeliverTo").Value = IIf(sDeliverTo = Nothing, "", sDeliverTo)
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
            uPar = .CreateParameter("@OrderDescription", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 60)
            .Parameters.Append(uPar)
            .Parameters("@OrderDescription").Value = sOrderDescription
            uPar = .CreateParameter("@DeliveryDate", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@DeliveryDate").Value = dDeliveryDate
            uPar = .CreateParameter("@MultiJobOrder", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@MultiJobOrder").Value = bMultiJobOrder
            uPar = .CreateParameter("@Notes", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000)
            .Parameters.Append(uPar)
            .Parameters("@Notes").Value = sNotes
            uPar = .CreateParameter("@StatusID", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@StatusID").Value = iStatusID
            uPar = .CreateParameter("@UserName", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20)
            .Parameters.Append(uPar)
            .Parameters("@UserName").Value = sUserName
            .CommandText = "Order_SaveRecord"
            .Execute()
        End With
        uRecSnap = Nothing
        uCommand = Nothing

        ' Close Connection
        If bConnection Then CloseConnection()
        SaveOrder = True

Err_SaveOrder:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsOrder", "SaveOrder", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveOrder" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveOrder = False
        End If

    End Function

    Public Function RemoveOrder() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveOrder

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Order Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            .CommandText = "Order_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveOrder = True

Err_RemoveOrder:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsOrder", "RemoveOrder", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveOrder" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveOrder = False
        End If

    End Function


    Public Function GetOrderID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetOrderID

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
            .CommandText = "Order_GetID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Order - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Order Class Object", sErrDescription)
            GetOrderID = 0
            Exit Function
        End If

        ' Store Data Values
        GetOrderID = uRecSnap("NewIDNo").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetOrderID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsOrder", "GetOrderID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetOrderID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetOrderID = 0
        End If

    End Function

    Public Function GetOrderTotals() As Double

        '** Get Order Totals
        ' Error Checking
        On Error GoTo Err_GetOrderTotals

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Item Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            .CommandText = "Order_GetTotal"
            uRecSnap = .Execute
        End With

        ' Check for Valid Order Item - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Order Class Object", sErrDescription)
            GetOrderTotals = 0
            Exit Function
        End If

        ' Store Data Values
        GetOrderTotals = uRecSnap("OrderTotal").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetOrderTotals:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsOrder", "GetOrderTotals", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetOrderTotals" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetOrderTotals = 0
        End If

    End Function

    Public Function LoadOrderProduct() As Boolean
        '**Finds the Current Order Product Data
        '**Raises an Error if Order not found

        ' Error Checking
        On Error GoTo Err_LoadOrderProduct

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Order Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            .CommandText = "Order_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Check for Valid Order Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Order Class Object", sErrDescription)
            LoadOrderProduct = False
            Exit Function
        End If

        ' Store Data Values
        iProductID = If(IsDBNull(uRecSnap("iProductID").Value), 0, uRecSnap("iProductID").Value)
        sProductDescription = If(IsDBNull(uRecSnap("ProductDescription").Value), 0, uRecSnap("ProductDescription").Value)
        iQuantity = If(IsDBNull(uRecSnap("iQuantity").Value), 0, uRecSnap("iQuantity").Value)
        dPrice = If(IsDBNull(uRecSnap("dPrice").Value), 0, uRecSnap("dPrice").Value)
        sOrderProductDescription = If(IsDBNull(uRecSnap("OrderProductDescription").Value), 0, uRecSnap("OrderProductDescription").Value)


        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadOrderProduct = True

Err_LoadOrderProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsOrder", "LoadOrderProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadOrderProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadOrderProduct = False
        End If

    End Function

    Public Function SaveOrderProduct() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveOrderProduct

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Order Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            uPar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RowNo").Value = iRowNo
            uPar = .CreateParameter("@Quantity", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Quantity").Value = iQuantity
            uPar = .CreateParameter("@Price", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Price").Value = dPrice
            uPar = .CreateParameter("@OrderProductDescription", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 80)
            .Parameters.Append(uPar)
            .Parameters("@OrderProductDescription").Value = sOrderProductDescription
            .CommandText = "OrderProduct_SaveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing

        ' Check for Stock Order (SupplierID = 9999) with Order Delivered (StatusID = 3)
        If iStatusID = 3 And iSupplierID = 9999 Then
            ' Run Stored Procedure - Save Order Record
            uCommand = New ADODB.Command
            With uCommand
                .ActiveConnection = uDBase
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandTimeout = 0
                uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
                .Parameters.Append(uPar)
                .Parameters("@OrderID").Value = iOrderID
                .CommandText = "Product_UpdateStock"
                .Execute()
            End With
            uRecSnap = Nothing
            uCommand = Nothing
            iStatusID = 4 ' Stock Updated
        End If

        ' Close Connections
        If bConnection Then CloseConnection()
        SaveOrderProduct = True

Err_SaveOrderProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsOrder", "SaveOrderProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveOrderProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveOrderProduct = False
        End If

    End Function

    Public Function RemoveOrderProduct() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveOrderProduct

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Remove Order Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@OrderID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@OrderID").Value = iOrderID
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            uPar = .CreateParameter("@RowNo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@RowNo").Value = iRowNo
            .CommandText = "OrderProduct_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveOrderProduct = True

Err_RemoveOrderProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsOrder", "RemoveOrderProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveOrderProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveOrderProduct = False
        End If

    End Function

 
End Class
