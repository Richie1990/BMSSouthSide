Public Class clsProduct

    ' Dimension Class Variables - Product
    Private iProductID As Integer
    Private iCategoryID As Integer
    Private sCategoryDescription As String
    Private sProductDescription As String
    Private sProductFullDescription As String
    Private dPrice As Decimal
    Private dMarkupPer As Decimal
    Private dLabourHours As Decimal
    Private dLabourRate As Decimal
    Private iStock As Integer
    Private iOnOrder As Integer


    Public Property ProductID() As Integer
        '** Expose ProductID
        Get
            ProductID = iProductID
        End Get
        Set(ByVal Value As Integer)
            iProductID = Value
        End Set

    End Property

    Public Property CategoryID() As Integer
        '** Expose Category ID
        Get
            CategoryID = iCategoryID
        End Get
        Set(ByVal Value As Integer)
            iCategoryID = Value
        End Set

    End Property

    Public Property CategoryDescription() As String
        '** Expose Category Description
        Get
            CategoryDescription = sCategoryDescription
        End Get
        Set(ByVal Value As String)
            sCategoryDescription = Value
        End Set

    End Property

    Public Property ProductDescription() As String
        '** Expose Product Description
        Get
            ProductDescription = sProductDescription
        End Get
        Set(ByVal Value As String)
            sProductDescription = Value
        End Set

    End Property

    Public Property ProductFullDescription() As String
        '** Expose Product Full Description
        Get
            ProductFullDescription = sProductFullDescription
        End Get
        Set(ByVal Value As String)
            sProductFullDescription = Value
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

    Public Property MarkupPer() As Decimal
        '** Expose MarkupPer
        Get
            MarkupPer = dMarkupPer
        End Get
        Set(ByVal Value As Decimal)
            dMarkupPer = Value
        End Set

    End Property

    Public Property LabourHours() As Decimal
        '** Expose Labour Hours
        Get
            LabourHours = dLabourHours
        End Get
        Set(ByVal Value As Decimal)
            dLabourHours = Value
        End Set

    End Property

    Public Property LabourRate() As Decimal
        '** Expose Labour Cost
        Get
            LabourRate = dLabourRate
        End Get
        Set(ByVal Value As Decimal)
            dLabourRate = Value
        End Set

    End Property

    Public Property Stock() As Integer
        '** Expose Stock
        Get
            Stock = iStock
        End Get
        Set(ByVal Value As Integer)
            iStock = Value
        End Set

    End Property

    Public Property OnOrder() As Integer
        '** Expose OnOrder
        Get
            OnOrder = iOnOrder
        End Get
        Set(ByVal Value As Integer)
            iOnOrder = Value
        End Set

    End Property

    Public Function LoadProduct() As Boolean
        '**Finds the Current Job Data
        '**Raises an Error if Product not found

        ' Error Checking
        On Error GoTo Err_LoadProduct

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' CHeck for Blank Product
        If iProductID = 0 Then
            LoadProduct = False
            Exit Function
        End If

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            .CommandText = "Product_LoadRecord"
            uRecSnap = .Execute
        End With

        ' Store Data Values
        iCategoryID = If(IsDBNull(uRecSnap("CategoryID").Value), 0, uRecSnap("CategoryID").Value)
        sCategoryDescription = If(IsDBNull(uRecSnap("CategoryDescription").Value), "", uRecSnap("CategoryDescription").Value)
        sProductDescription = If(IsDBNull(uRecSnap("ProductDescription").Value), "", uRecSnap("ProductDescription").Value)
        sProductFullDescription = If(IsDBNull(uRecSnap("FullDescription").Value), "", uRecSnap("FullDescription").Value)
        dPrice = If(IsDBNull(uRecSnap("Price").Value), 0, uRecSnap("Price").Value)
        dMarkupPer = If(IsDBNull(uRecSnap("MarkupPer").Value), 0, uRecSnap("MarkupPer").Value)
        dLabourHours = If(IsDBNull(uRecSnap("LabourHours").Value), 0, uRecSnap("LabourHours").Value)
        dLabourRate = If(IsDBNull(uRecSnap("LabourRate").Value), 0, uRecSnap("LabourRate").Value)
        iStock = If(IsDBNull(uRecSnap("Stock").Value), 0, uRecSnap("Stock").Value)
        iOnOrder = If(IsDBNull(uRecSnap("OnOrder").Value), 0, uRecSnap("OnOrder").Value)


        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing
        LoadProduct = True

Err_LoadProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsProduct", "LoadProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "LoadProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            LoadProduct = False
        End If

    End Function

    Public Function SaveProduct() As Boolean
        '** Save Current Data Record

        ' Error Checking
        On Error GoTo Err_SaveProduct

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check for New ID Request
        If iProductID = 0 Then iProductID = GetProductID()

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Save Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            uPar = .CreateParameter("@CategoryID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@CategoryID").Value = iCategoryID
            uPar = .CreateParameter("@Description", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 80)
            .Parameters.Append(uPar)
            .Parameters("@Description").Value = sProductDescription
            uPar = .CreateParameter("@FullDescription", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 240)
            .Parameters.Append(uPar)
            .Parameters("@FullDescription").Value = sProductFullDescription
            uPar = .CreateParameter("@Price", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Price").Value = dPrice
            uPar = .CreateParameter("@MarkupPer", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@MarkupPer").Value = dMarkupPer
            uPar = .CreateParameter("@LabourHours", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LabourHours").Value = dLabourHours
            uPar = .CreateParameter("@LabourRate", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@LabourRate").Value = dLabourRate
            uPar = .CreateParameter("@Stock", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@Stock").Value = iStock
            .CommandText = "Product_SaveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        SaveProduct = True

Err_SaveProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsProduct", "SaveProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "SaveProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            SaveProduct = False
        End If

    End Function

    Public Function RemoveProduct() As Boolean
        '** Remove Requested Item

        ' Error Checking
        On Error GoTo Err_RemoveProduct

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
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            .CommandText = "Product_RemoveRecord"
            .Execute()
        End With

        ' Close Connection
        uRecSnap = Nothing
        uCommand = Nothing
        If bConnection Then CloseConnection()
        RemoveProduct = True

Err_RemoveProduct:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsProduct", "RemoveProduct", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "RemoveProduct" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            RemoveProduct = False
        End If

    End Function

    Public Function GetProductID()

        '** Get New ID No
        ' Error Checking
        On Error GoTo Err_GetProductID

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            .CommandText = "Product_GetID"
            uRecSnap = .Execute
        End With

        ' Check for Valid Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Product Class Object", sErrDescription)
            GetProductID = 0
            Exit Function
        End If

        ' Store Data Values
        GetProductID = uRecSnap("NewIDNo").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetProductID:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsProduct", "GetProductID", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetProductID" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetProductID = 0
        End If

    End Function

    Public Function GetProductPrice()

        ' Error Checking
        On Error GoTo Err_GetProductPrice

        ' Dimension Local Variables
        Dim uRecSnap As ADODB.Recordset
        Dim uPar As ADODB.Parameter

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Load Product Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@ProductID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            .Parameters.Append(uPar)
            .Parameters("@ProductID").Value = iProductID
            .CommandText = "Product_GetPrice"
            uRecSnap = .Execute
        End With

        ' Check for Valid Product - Not Found
        If uRecSnap.BOF And uRecSnap.EOF Then
            Err.Raise(Err.Number, "Product Class Object", sErrDescription)
            GetProductPrice = 0
            Exit Function
        End If

        ' Store Data Values
        GetProductPrice = uRecSnap("Price").Value

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

Err_GetProductPrice:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("clsProduct", "GetProductPrice", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "GetProductPrice" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
            GetProductPrice = 0
        End If

    End Function

End Class
