Imports System.Configuration

Module modGlobal

    'Global Database Objects
    Public uDBase As ADODB.Connection
    Public uAccessDBase As ADODB.Connection
    Public uCommand As ADODB.Command
    Public bConnection As Boolean
    Public sErrDescription As String
    Public sSystemPath As String

    Public ReadOnly Property SystemPath() As String
        '** Expose System Path
        Get
            If My.Settings.DocumentsLocation = "" Then
                SystemPath = My.Settings.DropBoxPath & "\" & Environment.UserName & "\" & My.Settings.DropBoxName & "\Business Management System\"
                If Not (System.IO.Directory.Exists(SystemPath)) Then
                    SystemPath = My.Settings.DropBoxPath & "\" & Environment.UserName & "." & My.Settings.CompanyName & "\" & My.Settings.DropBoxName & "\Business Management System\"
                End If
            Else
                SystemPath = My.Settings.DocumentsLocation
            End If

        End Get


    End Property

    Public Sub WriteAuditLogRecord(ByVal sObject As String, ByVal sSource As String, ByVal sType As String, ByVal sDescription As String)
        '** Write Audit Log Record

        ' Ensure No Errors
        On Error Resume Next

        ' Dimension Local Variables
        Dim uLDBase As ADODB.Connection
        Dim uLCommand As ADODB.Command
        Dim uPar As ADODB.Parameter

        ' Open Own Connection
        uLDBase = New ADODB.Connection
        uLDBase.ConnectionString = My.Settings.ConnectionString.ToString
        uLDBase.CommandTimeout = 0
        uLDBase.Open()

        ' Run Stored Procedure - Insert Audit Log Record
        uLCommand = New ADODB.Command
        With uLCommand
            .ActiveConnection = uLDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            uPar = .CreateParameter("@UserName", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50)
            .Parameters.Append(uPar)
            .Parameters("@UserName").Value = Environment.UserName
            uPar = .CreateParameter("@Object", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Object").Value = sObject
            uPar = .CreateParameter("@Source", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30)
            .Parameters.Append(uPar)
            .Parameters("@Source").Value = sSource
            uPar = .CreateParameter("@Type", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 6)
            .Parameters.Append(uPar)
            .Parameters("@Type").Value = sType
            uPar = .CreateParameter("@Description", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 400)
            .Parameters.Append(uPar)
            .Parameters("@Description").Value = sDescription
            .CommandText = "AuditLog_SaveRecord"
            .Execute()
        End With

        ' Close Connection
        uLCommand = Nothing
        uLDBase = Nothing
        On Error GoTo 0

    End Sub

    Public Sub OpenConnection()
        '** Open Database Connection

        ' Error Checking
        '       On Error GoTo Err_OpenConnection

        ' Open Database Connection
        uDBase = New ADODB.Connection
        uDBase.ConnectionString = My.Settings.ConnectionString.ToString
        uDBase.CommandTimeout = 0
        uDBase.Open()

Err_OpenConnection:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description & Chr(13) & My.Settings.ConnectionString.ToString
            WriteAuditLogRecord("CJMGlobal", "OpenConnection", "Error", sErrDescription)
            CloseConnection()
            MsgBox("System Error occurred" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

    Public Sub OpenAccessConnection()
        '** Open Access Database Connection

        ' Error Checking
        On Error GoTo Err_OpenAccessConnection


        ' Open Database Connection
        uAccessDBase = New ADODB.Connection
        uAccessDBase.ConnectionString = My.Settings.AccessDatabase.ToString
        uAccessDBase.CommandTimeout = 0
        uAccessDBase.Open()

Err_OpenAccessConnection:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description & Chr(13) & My.Settings.AccessDatabase.ToString
            WriteAuditLogRecord("CJMGlobal", "OpenAccessConnection", "Error", sErrDescription)
            CloseConnection()
            MsgBox("System Error occurred" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Sub

 

    Sub CloseConnection()
        '** Close Database Connection
        uDBase = Nothing
        uCommand = Nothing

    End Sub

    Sub CloseAccessConnection()
        '** Close Remote Database Connection
        uAccessDBase.Close()
        uCommand = Nothing

    End Sub


    Function ExtractDataString(ByVal sData As String, ByVal iPos As Integer, ByVal sChar As String) As String
        '** Extract Data from a String delimited by any character

        ' Dimension Local Variables
        Dim iCnt As Integer
        Dim iPrevPos, iNextPos As Integer

        ' Extract Data
        iNextPos = 0
        For iCnt = 1 To iPos
            iPrevPos = iNextPos
            iNextPos = InStr(iNextPos + 1, sData, sChar)
        Next
        ExtractDataString = sData.Substring(iPrevPos, iNextPos - iPrevPos - 1)

    End Function

    Sub SetupCombos(ByVal uForm As Form)
        '** Populate Combo Boxes on Form with Data
        '** Get Source Data from Tag Property (i.e. Stored Procedure Name)

        ' Dimension Local Variables
        Dim ControlParent, ControlChild As Control

        ' Loop All Controls / Child Controls
        For Each ControlParent In uForm.Controls
            For Each ControlChild In ControlParent.Controls
                If TypeOf ControlChild Is ComboBox And ControlChild.Tag <> "" Then
                    PopulateCombo(ControlChild)
                End If
            Next ControlChild
        Next ControlParent
        ControlChild = Nothing
        ControlParent = Nothing

    End Sub

    Public Sub PopulateCombo(ByVal uComboBox As ComboBox)
        '** Populate Combo With Data from Database Call

        ' Dimension Local Variables
        Dim ComboArray As New ArrayList
        Dim uRecSnap As ADODB.Recordset

        ' Check For Open Connection
        If uDBase Is Nothing Then
            OpenConnection()
            bConnection = True
        End If

        ' Run Stored Procedure - Get Client Record
        uCommand = New ADODB.Command
        With uCommand
            .ActiveConnection = uDBase
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandTimeout = 0
            .CommandText = uComboBox.Tag
            uRecSnap = .Execute
        End With

        ' Clear Data
        uComboBox.DataSource = Nothing
        uComboBox.Items.Clear()
        ComboArray.Add(New ComboValueDescription(uComboBox, "0", ""))

        'Read Dataset & Add to ArrayList
        Do While Not uRecSnap.EOF
            ComboArray.Add(New ComboValueDescription(uComboBox, uRecSnap.Fields!ID.Value, uRecSnap.Fields!Value.Value.ToString))
            uRecSnap.MoveNext()
        Loop

        ' Close Connection
        uRecSnap.Close()
        If bConnection Then CloseConnection()
        uRecSnap = Nothing

        ' Populate Combo Box
        With uComboBox
            .DisplayMember = "Value"
            .ValueMember = "ID"
            .DataSource = ComboArray
        End With

    End Sub

    Public Sub SelectComboRecord(ByVal uComboBox As ComboBox, ByVal sItem As String)
        '** Searches for sItem in uCombo / Selects sItems. 

        ' Dimension Local Variables
        Dim iCnt As Integer

        'Look for sItem in the Combobox
        For iCnt = 0 To uComboBox.Items.Count - 1
            If uComboBox.Items.Item(iCnt).ToString.ToUpper = sItem.ToString.ToUpper Then
                uComboBox.SelectedIndex = iCnt
                Exit Sub
            End If
        Next

    End Sub

    Public Sub ClearControls(ByVal uForm As Form, ByVal uParent As Control)
        '** Clear All Child Controls in Parent Control

        ' Dimension Local Variables
        Dim ControlParent, ControlChild As Control

        ' Loop All Controls / Child Controls
        For Each ControlParent In uForm.Controls
            If ControlParent.Name = uParent.Name Then
                For Each ControlChild In ControlParent.Controls
                    If TypeOf ControlChild Is TextBox Then
                        ControlChild.Text = ""
                    End If
                    If TypeOf ControlChild Is ComboBox Then
                        ControlChild.Text = ""
                    End If
                    If TypeOf ControlChild Is DateTimePicker Then
                        ControlChild.Text = Now
                    End If
                    If TypeOf ControlChild Is UpDownBase Then
                        ControlChild.Text = 0
                    End If
                Next ControlChild
            End If
        Next ControlParent
        ControlChild = Nothing
        ControlParent = Nothing

    End Sub

    ' When you then query ComboBox1.SelectedValue this will return just the Value, without having to convert it back in to the data class.
    'Also the data class seams to work better if I defined the Value and Description as read-only properties.

    Public Class ComboValueDescription
        '** Class to store Control, Value & Description in a Combo Box
        Private uControl As Control
        Private uID As Object
        Private sValue As String

        Public ReadOnly Property Control() As Control
            Get
                Return uControl
            End Get
        End Property

        Public ReadOnly Property ID() As Object
            Get
                Return uID
            End Get
        End Property

        Public ReadOnly Property Value() As String
            Get
                Return sValue
            End Get
        End Property

        Public Sub New(ByVal NewControl As Control, ByVal NewID As Object, ByVal NewValue As String)
            uControl = NewControl
            uID = NewID
            sValue = NewValue
        End Sub

        Public Overrides Function ToString() As String
            Return sValue
        End Function

    End Class

    Public Class ListValueDescription
        '** Class to store Control, Value & Description in a List Box
        Private uControl As Control
        Private uID As Object
        Private sValue As String

        Public ReadOnly Property Control() As Control
            Get
                Return uControl
            End Get
        End Property

        Public ReadOnly Property ID() As Object
            Get
                Return uID
            End Get
        End Property

        Public ReadOnly Property Value() As String
            Get
                Return sValue
            End Get
        End Property

        Public Sub New(ByVal NewControl As Control, ByVal NewID As Object, ByVal NewValue As String)
            uControl = NewControl
            uID = NewID
            sValue = NewValue
        End Sub

        Public Overrides Function ToString() As String
            Return sValue
        End Function

    End Class

    Function ValCur(ByVal sValue As String) As Double
        '** Convert Currency Value to Double

        ' Error Checking
        On Error GoTo Err_ValCur

        ' Dimension Local Variables
        Dim iPos As Integer

        ' Check for Blank String
        If sValue = "" Then
            ValCur = 0
            Exit Function
        End If

        ' Remove Thousands Separator
        If InStr(sValue, ",") > 0 Then
            iPos = InStr(sValue, ",")
            sValue = sValue.Substring(0, iPos - 1) & sValue.Substring(iPos, Len(sValue) - iPos)
        End If

        ' Convert £ to Decimal Value / Check for Negative
        If InStr(sValue, "-") <> 0 Then
            If InStr(sValue, "£") <> 0 Or InStr(sValue, "¬") <> 0 Then
                ValCur = -Val(Trim(sValue).Substring(2, Len(Trim(sValue)) - 2))
            Else
                ValCur = -Val(sValue)
            End If
        Else
            If InStr(sValue, "£") <> 0 Or InStr(sValue, "¬") <> 0 Then
                ValCur = Val(Trim(sValue).Substring(1, Len(Trim(sValue)) - 1))
            Else
                ValCur = Val(sValue)
            End If
        End If

Err_ValCur:
        If Err.Number <> 0 Then
            sErrDescription = Err.Description
            WriteAuditLogRecord("modGlobal", "ValCur", "Error", sErrDescription)
            MsgBox("System Error occurred" & Chr(13) & "ValCur" & Chr(13) & sErrDescription, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AztecCRM - Error Reporting")
        End If

    End Function

End Module
