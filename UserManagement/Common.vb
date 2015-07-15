Imports System.Data.SqlClient
Imports System.IO

#Region " Enums "
Public Enum PageMode
    Initial = 1
    Postback = 2
    ReturnFromChild = 3
    CalledByOther = 4
End Enum
Public Enum RequestAction
    CreateNew = 1
    LoadExisting = 2
    SaveNew = 3
    SaveExisting = 4
    SaveNewOrExisting = 5
    ' ConfirmResults = 5
    NoSaveNew = 6
    NoSaveExisting = 7
    ReturnToParent = 8
    [Date] = 9
    Other = 10
End Enum
Public Enum ResponseAction
    DisplayBlank = 1
    DisplayUserInputNew = 2
    DisplayUserInputExisting = 3
    DisplayUserInputNewOrExisting = 4
    DisplayExisting = 5
    ReturnToCallingPage = 6
End Enum

Public Enum StringTreatEnum
    AsIs = 1
    SideQts = 2
    SecApost = 3
    SideQts_SecApost = 4
End Enum
Public Enum ReportTypeEnum
    Information = 1
    InformationNoLog = 2
    [Error] = 3
    ErrorNoShutdown = 4
    Timeout = 5
End Enum

#End Region

Public Class Results
    Private cSuccess As Boolean
    Private cMsg As String
    Private cResponseAction As ResponseAction
    Private cObtainConfirm As Boolean
    Public Property Success() As Boolean
        Get
            Return cSuccess
        End Get
        Set(ByVal Value As Boolean)
            cSuccess = Value
        End Set
    End Property
    Public Property Msg() As String
        Get
            Return cMsg
        End Get
        Set(ByVal Value As String)
            cMsg = Value
        End Set
    End Property
    Public Property ResponseAction() As ResponseAction
        Get
            Return cResponseAction
        End Get
        Set(ByVal Value As ResponseAction)
            cResponseAction = Value
        End Set
    End Property
    Public Property ObtainConfirm() As Boolean
        Get
            Return cObtainConfirm
        End Get
        Set(ByVal Value As Boolean)
            cObtainConfirm = Value
        End Set
    End Property
End Class

Public Class BulkResults
    Inherits Results
    Private cReportColl As New Collection
    Public ReadOnly Property ReportColl() As Collection
        Get
            Return cReportColl
        End Get
    End Property
End Class

Public Class Style
    Public Enum StyleType
        NormalEditable = 1
        NoneditableGrayed = 3
        NoneditableWhite = 2
        NotVisible = 4
    End Enum

    Public Shared Sub AddCustomStyle(ByVal tb As TextBox, ByRef Elements As Collection, ByVal Visible As Boolean, ByVal [ReadOnly] As Boolean)
        Dim StyleStr As String
        Dim i As Integer
        For i = 1 To Elements.Count
            StyleStr &= Elements(i).Key & ":" & Elements(i).Value & ";"
        Next
        tb.Attributes.Add("style", StyleStr)
        tb.Visible = Visible
        tb.ReadOnly = [ReadOnly]
    End Sub

    Public Shared Sub AddStyle(ByVal tb As TextBox, ByVal StyleType As StyleType, ByVal Width As Integer, Optional ByVal IsTextArea As Boolean = False)
        If Not IsTextArea Then
            Select Case StyleType
                Case StyleType.NormalEditable
                    tb.Attributes.Add("style", "width:" & Width & ";")
                    tb.Visible = True
                    tb.ReadOnly = False

                Case StyleType.NoneditableGrayed
                    tb.Attributes.Add("style", "width:" & Width & ";border-width:1px;background: #eeeedd;readOnly: true;")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NoneditableWhite
                    tb.Attributes.Add("style", "width:" & Width & ";border-width:1px;border-style:solid;border-color:cccccc;background: #ffffff")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NotVisible
                    tb.Visible = False
            End Select
        Else                                                                ' is a textarea
            Select Case StyleType
                Case StyleType.NormalEditable
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;")
                    tb.Visible = True
                    tb.ReadOnly = False

                Case StyleType.NoneditableGrayed
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:hidden;readOnly: true")
                    tb.Visible = True
                    tb.ReadOnly = True


                Case StyleType.NoneditableWhite
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;overflow:hidden;readOnly: true")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NotVisible
                    tb.Visible = False
            End Select
        End If

    End Sub
    Public Shared Sub AddStyle(ByVal tb As TextBox, ByVal IsVisible As Boolean, ByVal IsReadOnly As Boolean, ByVal Width As Integer, Optional ByVal IsTextArea As Boolean = False)
        tb.Visible = IsVisible
        tb.ReadOnly = IsReadOnly
        If Not IsTextArea Then
            If IsVisible Then
                If IsReadOnly Then
                    tb.Attributes.Add("style", "width:" & Width & ";border-width:1px;background: #eeeedd;readOnly: true;")
                    tb.ReadOnly = True
                Else
                    'tb.Attributes.Add("style", "width:" & Width & ";background: #ffffff;")
                    tb.Attributes.Add("style", "width:" & Width & ";")
                End If
            Else
                ' tb.Attributes.Add("style", "VISIBILITY: hidden")
                tb.Attributes.Add("style", "display: none")
            End If
        Else
            If IsVisible Then
                If IsReadOnly Then
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:hidden;readOnly: true")
                    tb.ReadOnly = True
                Else
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;")
                End If
            End If
        End If
    End Sub
End Class

Public Class Common
    Private cEnviro As Enviro
    Private cDBase As DBase

    Public Sub New()
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        cEnviro = SessionObj("Enviro")
        cDBase = New DBase
    End Sub

    Public Sub GenerateError()
        Dim a, b, c As Integer
        a = b / c
    End Sub

    Public Function GetNewRecordID(ByVal TableName As String, ByVal KeyFldName As String) As Integer
        Dim RandValue As Integer
        Dim dt As DataTable

        Try

            Do
                Do
                    Randomize()
                    RandValue = CType(Rnd() * 1000000, System.Int64)
                Loop Until RandValue > 99999
                dt = GetDT("SELECT Count (*) FROM " & TableName & " WHERE " & KeyFldName & " = " & RandValue)
                If dt.Rows(0)(0) = 0 Then
                    Exit Do
                End If
            Loop

            Return RandValue

        Catch ex As Exception
            Throw New Exception("Error #2210: Common GetNewRecordID. " & ex.Message, ex)
        End Try

    End Function

    Public Sub UpdateTimeSensitiveData(ByVal UserID As String, ByVal LicenseState As String)
        Dim sb As New System.Text.StringBuilder
        Dim Querypack As DBase.QueryPack

        Try

            sb.Append("UPDATE ua ")
            sb.Append("SET ua.StatusCode = 'T' ")
            sb.Append("FROM UserAppointments ua ")
            sb.Append("INNER JOIN UserLicenses ul ON (ua.UserID = ul.UserID AND ua.State = ul.State) ")
            sb.Append("WHERE ")
            sb.Append("ul.OKToRenewInd <> 1 AND ProjectReports.dbo.ufn_DateCompare(ul.ExpirationDate, getDate(), 1) = -1 ")
            If UserID <> Nothing AndAlso LicenseState <> Nothing Then
                sb.Append("AND ul.UserID = '" & UserID & "' AND ul.State = '" & LicenseState & "'")
            End If

            Querypack = ExecuteNonQueryWithQuerypack(sb.ToString)
            If Not Querypack.Success Then
                Throw New Exception("Error #2214a: Index UpdateTimeSensitiveData. " & Querypack.TechErrMsg)
            End If

        Catch ex As Exception
            Throw New Exception("Error #2214b: Index UpdateTimeSensitiveData. " & ex.Message)
        End Try
    End Sub

    Public Function GetSecurityWhereClause(ByVal Prefix As String, ByVal AccessLevel As RightsClass.AccessLevelEnum, ByVal LocationID As String) As String
        Dim Results As String
        Select Case AccessLevel
            Case RightsClass.AccessLevelEnum.AllAccess
                Results = String.Empty
            Case RightsClass.AccessLevelEnum.SupervisorAccess
                'Results = " (UserID = '" & cEnviro.LoggedInUserID & " ' or Role = 'ENROLLER') and LocationID = '" & LocationID & "'"
                Results = String.Empty
            Case RightsClass.AccessLevelEnum.EnrollerAccess
                If Prefix = Nothing Then
                    Results = "UserID = '" & cEnviro.LoggedInUserID & "'"
                Else
                    Results = Prefix & "." & "UserID = '" & cEnviro.LoggedInUserID & "'"
                End If
        End Select
        Return Results
    End Function

    Public Function StateRequiresStateSpecificLTCCert(ByVal State As String) As Boolean
        Dim dt As DataTable

        If State = String.Empty Then
            Return False
        End If

        dt = GetDT("SELECT LTCRequiresStateSpecificCert FROM Codes_State WHERE upper(StateCode) = '" & State.ToUpper & "'")
        If IsDBNull(dt.Rows(0)(0)) Then
            Return False
        Else
            Return dt.Rows(0)(0)
        End If
        dt.Dispose()
    End Function

    Public Sub DropdownFindByValueSelect(ByRef dd As DropDownList, ByVal Value As Object)
        Dim i As Integer
        Dim TestValue As String
        If IsDBNull(Value) OrElse Value = Nothing Then
            TestValue = String.Empty
        Else
            TestValue = Value
        End If
        TestValue = TestValue.ToLower

        For i = 0 To dd.Items.Count - 1
            If dd.Items(i).Value.ToLower = TestValue Then
                dd.Items(i).Selected = True
                Exit For
            End If
        Next
    End Sub

    Public Function Left(ByVal Value As Object, ByVal Length As Integer) As String
        Dim Results As String = String.Empty

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return Results
        Else
            If Value.length <= Length Then
                Return Value.ToString
            Else
                Return Value.Substring(0, Length)
            End If
        End If
    End Function

    Public Function ToProper(ByVal Value As Object) As String
        Dim i, j As Integer
        Dim Box As Object
        Dim Results As String = String.Empty
        Dim CurValue As String = String.Empty
        Dim NewValue As String

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return Results
        Else
            Box = Split(Value, " ")
            For i = 0 To Box.GetUpperBound(0)
                CurValue = Box(i)
                NewValue = String.Empty
                For j = 0 To CurValue.Length - 1
                    If j = 0 Then
                        NewValue = CurValue.Substring(0, 1).ToUpper
                    Else
                        NewValue &= CurValue.Substring(j, 1).ToLower
                    End If
                Next
                Results &= " " & NewValue
            Next
            Results = Trim(Results)
            Return Results
        End If
    End Function

    Public Sub PrintExcelSqlServer(ByVal Query As String, ByVal Path As String, ByVal FileName As String, ByVal Extension As String)
        Dim sb As New System.Text.StringBuilder

        Try
            sb.Append("exec usp_GenExcelDownloadFile ")
            sb.Append("@Query = '" & Query & "', ")
            sb.Append("@Path = '" & Path & "', ")
            sb.Append("@FileName = '" & FileName & "', ")
            sb.Append("@Extension = '" & Extension & "'")
            ExecuteNonQuery(sb.ToString)

        Catch ex As Exception
            Throw New Exception("Error #2207: Common PrintExcelSqlServer. " & ex.Message)
        End Try
    End Sub

    Public Sub PrintCSVVersionLocal(ByRef dt As DataTable, ByVal FullPath As String)
        Dim i As Integer
        Dim ReportHeader As String
        Dim sb As New System.Text.StringBuilder
        Dim ColNum As Integer
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim FirstRow As Boolean = True
        Dim RecCount As Integer = 0


        For i = 0 To dt.Rows.Count - 1

            ' ___ Header row.
            If FirstRow Then
                For ColNum = 0 To dt.Columns.Count - 1
                    If ColNum > 0 Then
                        sb.Append(",")
                    End If
                    sb.Append("""")
                    sb.Append(dt.Columns(ColNum).ColumnName)
                    sb.Append("""")
                Next
                sb.Append(vbCrLf)
                FirstRow = False
            End If

            ' ___ Data rows.
            For ColNum = 0 To dt.Columns.Count - 1
                If ColNum > 0 Then
                    sb.Append(",")
                End If
                sb.Append("""")
                sb.Append(dt.Rows(i)(ColNum).ToJSParm)
                sb.Append("""")
            Next
            sb.Append(vbCrLf)

        Next

        sb.Append(vbCrLf)

        If File.Exists(FullPath) Then
            File.Delete(FullPath)
        End If
        fs = File.Create(FullPath)
        fs.Close()

        sw = New StreamWriter(FullPath)
        sw.Write(sb.ToString)
        sw.Close()
        sw.Close()
        fs.Close()
    End Sub

    Public Function GetDownloadPath(ByVal Page As Page) As Collection
        Dim PathsColl As New Collection
        Dim DirPath As String
        Dim FileName As String
        Dim FullPath As String

        DirPath = Page.Request.ServerVariables("APPL_PHYSICAL_PATH") & "TempData\"
        'FileName = "Rpt" & DateDiff("s", "1/1/2000", Now) & "a" & CInt(Rnd() * 1000) & ".csv"
        FileName = "Rpt_" & cEnviro.LoggedInUserID & "_" & Date.Now.ToUniversalTime.AddHours(-5).ToString("yyyyMMdd_HHmmss_fff") & ".csv"
        FullPath = DirPath & FileName
        PathsColl.Add(DirPath, "DirPath")
        PathsColl.Add(FullPath, "AbsPath")
        PathsColl.Add("./TempData/" & FileName, "RelPath")
        Return PathsColl
    End Function

    Public Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String)
        Dim Report As New Report
        Report.SendEmail(SendTo, SendTo, cc, Subject, TextBody)
    End Sub

    Public Function GetPageCaption() As String
        Return "User Management v" & cEnviro.VersionNumber
    End Function

#Region " Data "
    Public Function ExecuteNonQueryWithQuerypack(ByVal Sql As String) As DBase.QueryPack
        Dim Querypack As DBase.QueryPack
        Querypack = ExecuteNonQueryWithQuerypack(cEnviro.DBHost, cEnviro.DefaultDatabase, Sql)
        Return Querypack
    End Function

    Public Function ExecuteNonQueryWithQuerypack(ByVal DBHost As String, ByVal Database As String, ByVal Sql As String) As DBase.QueryPack
        Dim Querypack As New DBase.QueryPack

        Try
            Dim SqlConnection1 As New SqlClient.SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
            SqlConnection1.Open()
            Dim SqlCmd As New System.Data.SqlClient.SqlCommand(Sql, SqlConnection1)
            SqlCmd.CommandType = System.Data.CommandType.Text
            SqlCmd.ExecuteNonQuery()
            SqlCmd.Dispose()
            SqlConnection1.Close()
            Querypack.Success = True
        Catch ex As Exception
            Querypack.Success = False
            Querypack.TechErrMsg = ex.Message
        End Try
        Return Querypack
    End Function


    Public Sub ExecuteNonQuery(ByVal Sql As String)
        ExecuteNonQuery(cEnviro.DBHost, cEnviro.DefaultDatabase, Sql)
    End Sub

    Public Sub ExecuteNonQuery(ByVal DBHost As String, ByVal Database As String, ByVal Sql As String)
        Dim SqlConnection1 As New SqlClient.SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
        SqlConnection1.Open()
        Dim SqlCmd As New System.Data.SqlClient.SqlCommand(Sql, SqlConnection1)
        SqlCmd.CommandType = System.Data.CommandType.Text
        SqlCmd.ExecuteNonQuery()
        SqlCmd.Dispose()
        SqlConnection1.Close()
    End Sub

    Public Function DoesTableExist(ByVal TableName As String) As Boolean
        Dim sb As New System.Text.StringBuilder
        Dim dt As DataTable

        sb.Append("USE " & cEnviro.DefaultDatabase & " ")
        sb.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        sb.Append(TableName)
        sb.Append("') ")
        sb.Append("SELECT Results = 1 ")
        sb.Append("ELSE ")
        sb.Append("SELECT Results = 0")

        dt = GetDT(sb.ToString, cEnviro.DBHost, cEnviro.DefaultDatabase)
        If dt.Rows(0)(0) = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DoesTableExist(ByVal DBHost As String, ByVal DBName As String, ByVal TableName As String) As Boolean
        Dim sb As New System.Text.StringBuilder
        Dim dt As DataTable

        sb.Append("USE " & DBName & " ")
        sb.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        sb.Append(TableName)
        sb.Append("') ")
        sb.Append("SELECT Results = 1 ")
        sb.Append("ELSE ")
        sb.Append("SELECT Results = 0")

        dt = GetDT(sb.ToString, DBHost, DBName)
        If dt.Rows(0)(0) = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DoesDatabaseAndTableExist(ByVal DBName As String, ByVal TableName As String) As Boolean
        Return DoesDatabaseAndTableExist(cEnviro.DBHost, DBName, TableName)
    End Function

    Public Function DoesDatabaseAndTableExist(ByVal DBHost As String, ByVal DBName As String, ByVal TableName As String) As Boolean
        Dim Sql As String
        Dim Querypack As DBase.QueryPack

        Sql = DoesTableExistSql(DBName, TableName)
        Querypack = GetDTWithQueryPack(Sql, DBHost, DBName)
        If Querypack.Success Then
            If Querypack.dt.rows(0)(0) = 1 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

        'Sql = DoesTableExistSql(DBName, TableName)
        'Dim CmdAsst As New CmdAsst(CommandType.Text, Sql)
        'Dim QueryPack As DBase.QueryPack
        'QueryPack = CmdAsst.Execute
        'If QueryPack.Success Then
        '    If QueryPack.dt.rows(0)(0) = 1 Then
        '        Return True
        '    Else
        '        Return False
        '    End If
        'Else
        '    Return False
        'End If
    End Function

    Public Function DoesTableExistSql(ByVal DBName As String, ByVal TableName As String) As String
        Dim sb As New System.Text.StringBuilder

        sb.Append("USE " & DBName & " ")
        sb.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        sb.Append(TableName)
        sb.Append("') ")
        sb.Append("SELECT Results = 1 ")
        sb.Append("ELSE ")
        sb.Append("SELECT Results = 0")
        Return sb.ToString
    End Function

    Public Function GetFieldList(ByVal TableName As String) As DataTable
        Return GetFieldList(cEnviro.DBHost, cEnviro.DefaultDatabase, TableName)
    End Function

    Public Function GetFieldList(ByVal DBHost As String, ByVal Database As String, ByVal TableName As String) As DataTable
        Dim dt As DataTable
        Dim Sql As New System.Text.StringBuilder

        'Sql.append("SELECT * from information_schema.tables where table_type='BASE TABLE'")

        Sql.Append("SELECT column_name,ordinal_position,column_default,data_type, ")
        Sql.Append("Is_nullable from information_schema.columns ")
        Sql.Append("WHERE table_name='" & TableName & "'")
        Return GetDT(Sql.ToString, DBHost, Database)
    End Function

    Public Function GetExtendedTable(ByRef dt As DataTable) As DataTable
        Return cDBase.GetDTExtended(dt)
    End Function


    Public Function GetDT(ByVal Sql As String, Optional ByVal ExtendedTbl As Boolean = False) As DataTable
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable

        Dim SqlCmd As New SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString)
        DataAdapter = New SqlDataAdapter(SqlCmd)
        DataAdapter.Fill(dt)
        If ExtendedTbl Then
            dt = cDBase.GetDTExtended(dt)
        End If
        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return dt
    End Function

    Public Function GetDT(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, Optional ByVal ExtendedTbl As Boolean = False) As DataTable
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable

        Dim SqlCmd As New SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
        DataAdapter = New SqlDataAdapter(SqlCmd)
        DataAdapter.Fill(dt)
        If ExtendedTbl Then
            dt = cDBase.GetDTExtended(dt)
        End If
        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return dt
    End Function

    Public Function GetDTWithQueryPack(ByVal Sql As String, Optional ByVal ExtendedTbl As Boolean = False) As DBase.QueryPack
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable
        Dim QueryPack As New DBase.QueryPack

        Dim SqlCmd As New SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString())
        DataAdapter = New SqlDataAdapter(SqlCmd)
        Try
            DataAdapter.Fill(dt)
            If ExtendedTbl Then
                dt = cDBase.GetDTExtended(dt)
            End If
            QueryPack.Success = True
            QueryPack.dt = dt
        Catch ex As Exception
            QueryPack.Success = False
            QueryPack.TechErrMsg = ex.Message
        End Try

        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return QueryPack
    End Function

    Public Function GetDTWithQueryPack(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, Optional ByVal ExtendedTbl As Boolean = False) As DBase.QueryPack
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable
        Dim QueryPack As New DBase.QueryPack

        Dim SqlCmd As New SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
        DataAdapter = New SqlDataAdapter(SqlCmd)
        Try
            DataAdapter.Fill(dt)
            If ExtendedTbl Then
                dt = cDBase.GetDTExtended(dt)
            End If
            QueryPack.Success = True
            QueryPack.dt = dt
        Catch ex As Exception
            QueryPack.Success = False
            QueryPack.TechErrMsg = ex.Message
        End Try

        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return QueryPack
    End Function

    Public Function GetDTSqlDataElements(ByVal Sql As String, ByVal TableName As String) As DataTable
        Dim dt2 As DataTable
        dt2 = GetDTSqlDataElements(Sql, cEnviro.DBHost, cEnviro.DefaultDatabase, TableName)
        Return dt2
    End Function

    Public Function GetDTSqlDataElements(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal TableName As String) As DataTable
        Dim i, j As Integer
        Dim Row As Integer
        Dim Col As Integer
        Dim dt As DataTable
        Dim dt2 As New DataTable
        Dim dr2 As DataRow
        Dim DictItem As DBase.DictItem
        Dim DictColl As Collection
        Dim Found As Boolean
        Dim IdxColl As New Collection

        ' ___ Get the original table
        dt = GetDT(Sql, DBHost, Database)

        ' ___ Add the column definitions to the dictionary if not already present
        DictColl = cDBase.DictColl
        For i = 1 To DictColl.Count
            Found = False
            For j = 0 To dt.Columns.Count - 1
                If DictColl(i).Database = Database AndAlso DictColl(i).Tablename = TableName Then
                    Found = True
                    Exit For
                End If
            Next
        Next
        If Not Found Then
            cDBase.DictionaryAddTable(Me, TableName)
        End If

        ' ___ Build the index collection
        For Col = 0 To dt.Columns.Count - 1
            For j = 1 To DictColl.Count
                If DictColl(j).Database = Database AndAlso DictColl(j).Tablename = TableName AndAlso DictColl(j).ColumnName = dt.Columns(Col).ColumnName Then
                    IdxColl.Add(j)
                    Exit For
                End If
            Next
        Next

        ' ___ Prepare the new table
        For Col = 0 To dt.Columns.Count - 1
            dt2.Columns.Add(dt.Columns(Col).ColumnName, GetType(DBase.SqlDataElement))
        Next

        For Row = 0 To dt.Rows.Count - 1
            dr2 = dt2.NewRow
            For Col = 0 To dt.Columns.Count - 1

                ' ___ Construct the SqlDataElement with the dictionary item and field value
                Dim SqlDataElement As New DBase.SqlDataElement(DictColl(IdxColl(Col + 1)), dt.Rows(Row)(Col))
                dr2(Col) = SqlDataElement

            Next
            dt2.Rows.Add(dr2)
        Next
        Return dt2
    End Function

#End Region

#Region " Page handling "
    Public Function GetPageMode(ByVal Page As Page, ByVal Sess As PageSession) As PageMode
        Dim PageMode As PageMode

        If Page.IsPostBack AndAlso Page.Request.Form("__EVENTTARGET") = "" Then
            PageMode = PageMode.Postback
        Else
            Select Case Page.Request.QueryString("CalledBy")
                Case "Child"
                    PageMode = PageMode.ReturnFromChild
                Case "Other"
                    If Sess.PageInitiallyLoaded Then
                        PageMode = PageMode.CalledByOther
                    Else
                        PageMode = PageMode.Initial
                        Sess.PageInitiallyLoaded = True
                    End If
                Case Else
                    PageMode = PageMode.Initial
                    Sess.PageInitiallyLoaded = True
            End Select
        End If
        Return PageMode
    End Function

    Public Function GetRequestAction(ByVal Page As Page) As RequestAction
        Dim ActionType As String
        Dim RequestAction As RequestAction
        Dim hdResponseAction As String
        Dim hdAction As String
        Dim CallType As String

        If Page.Request.QueryString("CallType") = Nothing OrElse Page.Request.QueryString("CallType") = String.Empty Then
            CallType = String.Empty
        Else
            CallType = Page.Request.QueryString("CallType")
            CallType = CallType.ToLower
        End If

        If Page.Request.Form("hdAction") = Nothing OrElse Page.Request.Form("hdAction") = String.Empty Then
            hdAction = String.Empty
        Else
            hdAction = Page.Request.Form("hdAction")
            hdAction = hdAction.ToLower
        End If

        If Not Page.IsPostBack Then
            ActionType = "record"
        Else
            If hdAction = "update" Then
                ActionType = "record"
            ElseIf hdAction = "return" Then
                ActionType = "return"
            ElseIf hdAction = "confirmation" Then
                ActionType = "confirmation"
            ElseIf hdAction = "clientselectionchanged" Then
                ActionType = "clientselectionchanged"
            End If
        End If

        Select Case ActionType
            Case "return"
                RequestAction = RequestAction.ReturnToParent

            Case "record"
                If Not Page.IsPostBack Then
                    Select Case CallType
                        Case "new", ""
                            RequestAction = RequestAction.CreateNew
                        Case "existing"
                            RequestAction = RequestAction.LoadExisting
                    End Select
                Else
                    Select Case Page.Request.Form("hdResponseAction")
                        Case ResponseAction.DisplayBlank.ToString
                            RequestAction = RequestAction.SaveNew
                        Case ResponseAction.DisplayExisting.ToString
                            RequestAction = RequestAction.SaveExisting
                        Case ResponseAction.DisplayUserInputNew.ToString
                            RequestAction = RequestAction.SaveNew
                        Case ResponseAction.DisplayUserInputExisting.ToString
                            RequestAction = RequestAction.SaveExisting
                        Case ResponseAction.DisplayUserInputNewOrExisting.ToString
                            RequestAction = RequestAction.SaveNewOrExisting
                    End Select
                End If

            Case "confirmation"
                hdResponseAction = Page.Request.Form("hdResponseAction")
                If hdResponseAction = "DisplayBlank" Or hdResponseAction = "DisplayUserInputNew" Then
                    If Page.Request.Form("hdConfirm") = "yes" Then
                        RequestAction = RequestAction.SaveNew
                    Else
                        RequestAction = RequestAction.NoSaveNew
                    End If
                ElseIf hdResponseAction = "DisplayExisting" Or hdResponseAction = "DisplayUserInputExisting" Then
                    If Page.Request.Form("hdConfirm") = "yes" Then
                        RequestAction = RequestAction.SaveExisting
                    Else
                        RequestAction = RequestAction.NoSaveExisting
                    End If
                End If

            Case "clientselectionchanged"
                RequestAction = RequestAction.Other

        End Select

        Return RequestAction
    End Function

    Public Function GetResponseActionFromRequestActionOther(ByRef Page As Page) As ResponseAction
        Select Case Page.Request("hdResponseAction").ToString
            Case ResponseAction.DisplayBlank.ToString
                Return ResponseAction.DisplayUserInputNew
            Case ResponseAction.DisplayExisting.ToString
                Return ResponseAction.DisplayUserInputExisting
            Case ResponseAction.DisplayUserInputExisting.ToString
                Return ResponseAction.DisplayUserInputExisting
            Case ResponseAction.DisplayUserInputNew.ToString
                Return ResponseAction.DisplayUserInputNew
            Case ResponseAction.ReturnToCallingPage.ToString
                Return ResponseAction.ReturnToCallingPage
        End Select
    End Function
#End Region

#Region " In handlers "
    Public Function StrInHandler(ByVal Input As Object) As Object
        Dim Output As Object

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf (Not IsNumeric(Input)) AndAlso Input = Nothing Then
            Return String.Empty
            'ElseIf (Not IsDate(Input)) AndAlso Input.length = 0 Then
            '    Return String.Empty
        Else
            Output = Input
            If Input = Nothing Then
                Return String.Empty
            End If
            Return Output
        End If
    End Function

    Public Function DateInHandler(ByVal Input As Object) As Object
        ' 12/31/2399
        Dim Output As Object
        Output = Input

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf Input = "01/01/1900" Then
            Return String.Empty
        ElseIf Input = "01/01/1950" Then
            Return String.Empty
        Else
            Return Output
        End If
    End Function

    Public Function NumInHandler(ByVal Input As Object, ByVal NullAsZero As Boolean) As Object
        If IsDBNull(Input) Then
            If NullAsZero Then
                Return 0
            Else
                Return String.Empty
            End If
        Else
            Return Input
        End If
    End Function

    Public Function GuidInHandler(ByVal Input As Object) As Object
        If IsDBNull(Input) Then
            Return String.Empty
        Else
            Return Input.ToString
        End If
    End Function


    Public Function StrXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        Dim Output As Object
        Dim ReturnNull As Boolean

        If IsDBNull(Input) Then
            ReturnNull = True
        ElseIf (Not IsNumeric(Input)) AndAlso Input = Nothing Then
            ReturnNull = True
        Else
            Output = Replace(Input, "~", "'")
            If Output = Nothing Then
                ReturnNull = True
            End If
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return String.Empty
            End If
        Else
            Return Output
        End If

    End Function
    Public Function DateXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        ' 12/31/2399
        Dim Output As Object
        Dim ReturnNull As Boolean
        Output = Input

        If IsDBNull(Input) OrElse Input = Nothing Then
            ReturnNull = True
        Else
            Output = Input
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return "1/1/1950"
            End If
        Else
            Return Output
        End If
    End Function

    Public Function NumXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        Dim Output As Object
        Dim ReturnNull As Boolean

        If IsDBNull(Input) Then
            ReturnNull = True
        Else
            Output = Input
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return 0
            End If
        Else
            Return Output
        End If

    End Function
#End Region

#Region " Out handlers"
    'Public Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
    '    Dim ReturnNull As Boolean
    '    Dim Output As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input.length > 0 Then
    '        Output = Replace(Input, "'", "~")
    '    Else
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    End If

    '    If ReturnNull Then
    '        Return "null"
    '    Else
    '        If AddSingleQuotes Then
    '            Return "'" & Output & "'"
    '        Else
    '            Return Output
    '        End If
    '    End If
    'End Function

    'Public Function OrigStrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
    '    Dim ReturnNull As Boolean
    '    Dim Output As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input.length > 0 Then
    '        'Output = Replace(Input, "'", "~")

    '        If AddSingleQuotes Then
    '            Output = Replace(Input, "'", "''")
    '        Else
    '            Output = Input
    '        End If

    '    Else
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    End If

    '    If ReturnNull Then
    '        Return "null"
    '    Else
    '        If AddSingleQuotes Then
    '            Return "'" & Output & "'"
    '        Else
    '            Return Output
    '        End If
    '    End If
    'End Function

    Public Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal StringTreat As StringTreatEnum) As Object
        Dim Output As String

        If IsDBNull(Input) Then
            If AllowNull Then
                Output = "null"
            Else
                Output = String.Empty
            End If
        ElseIf Input Is Nothing Then
            If AllowNull Then
                Output = "null"
            Else
                Output = String.Empty
            End If
        ElseIf Input.length > 0 Then
            Output = Input
            If (StringTreat = StringTreatEnum.SecApost) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
                Output = Replace(Output, "'", "''")
            End If
            If (StringTreat = StringTreatEnum.SideQts) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
                Output = "'" & Output & "'"
            End If
        Else
            If AllowNull Then
                Output = "null"
            Else
                Output = String.Empty
                If (StringTreat = StringTreatEnum.SideQts) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
                    Output = "'" & Output & "'"
                End If
            End If
        End If
        Return Output
    End Function

    Public Function DateOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim ReturnNull As Boolean
        Dim Output As Object

        If IsDBNull(Input) OrElse Input = Nothing Then
            If AllowNull Then
                ReturnNull = True
            Else
                Output = "01/01/1950"
            End If
        Else
            Output = Input
        End If

        If ReturnNull Then
            Return "null"
        Else
            If AddSingleQuotes Then
                Return "'" & Output & "'"
            Else
                Return Output
            End If
        End If
    End Function
    Public Function PhoneOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim i As Integer
        Dim Output As String = String.Empty
        Dim Working As String
        Working = StrOutHandler(Input, AllowNull, StringTreatEnum.SideQts)

        If Working = "null" Or Working = String.Empty Then
        Else
            If Working.Length >= 10 Then
                For i = 0 To Working.Length - 1
                    If IsNumeric(Working.Substring(i, 1)) Then
                        Output &= Working.Substring(i, 1)
                    End If
                Next
            End If
        End If

        If Output.Length = 10 Then
            Output = InsertAt(Output, "(", 1)
            Output = InsertAt(Output, ") ", 5)
            Output = InsertAt(Output, "-", 10)
        Else
            Output = Input
        End If

        If AddSingleQuotes Then
            Output = "'" & Output & "'"
        End If

        Return Output
    End Function

    Public Function BoolToBit(ByVal Input As Object) As Integer
        If IsDBNull(Input) Then
            Return 0
        Else
            If CType(Input, Boolean) Then
                Return 1
            Else
                Return 0
            End If
        End If
    End Function

    Public Function BoolToText(ByVal Input As Object) As String
        If IsDBNull(Input) Then
            Return "No"
        Else
            If CType(Input, Boolean) Then
                Return "Yes"
            Else
                Return "No"
            End If
        End If
    End Function

    Public Function BitOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        If IsDBNull(Input) Then
            If AllowNull Then
                Return "null"
            Else
                Return 0
            End If
        Else
            If CType(Input, Boolean) Then
                Return 1
            Else
                Return 0
            End If
        End If
    End Function
#End Region

#Region " Validate "
    Public Function IsAlphaNumeric(ByVal strToCheck As String) As Boolean
        'Dim pattern As System.Text.RegularExpressions.Regex = New Regex("[^a-zA-Z0-9]")
        Dim Pattern As System.Text.RegularExpressions.Regex
        Pattern = New System.Text.RegularExpressions.Regex("[^a-zA-Z0-9]")
        Return Not Pattern.IsMatch(strToCheck)
    End Function

    Public Function IsBlank(ByVal Value As Object) As Boolean
        If IsDBNull(Value) Then
            Return True
        ElseIf Value = Nothing Then
            Return True
        Else
            If Value.length = 0 Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Sub ValidateStringField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal MinLength As Integer, ByVal ErrMsg As String)
        If Value.length < MinLength Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateStringField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal MinLength As Integer, ByVal MaxLength As Integer, ByVal ErrMsg As String)
        If Value.length < MinLength Or Value.length > MaxLength Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateApostrophe(ByRef ErrColl As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If InStr(Value, "'") > 0 Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateNumericField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim PassTest As Boolean
        If IsDBNull(Value) OrElse Value.Length = 0 Then
            If AllowNull Then
                PassTest = True
            Else
                PassTest = False
            End If
        Else
            If IsNumeric(Value) Then
                PassTest = True
            Else
                PassTest = False
            End If
        End If

        If Not PassTest Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateRadio(ByRef ErrColl As Collection, ByVal SelectedIndex As Integer, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        If (SelectedIndex < 0) AndAlso (Not AllowNull) Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateNumericRange(ByRef ErrColl As Collection, ByVal Value As Object, ByVal Min As Integer, ByVal Max As Integer, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim PassTest As Boolean
        If IsDBNull(Value) OrElse Value.Length = 0 Then
            If AllowNull Then
                PassTest = True
            Else
                PassTest = False
            End If
        Else
            If IsNumeric(Value) Then
                If Value >= Min AndAlso Value <= Max Then
                    PassTest = True
                Else
                    PassTest = False
                End If
            End If
        End If

        If Not PassTest Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateDateField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim Valid As Boolean
        If IsDBNull(Value) OrElse Value = Nothing Then
            If AllowNull Then
                Valid = True
            Else
                Valid = False
            End If
        ElseIf IsDate(Value) Then
            Valid = True
        Else
            Valid = False
        End If
        If Not Valid Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Function IsValidPhoneNumber(ByVal Value As Object) As Boolean
        Dim i As Integer
        Dim NumCount As Integer

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return False
        End If

        If Value.length >= 10 Then
            For i = 0 To Value.Length - 1
                If IsNumeric(Value.Substring(i, 1)) Then
                    NumCount += 1
                End If
            Next
        End If

        If NumCount = 10 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub ValidatePhoneNumber(ByRef Errcoll As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If Not IsValidPhoneNumber(Value) Then
            If Errcoll.Count = 0 Then
                Errcoll.Add(ErrMsg)
            Else
                Errcoll.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateEmailAddress(ByRef ErrColl As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If Not IsValidEmailAddress(Value) Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Function IsValidEmailAddress(ByVal Value As Object) As Boolean
        Dim OKSoFar As Boolean = True
        Const InvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
        Dim i As Integer
        Dim Num As Integer
        Dim DotPos As Integer
        Dim Part2 As String

        ' ___ Check for null or empty value
        If IsDBNull(Value) OrElse Value = Nothing Then
            OKSoFar = False
        End If

        ' ___ Check for minimum length
        If OKSoFar Then
            If Value.Length < 5 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for a double quote
        If OKSoFar Then
            OKSoFar = Not InStr(1, Value, Chr(34)) > 0  'Check to see if there is a double quote
        End If

        ' ___ Check for consecutive dots
        If OKSoFar Then
            OKSoFar = Not InStr(1, Value, "..") > 0
        End If

        ' ___ Check for invalid characters
        If OKSoFar Then
            For i = 0 To InvalidChars.Length - 1
                If InStr(1, Value, InvalidChars.Substring(i, 1)) > 0 Then
                    OKSoFar = False
                    Exit For
                End If
            Next
        End If

        ' ___ Check for number of @ symbols
        If OKSoFar Then
            For i = 0 To Value.Length - 1
                If InStr(Value.Substring(i, 1), "@") > 0 Then
                    Num += 1
                End If
            Next
            If Num > 1 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for the @ symbol in starting before the third position
        If OKSoFar Then
            If InStr(Value, "@") < 2 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for number of dots
        If OKSoFar Then
            Num = 0
            Part2 = Value.substring(InStr(Value, "@"))
            For i = 0 To Part2.Length - 1
                If InStr(Part2.Substring(i, 1), ".") > 0 Then
                    Num += 1
                End If
            Next
            If Num > 1 Then
                OKSoFar = False
            End If
        End If

        ' ___ Dot is present and not immediately after ampersand and not at end. 
        '___  Dot separated from ampersand by at least one character
        If OKSoFar Then
            DotPos = InStr(Part2, ".")
            If DotPos < 2 Or DotPos = Part2.Length Then
                OKSoFar = False
            End If
        End If

        Return OKSoFar
    End Function

    Public Sub ValidateDropDown(ByRef ErrColl As Collection, ByRef dd As DropDownList, ByVal MinSelectedIndex As Integer, ByVal ErrMsg As String)
        If dd.SelectedIndex < MinSelectedIndex Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateDropDownSelect0(ByRef ErrColl As Collection, ByRef dd As DropDownList, ByVal ErrMsg As String)
        If dd.SelectedIndex > 0 Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateCheckbox(ByRef ErrColl As Collection, ByRef chkBox As CheckBox, ByVal ValidState As Integer, ByVal ErrMsg As String)
        Dim IsValid As Boolean = True
        If ValidState = 0 AndAlso Not chkBox.Checked Then
            IsValid = False
        ElseIf ValidState = 1 AndAlso Not chkBox.Checked Then
            IsValid = False
        End If
        If Not IsValid Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateErrorOnly(ByRef ErrColl As Collection, ByVal ErrMsg As String)
        If ErrColl.Count = 0 Then
            ErrColl.Add(ErrMsg)
        Else
            ErrColl.Add(", " & ErrMsg)
        End If
    End Sub

    Public Function ErrCollToString(ByRef ErrColl As Collection, ByVal Intro As String) As String
        Dim sb As New System.Text.StringBuilder
        Dim i As Integer
        If ErrColl.Count > 0 Then
            For i = 1 To ErrColl.Count
                sb.Append(ErrColl(i))
            Next
        End If
        Return Intro & " " & sb.ToString & "."
    End Function
#End Region

#Region " This to that "
    Public Function BitToRadio(ByVal Value As Object, ByVal TrueIndex As Integer, ByVal AllowNoneSelected As Boolean) As Integer
        Dim FalseIndex As Integer
        FalseIndex = System.Math.Abs(TrueIndex - 1)

        If IsDBNull(Value) Then
            If AllowNoneSelected Then
                Return -1
            Else
                Return FalseIndex
            End If
        Else
            If Value Then
                Return TrueIndex
            Else
                Return FalseIndex
            End If
        End If
    End Function

    Public Function BitToString(ByVal Value As Object, ByVal TrueString As String, ByVal FalseString As String, ByVal AllowNull As Boolean) As String
        If IsDBNull(Value) Then
            If AllowNull Then
                Return String.Empty
            Else
                Return FalseString
            End If
        End If
        If Value Then
            Return TrueString
        Else
            Return FalseString
        End If
    End Function

    Public Function ChkToInd(ByVal chkBox As CheckBox) As Integer
        If chkBox.Checked Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Public Sub IndToChk(ByVal Ind As Object, ByVal chkBox As CheckBox)
        If IsDBNull(Ind) Then
            chkBox.Checked = False
        Else
            If Ind Then
                chkBox.Checked = True
            Else
                chkBox.Checked = False
            End If
        End If
    End Sub
#End Region

#Region " Everything else "

    Public Sub RecordLoggedInUserData(ByVal LoggedInUserID As String, ByVal SessionID As String, ByVal LastLoginIP As String)
        Dim dt As DataTable
        Dim Sql As String

        dt = GetDT("SELECT LastSessionID FROM Users WHERE UserID = '" & LoggedInUserID & "'")
        If dt.Rows(0)("LastSessionID") <> SessionID Then
            Sql = "UPDATE Users SET LastSessionID = '" & SessionID & "', LastLoginDate = '" & GetServerDateTime() & "', LastLoginIP = '" & LastLoginIP & "' WHERE UserID = '" & LoggedInUserID & "'"
            ExecuteNonQuery(Sql)
        End If
    End Sub

    Public Function NumPart(ByVal Value As String) As Integer
        Dim i As Integer
        Dim Result As String
        For i = 0 To Value.Length - 1
            If Asc(Value.Substring(i, 1)) > 47 And Asc(Value.Substring(i, 1)) < 58 Then
                Result &= Value.Substring(i, 1)
            End If
        Next
        If Result = Nothing Then
            Return -1
        Else
            Return CType(Result, System.Int64)
        End If
    End Function

    Public Function FormatCalendar(ByRef Calendar1 As System.Web.UI.WebControls.Calendar)
        With Calendar1
            .Font.Name = "Verdana;Arial"
            .NextPrevFormat = NextPrevFormat.ShortMonth
            .Font.Size = System.Web.UI.WebControls.FontUnit.Point(8)
            .DayHeaderStyle.BackColor = Color.LightSkyBlue
            .DayStyle.BackColor = Color.White
            .OtherMonthDayStyle.ForeColor = Color.Gray
            .OtherMonthDayStyle.BackColor = Color.White
            .TitleStyle.BackColor = Color.LightSkyBlue
            .TitleStyle.ForeColor = Color.Black
            .TitleStyle.Font.Bold = True
            .SelectedDayStyle.BackColor = Color.Navy
            .SelectedDayStyle.Font.Bold = False
            ' .SelectedDate = Common.GetServerDateTime
        End With
    End Function

    Public Function GetServerDateTime() As DateTime
        Return Date.Now.ToUniversalTime.AddHours(-5)
        'Dim CmdAsst As New CmdAsst(CommandType.StoredProcedure, "ServerDateTime")
        'Dim QueryPack As CmdAsst.QueryPack
        'Try
        '    QueryPack = CmdAsst.Execute
        'Catch
        '    QueryPack.Success = False
        'End Try
        'Return QueryPack.dt.rows(0)("ServerDateTime")
    End Function

    Public Function ConditionStringForHTML(ByVal Value As Object) As String
        Dim Results As String
        If IsDBNull(Value) Then
            Results = String.Empty
        Else
            Results = Value.ToString
        End If
        Results = Replace(Results, Chr(10).ToString, "<br />")
        Return Results
    End Function

    Public Function GetRightsStr(ByRef dt As DataTable) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        If dt.Rows.Count = 0 Then
            Return String.Empty
        Else
            For i = 0 To dt.Rows.Count - 1
                sb.Append("|" & StrInHandler(dt.Rows(i)("RightCd")))
            Next
            Return sb.ToString.Substring(1)
        End If
    End Function

    Public Function InsertAt(ByVal Value As String, ByVal InsChar As String, ByVal Pos As Integer) As String
        Dim ValuePos As Integer = 1
        Dim Output As String = String.Empty
        Dim OutputPos As Integer = 1
        Do
            If OutputPos = Pos Then
                Output &= InsChar
                OutputPos += 1
            Else
                Output &= Value.Substring(ValuePos - 1, 1)
                ValuePos += 1
                OutputPos += 1
                If ValuePos > Value.Length Then
                    Exit Do
                End If
            End If
        Loop
        Return Output
    End Function

    Public Function ToNull(ByVal Input As Object) As Object
        If IsDBNull(Input) Then
            Return DBNull.Value
        ElseIf Input Is Nothing Then
            Return DBNull.Value
        ElseIf Input.length = 0 Then
            Return DBNull.Value
        Else
            Return Input
        End If
    End Function



    'Public Function DateSqlWhere(ByRef Input As Object) As String
    '    If IsDBNull(Input) Then
    '        Return " is null "
    '    ElseIf Input = Nothing Then
    '        Return " is null "
    '    ElseIf Trim(Input).Length = 0 Then
    '        Return " is null "
    '    Else
    '        Return " = '" & Input & "' "
    '    End If
    'End Function

    'Public Function DateSqlWhereNoNull(ByRef Input As Object) As String
    '    If IsDBNull(Input) Then
    '        Return " = '01/01/1900' "
    '    ElseIf Input = Nothing Then
    '        Return " = '01/01/1900' "
    '    ElseIf Trim(Input).Length = 0 Then
    '        Return " = '01/01/1900' "
    '    Else
    '        Return " = '" & Input & "' "
    '    End If
    'End Function

    Public Function IsBVIDate(ByVal Input As Object) As Boolean
        If IsDBNull(Input) Then
            Return False
        ElseIf Input = Nothing Then
            Return False
        ElseIf Input = "01/01/1950" Then
            Return False
        ElseIf Input.ToString = String.Empty Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function GetRandomlyGeneratedPassword(ByVal Length As Integer) As String
        Dim i As Integer
        Dim RndValue As Integer
        Dim sb As New System.Text.StringBuilder

        ' ___ Generate random password 8 digits long
        For i = 1 To Length
            Randomize()
            RndValue = CInt(Int(62 * Rnd() + 1))
            Select Case RndValue
                Case 1 To 10
                    sb.Append(Chr(RndValue + 47))
                Case 11 To 36
                    sb.Append(Chr(RndValue + 54))
                Case 37 To 62
                    sb.Append(Chr(RndValue + 28))
            End Select
        Next
        Return sb.ToString.ToLower

    End Function

    Public Function GetCurRightsHidden(ByVal RightsColl As Collection) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        For i = 1 To RightsColl.Count
            sb.Append(RightsColl(i) & "|")
        Next
        sb.Length -= 1
        Return "<input type='hidden' id='currentrights' name='currentrights' value=""" & sb.ToString & """>"
    End Function

    Public Function GetCurRightsHidden(ByVal RightsStr As String) As String
        Return "<input type='hidden'id='currentrights' name='currentrights' value=""" & RightsStr & """ > "
    End Function
#End Region

End Class

'Public Class CmdAsst
'    Private cSqlCmd As SqlClient.SqlCommand
'    Private cDBName As String

'    Public Sub New(ByVal CmdType As CommandType, ByRef SPNameOrSql As String, Optional ByVal DBName As String = Nothing)
'        cSqlCmd = New SqlClient.SqlCommand(SPNameOrSql)
'        cSqlCmd.CommandType = CmdType
'        cDBName = DBName
'    End Sub

'#Region " Add Parameters "
'    Public Sub AddInt(ByVal VarName As String, ByVal Value As Object)
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.ParameterDirection.Input, False, CType(10, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddBit(ByVal VarName As String, ByVal Value As Object)
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Bit, 1"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddDateTime(ByVal VarName As String, ByVal Value As Object) ' datetime
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.DateTime, 8"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddMoney(ByVal VarName As String, ByVal Value As Object) ' decimal
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Money, 4, System.Data.ParameterDirection.Input, False, CType(19, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddVarChar(ByVal VarName As String, ByVal Value As Object, ByVal Length As Integer)
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.VarChar, "))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddFloat(ByVal VarName As String, ByVal Value As Object) ' double
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Float, 8, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'#End Region

'    Public Function Execute() As DBase.QueryPack
'        Dim QueryPack As New DBase.QueryPack
'        Dim DataAdapter As SqlDataAdapter
'        Dim dt As New DataTable

'        If cDBName = Nothing Then
'            cSqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString)
'        Else
'            cSqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString(cEnviro.DBHost, cDBName))
'        End If
'        DataAdapter = New SqlDataAdapter(cSqlCmd)
'        Try
'            DataAdapter.Fill(dt)
'            QueryPack.Success = True
'            QueryPack.dt = dt
'        Catch ex As Exception
'            QueryPack.Success = False
'            QueryPack.TechErrMsg = ex.Message
'        End Try
'        DataAdapter.Dispose()
'        cSqlCmd.Dispose()
'        Return QueryPack
'    End Function
'End Class

Public Class KeyValueObj
    Private cKey As String
    Private cValue As String

    Public Property Key()
        Get
            Return cKey
        End Get
        Set(ByVal Value)
            cKey = Value
        End Set
    End Property
    Public Property Value()
        Get
            Return cValue
        End Get
        Set(ByVal Value)
            cValue = Value
        End Set
    End Property
    Public Sub New(ByVal Key As String, ByVal Value As String)
        cKey = Key
        cValue = Value
    End Sub
End Class