Imports System.Data.SqlClient
Imports System.Runtime.InteropServices.Marshal
Imports Microsoft.Office.Interop
'http://www.aspnetpro.com/NewsletterArticle/2003/09/asp200309so_l/asp200309so_l.asp


Public Class Excel
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents dgRoleID As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dgStatusCode As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dgClientID As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dgSiteID As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dgResidentState As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dgWorkState As System.Web.UI.WebControls.DataGrid

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private cCommon As Common

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cCommon = New Common
        '  Codes_Role()
        ' Codes_CarrierID()
        'Codes_LocationID()
        ' Codes_State()
        'Codes_CompanyID()
        'Codes_Status()
        Codes_ClientID()

        ' CodeTableEntries()
        Exit Sub

        'Dim dt As DataTable
        'dt = GetUserTable()
        'CodeTableEntries()

        '     WriteToUserProduction(dt)

    End Sub

#Region " User table "
#Region " Write user table records "
    Private Sub WriteToUserProduction(ByVal dt As DataTable)
        Dim i As Integer
        Dim Sql As String
        Dim Results As Results

        For i = 0 To dt.Rows.Count - 1
            InsertUsersRow(dt, i)
        Next

    End Sub

    Private Sub InsertUsersRow(ByVal dt As DataTable, ByVal i As Integer)
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results

        Try

            If cCommon.IsBlank(dt.Rows(i)("UserID")) Then
                Exit Sub
            End If

            Sql.Append("INSERT INTO [UserManagement].[dbo].[Users]([UserID], [FirstName], [LastName], [MI], [Role], [CompanyID], [LocationID], [PrimaryContactNumber], [AltContactNumber], [Email], [LastStatusChangeDate], [LastSessionID], [StatusCode], [LastLoginDate], [LastLoginIP], [WorkState], [ResidentState], [AddDate], [ChangeDate])")
            Sql.Append(" VALUES (")

            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("UserID"), False, StringTreatEnum.SideQts) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("FirstName"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("LastName"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(String.Empty, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("RoleID"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("ClientID"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("SiteID"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("Phone"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append("'', ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("Email"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append("'1/1/1950', ")
            Sql.Append("'', ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("StatusCode"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("LastLoginDate"), False, True) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("LastLoginIP"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append("'**', ")
            Sql.Append("'**', ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("AddDate"), False, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("ChangeDate"), False, True) & ")")

            'ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-SQL", "UserManagement", Sql.ToString)

            'Dim SqlConnection1 As New SqlClient.SqlConnection(cEnviro.ConnectionString)
            'Dim SqlCmd As New System.Data.SqlClient.SqlCommand(Sql.ToString, SqlConnection1)
            'SqlCmd.CommandType = System.Data.CommandType.Text
            'SqlConnection1.Open()
            'SqlCmd.ExecuteNonQuery()
            'SqlCmd.Dispose()
            'SqlConnection1.Close()

        Catch ex As Exception
            Stop
        End Try
    End Sub
#End Region

#Region " Get the user table data and condition it "
    Private Function GetUserTable() As DataTable
        Dim i As Integer
        Dim dt As DataTable

        Try

            ' ___ Get a copy of the production source table
            'dt = GetDTAllDB("SELECT * FROM BVIUser", "Prod", "BVI")
            dt = cCommon.GetDT("SELECT * FROM BVIUser", "HBG-SQL", "BVI")
            dt.Columns.Add("RecNum", GetType(System.Int64))

            ' Step #1: Number the rows
            For i = 0 To dt.Rows.Count - 1
                dt.Rows(i)("RecNum") = i + 1
            Next

            ' ___ Step #2: RoleID
            For i = 0 To dt.Rows.Count - 1
                If cCommon.IsBlank(dt.Rows(i)("RoleID")) Then
                    dt.Rows(i)("RoleID") = "*none*"
                Else
                    Select Case dt.Rows(i)("RoleID")
                        Case "A", "a"
                            dt.Rows(i)("RoleID") = "ADMIN"
                        Case "C", "c"
                            dt.Rows(i)("RoleID") = "CLIENT"
                        Case Else
                            dt.Rows(i)("RoleID") = "*invalid*"
                    End Select
                End If
            Next

            ' ___ Step #3: StatusCode
            For i = 0 To dt.Rows.Count - 1
                If cCommon.IsBlank(dt.Rows(i)("StatusCode")) Then
                    dt.Rows(i)("StatusCode") = "*none*"
                Else
                    Select Case dt.Rows(i)("StatusCode")
                        Case "A", "a"
                            dt.Rows(i)("StatusCode") = "ACTIVE"
                        Case "I", "i"
                            dt.Rows(i)("StatusCode") = "INACTIVE"
                        Case Else
                            dt.Rows(i)("StatusCode") = "*invalid*"
                    End Select
                End If
            Next

            ' ___ Step #4: ClientID
            For i = 0 To dt.Rows.Count - 1
                If cCommon.IsBlank(dt.Rows(i)("ClientID")) Then
                    dt.Rows(i)("ClientID") = "*none*"
                End If
            Next

            ' ___ Step #5: SiteID
            For i = 0 To dt.Rows.Count - 1
                If cCommon.IsBlank(dt.Rows(i)("SiteID")) Then
                    dt.Rows(i)("SiteID") = "*none*"
                End If
            Next

            ' ___ Step #6: Enter * where value is blank  and nulls not permitted
            For i = 0 To dt.Rows.Count - 1
                If cCommon.IsBlank(dt.Rows(i)("FirstName")) Then
                    dt.Rows(i)("FirstName") = "*"
                End If

                If cCommon.IsBlank(dt.Rows(i)("LastName")) Then
                    dt.Rows(i)("FirstName") = "*"
                End If

                If cCommon.IsValidPhoneNumber(dt.Rows(i)("Phone")) Then
                    dt.Rows(i)("Phone") = cCommon.PhoneOutHandler(dt.Rows(i)("Phone"), False, False)
                End If

                If Not cCommon.IsValidEmailAddress(dt.Rows(i)("Email")) Then
                    dt.Rows(i)("Email") = "notprovided@notprovided.com"
                End If
            Next

            OutputToExcel(dt)

            Return dt
        Catch ex As Exception
            Stop
        End Try

    End Function
#End Region

#End Region

#Region " Code Tables "

    ' ____ Codes_ClientID
    Private Sub Codes_ClientID()
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim i As Integer
        'dt = GetDTAllDB("SELECT * FROM Codes_ClientID", "Test", "UserManagement")
        dt = cCommon.GetDT("SELECT * FROM Codes_ClientID", "HBG-TST", "UserManagement")
        For i = 0 To dt.Rows.Count - 1
            Sql.Length = 0
            Sql.Append("INSERT INTO Codes_ClientID (ClientID, Name, Project, LogicalDelete) VALUES (")
            Sql.Append("'" & dt.Rows(i)("ClientID") & "', ")
            Sql.Append("'" & dt.Rows(i)("Name") & "', ")
            Sql.Append("'" & dt.Rows(i)("Project") & "', ")
            Sql.Append(cCommon.BitToString(dt.Rows(i)("LogicalDelete"), "1", "0", False) & ")")
            'ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-SQL", "UserManagement", Sql.ToString)
        Next
    End Sub


    ' ____ Codes_Role
    Private Sub Codes_Role()
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim i As Integer
        'dt = GetDTAllDB("SELECT * FROM Codes_Role", "Test", "UserManagement")
        dt = cCommon.GetDT("SELECT * FROM Codes_Role", "HBG-TST", "UserManagement")
        For i = 0 To dt.Rows.Count - 1
            Sql.Length = 0
            Sql.Append("INSERT INTO Codes_Role (Role, Description) VALUES (")
            Sql.Append("'" & dt.Rows(i)("Role") & "', ")
            Sql.Append("'" & dt.Rows(i)("Description") & "')")
            'ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-SQL", "UserManagement", Sql.ToString)
        Next
    End Sub

    Private Sub Codes_CarrierID()
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim i As Integer
        'dt = GetDTAllDB("SELECT * FROM Codes_CarrierID", "Test", "UserManagement")
        dt = cCommon.GetDT("SELECT * FROM Codes_CarrierID", "HBG-TST", "UserManagement")
        For i = 0 To dt.Rows.Count - 1
            Sql.Length = 0
            Sql.Append("INSERT INTO Codes_CarrierID (CarrierID, Description, AddDate, ChangeDate) VALUES (")
            Sql.Append("'" & dt.Rows(i)("CarrierID") & "', ")
            Sql.Append("'" & dt.Rows(i)("Description") & "', ")
            Sql.Append("'" & dt.Rows(i)("AddDate") & "', ")
            Sql.Append("'" & dt.Rows(i)("ChangeDate") & "') ")
            'ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-SQL", "UserManagement", Sql.ToString)
        Next
    End Sub

    Private Sub Codes_LocationID()
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim i As Integer
        'dt = GetDTAllDB("SELECT * FROM Codes_LocationID", "Test", "UserManagement")
        dt = cCommon.GetDT("SELECT * FROM Codes_LocationID", "HBG-TST", "UserManagement")
        For i = 0 To dt.Rows.Count - 1
            Sql.Length = 0
            Sql.Append("INSERT INTO Codes_LocationID (LocationID, Description) VALUES (")
            Sql.Append("'" & dt.Rows(i)("LocationID") & "', ")
            Sql.Append("'" & dt.Rows(i)("Description") & "')")
            'ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-SQL", "UserManagement", Sql.ToString)
        Next
    End Sub

    Private Sub Codes_State()
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim i As Integer
        'dt = GetDTAllDB("SELECT * FROM Codes_State", "Test", "UserManagement")
        dt = cCommon.GetDT("SELECT * FROM Codes_State", "HBG-TST", "UserManagement")
        For i = 0 To dt.Rows.Count - 1
            Sql.Length = 0
            Sql.Append("INSERT INTO Codes_State (StateCode, StateDesc, Country, OrderBy, LogicalDelete) VALUES (")
            Sql.Append("'" & dt.Rows(i)("StateCode") & "', ")
            Sql.Append("'" & dt.Rows(i)("StateDesc") & "', ")
            Sql.Append("'" & dt.Rows(i)("Country") & "', ")
            Sql.Append(dt.Rows(i)("OrderBy") & ", ")
            Sql.Append(cCommon.BitToString(dt.Rows(i)("LogicalDelete"), "1", "0", False) & ")")
            'ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-SQL", "UserManagement", Sql.ToString)
        Next
    End Sub

    Private Sub Codes_CompanyID()
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim i As Integer
        'dt = GetDTAllDB("SELECT * FROM Client", "Prod", "BVI")
        dt = cCommon.GetDT("SELECT * FROM Client", "HBG-SQL", "BVI")
        For i = 0 To dt.Rows.Count - 1
            Sql.Length = 0
            Sql.Append("INSERT INTO Codes_CompanyID (CompanyID, Description, PrimaryContactName, PrimaryContactPhone, ContactNotes, AddDate, ChangeDate) VALUES (")
            Sql.Append("'" & dt.Rows(i)("ClientID") & "', ")
            Sql.Append("'" & dt.Rows(i)("ClientName") & "', ")
            Sql.Append("'**', ")
            Sql.Append("'**', ")
            Sql.Append("'**', ")
            Sql.Append("'" & dt.Rows(i)("AddDate") & "', ")
            Sql.Append("'" & dt.Rows(i)("ChangeDate") & "')")
            'ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-SQL", "UserManagement", Sql.ToString)
        Next
    End Sub

    Private Sub Codes_Status()
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim i As Integer
        'dt = GetDTAllDB("SELECT * FROM Codes_Status", "Test", "UserManagement")
        dt = cCommon.GetDT("SELECT * FROM Codes_Status", "HBG-TST", "UserManagement")
        For i = 0 To dt.Rows.Count - 1
            Sql.Length = 0
            Sql.Append("INSERT INTO Codes_Status (Status, Description) VALUES (")
            Sql.Append("'" & dt.Rows(i)("Status") & "', ")
            Sql.Append("'" & dt.Rows(i)("Description") & "')")
            'ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-SQL", "UserManagement", Sql.ToString)
        Next
    End Sub


    Private Sub OldCodeTableEntries()

        Try
            Dim dtRoleID As DataTable
            'dtRoleID = GetDTAllDB("SELECT DISTINCT RoleID FROM BVIUser Order By RoleID", "Prod", "BVI")
            dtRoleID = cCommon.GetDT("SELECT DISTINCT RoleID FROM BVIUser Order By RoleID", "HBG-SQL", "BVI")
            dgRoleID.DataSource = dtRoleID
            dgRoleID.DataBind()

            Dim dtStatusCode As DataTable
            'dtStatusCode = GetDTAllDB("SELECT DISTINCT StatusCode FROM BVIUser Order By StatusCode", "Prod", "BVI")
            dtStatusCode = cCommon.GetDT("SELECT DISTINCT StatusCode FROM BVIUser Order By StatusCode", "HBG-SQL", "BVI")
            dgStatusCode.DataSource = dtStatusCode
            dgStatusCode.DataBind()

            Dim dtClientID As DataTable
            'dtClientID = GetDTAllDB("SELECT DISTINCT ClientID FROM BVIUser Order By ClientID", "Prod", "BVI")
            dtClientID = cCommon.GetDT("SELECT DISTINCT ClientID FROM BVIUser Order By ClientID", "HBG-SQL", "BVI")
            dgClientID.DataSource = dtClientID
            dgClientID.DataBind()

            Dim dtSiteID As DataTable
            'dtSiteID = GetDTAllDB("SELECT DISTINCT SiteID FROM BVIUser Order By SiteID", "Prod", "BVI")
            dtSiteID = cCommon.GetDT("SELECT DISTINCT SiteID FROM BVIUser Order By SiteID", "HBG-SQL", "BVI")
            dgSiteID.DataSource = dtSiteID
            dgSiteID.DataBind()

        Catch ex As Exception
            Stop
        End Try

    End Sub

    Private Sub CodeTableEntries()
        Dim Sql As String

        Try

            Dim dtResidentState As DataTable
            'dtResidentState = GetDTAllDB("SELECT DISTINCT ResidentState FROM Users Order By ResidentState", "Prod", "UserManagement")
            dtResidentState = cCommon.GetDT("SELECT DISTINCT ResidentState FROM Users Order By ResidentState", "HBG-SQL", "UserManagement")
            dgResidentState.DataSource = dtResidentState
            dgResidentState.DataBind()

            Dim dtWorkState As DataTable
            'dtWorkState = GetDTAllDB("SELECT DISTINCT WorkState FROM Users Order By WorkState", "Prod", "UserManagement")
            dtWorkState = cCommon.GetDT("SELECT DISTINCT WorkState FROM Users Order By WorkState", "HBG-SQL", "UserManagement")
            dgWorkState.DataSource = dtWorkState
            dgWorkState.DataBind()

            Dim dtRoleID As DataTable
            'dtRoleID = GetDTAllDB("SELECT DISTINCT Role FROM Users Order By Role", "Prod", "UserManagement")
            dtRoleID = cCommon.GetDT("SELECT DISTINCT Role FROM Users Order By Role", "HBG-SQL", "UserManagement")
            dgRoleID.DataSource = dtRoleID
            dgRoleID.DataBind()

            Dim dtStatusCode As DataTable
            'dtStatusCode = GetDTAllDB("SELECT DISTINCT StatusCode FROM Users Order By StatusCode", "Prod", "UserManagement")
            dtStatusCode = cCommon.GetDT("SELECT DISTINCT StatusCode FROM Users Order By StatusCode", "HBG-SQL", "UserManagement")
            dgStatusCode.DataSource = dtStatusCode
            dgStatusCode.DataBind()

            Dim dtCompanyID As DataTable
            'dtCompanyID = GetDTAllDB("SELECT DISTINCT CompanyID FROM Users Order By CompanyID", "Prod", "UserManagement")
            dtCompanyID = cCommon.GetDT("SELECT DISTINCT CompanyID FROM Users Order By CompanyID", "HBG-SQL", "UserManagement")
            dgClientID.DataSource = dtCompanyID
            dgClientID.DataBind()

            Dim dtLocationID As DataTable
            'dtLocationID = GetDTAllDB("SELECT DISTINCT LocationID FROM Users Order By LocationID", "Prod", "UserManagement")
            dtLocationID = cCommon.GetDT("SELECT DISTINCT LocationID FROM Users Order By LocationID", "HBG-SQL", "UserManagement")
            dgSiteID.DataSource = dtLocationID
            dgSiteID.DataBind()

        Catch ex As Exception
            Stop
        End Try

    End Sub
#End Region

#Region " Excel "
    Private Sub OutputToExcel(ByVal dt As DataTable)
        Dim oExcel As New Microsoft.Office.Interop.Excel.Application
        Dim oBooks As Microsoft.Office.Interop.Excel.Workbooks
        Dim oBook As Microsoft.Office.Interop.Excel.Workbook
        Dim oSheets As Microsoft.Office.Interop.Excel.Sheets
        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim oCells As Microsoft.Office.Interop.Excel.Range
        Dim sFile As String
        Dim sTemplate As String

        sFile = Server.MapPath(Request.ApplicationPath) & "\MigrationOutput.xls"
        sTemplate = Server.MapPath(Request.ApplicationPath) & "\MyTemplate.xls"
        oExcel.Visible = False
        oExcel.DisplayAlerts = False

        ' ___ Start a new workbook
        oBooks = oExcel.Workbooks
        oBooks.Open(Server.MapPath(Request.ApplicationPath) & "\MyTemplate.xls")     'Load colorful template with chart
        oBook = oBooks.Item(1)

        oSheets = oBook.Worksheets
        oSheet = CType(oSheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)
        oSheet.Name = "First Sheet"

        oCells = oSheet.Cells
        DumpData(dt, oCells) 'Fill in the data
        oSheet.SaveAs(sFile) 'Save in a temporary file
        oBook.Close()

        '___ Quit Excel and thoroughly deallocate everything
        oExcel.Quit()
        ReleaseComObject(oCells) : ReleaseComObject(oSheet)
        ReleaseComObject(oSheets) : ReleaseComObject(oBook)
        ReleaseComObject(oBooks) : ReleaseComObject(oExcel)
        oExcel = Nothing : oBooks = Nothing : oBook = Nothing
        oSheets = Nothing : oSheet = Nothing : oCells = Nothing
        System.GC.Collect()
        ' Response.Redirect(sFile) 'Send the user to the file
    End Sub

    'Private Function GetProductionRecords() As DataTable
    '    Dim Sql As String
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable

    '    Dim SqlCmd As New SqlCommand("Select * From BVIUser")
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection("user id=BVI_SQL_SERVER;password=noisivtifeneb;database=BVI;server=192.168.1.10")
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    DataAdapter.Fill(dt)
    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return dt
    'End Function

    'Outputs a DataTable to an Excel Worksheet
    Private Function DumpData(ByVal dt As DataTable, ByVal oCells As Microsoft.Office.Interop.Excel.Range) As String
        Dim dr As DataRow
        Dim ary() As Object
        Dim iRow As Integer, iCol As Integer

        'Output Column Headers
        For iCol = 0 To dt.Columns.Count - 1
            oCells(1, iCol + 1) = dt.Columns(iCol).ToString
        Next

        'Output Data
        For iRow = 0 To dt.Rows.Count - 1
            dr = dt.Rows.Item(iRow)
            ary = dr.ItemArray
            For iCol = 0 To UBound(ary)
                oCells(iRow + 2, iCol + 1) = ary(iCol).ToString
                ' Response.Write(ary(iCol).ToString & vbTab)
            Next
        Next
    End Function

    Private Sub ImportTableToExcel()
        Dim oExcel As New Microsoft.Office.Interop.Excel.Application
        Dim oBooks As Microsoft.Office.Interop.Excel.Workbooks
        Dim oBook As Microsoft.Office.Interop.Excel.Workbook
        Dim oSheets As Microsoft.Office.Interop.Excel.Sheets
        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim oCells As Microsoft.Office.Interop.Excel.Range
        Dim sFile As String
        Dim sTemplate As String
        Dim dt As DataTable

        'dt = GetDTAllDB("SELECT * FROM BVIUser", "Prod", "BVI")
        dt = cCommon.GetDT("SELECT * FROM BVIUser", "HBG-SQL", "BVI")

        sFile = Server.MapPath(Request.ApplicationPath) & "\MyExcel.xls"
        sTemplate = Server.MapPath(Request.ApplicationPath) & "\MyTemplate.xls"
        oExcel.Visible = False
        oExcel.DisplayAlerts = False

        ' ___ Start a new workbook
        oBooks = oExcel.Workbooks
        oBooks.Open(Server.MapPath(Request.ApplicationPath) & "\MyTemplate.xls")     'Load colorful template with chart
        oBook = oBooks.Item(1)

        oSheets = oBook.Worksheets
        oSheet = CType(oSheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)
        oSheet.Name = "First Sheet"

        oCells = oSheet.Cells
        DumpData(dt, oCells) 'Fill in the data
        oSheet.SaveAs(sFile) 'Save in a temporary file
        oBook.Close()

        '___ Quit Excel and thoroughly deallocate everything
        oExcel.Quit()
        ReleaseComObject(oCells) : ReleaseComObject(oSheet)
        ReleaseComObject(oSheets) : ReleaseComObject(oBook)
        ReleaseComObject(oBooks) : ReleaseComObject(oExcel)
        oExcel = Nothing : oBooks = Nothing : oBook = Nothing
        oSheets = Nothing : oSheet = Nothing : oCells = Nothing
        System.GC.Collect()
        ' Response.Redirect(sFile) 'Send the user to the file
    End Sub
#End Region

#Region " Database "
    'Private Function GetDTAllDB(ByVal Sql As String, ByVal SqlServer As String, ByVal DB As String) As DataTable
    '    Try
    '        Dim DataAdapter As Data.SqlClient.SqlDataAdapter
    '        Dim dt As New DataTable

    '        Dim SqlCmd As New Data.SqlClient.SqlCommand(Sql)
    '        SqlCmd.CommandType = CommandType.Text
    '        SqlCmd.Connection = New Data.SqlClient.SqlConnection(cEnviro.GetConnectionString(SqlServer, DB))
    '        DataAdapter = New Data.SqlClient.SqlDataAdapter(SqlCmd)
    '        DataAdapter.Fill(dt)
    '        DataAdapter.Dispose()
    '        SqlCmd.Dispose()
    '        Return dt
    '    Catch ex As Exception
    '        Stop
    '    End Try
    'End Function

    'Private Sub ExecuteNonQueryAllDB(ByVal Sql As String, ByVal SqlServer As String, ByVal DB As String)
    '    Try
    '        Dim SqlConnection1 As New Data.SqlClient.SqlConnection(cEnviro.SqlServerConnectionString(SqlServer, DB))
    '        SqlConnection1.Open()
    '        Dim SqlCmd As New System.Data.SqlClient.SqlCommand(Sql, SqlConnection1)
    '        SqlCmd.CommandType = System.Data.CommandType.Text
    '        SqlCmd.ExecuteNonQuery()
    '        SqlCmd.Dispose()
    '        SqlConnection1.Close()
    '    Catch ex As Exception
    '        Stop
    '    End Try
    'End Sub
#End Region

End Class



