Imports System.Data
Imports System.Data.sqlclient
Imports System.Data.OleDb

Partial Class BoardRecords
    Inherits System.Web.UI.Page

    Dim cCommon As New Common
    Dim cEnviro As Enviro


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        cEnviro = Session("Enviro")
        TakeOne()
    End Sub

    Private Sub TakeOne()
        Dim i As Integer
        Dim dt As DataTable
        dt = GetExcelData()
        For i = 0 To dt.Rows.Count - 1
            PerformSave(dt, i)
        Next
    End Sub

    Private Function GetProductionRecords() As DataTable
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable

        Dim SqlCmd As New SqlCommand
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New SqlConnection("user id=BVI_SQL_SERVER;password=noisivtifeneb;database=UserManagement;server=192.168.1.15")
        DataAdapter = New SqlDataAdapter(SqlCmd)
        DataAdapter.Fill(dt)
        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return dt
    End Function


    Private Sub PerformSave(ByRef dt As DataTable, ByVal i As Integer)
        ' Dim RequestAction As RequestAction
        Dim Sql As New System.Text.StringBuilder
        ' RequestAction = RequestAction.SaveNew

        'If RequestAction = RequestAction.SaveNew Then
        'Sql.Append("INSERT INTO Codes_CarrierID (CarrierID, Description, ChangeDate)")

        Sql.Append("INSERT INTO Users (UserID, FirstName, LastName, MI, Role, CompanyID, LocationID, PrimaryContactNumber, AltContactNumber, Email, LastStatusChangeDate, Status, WorkState, ResidentState, Adddate, ChangeDate")
        Sql.Append(" Values ")

        'Sql.Append("(" & Common.StrOutHandler(dt.Rows(i)("UserName")) & ", ")
        'Sql.Append("(" & Common.StrOutHandler(dt.Rows(i)("FirstName")) & ", ")
        'Sql.Append("(" & Common.StrOutHandler(dt.Rows(i)("LastName")) & ", ")
        'Sql.Append("("", ")
        'Sql.Append("'ADMIN', ")
        'Sql.Append("'BVI', ")
        'Sql.Append("(" & Common.StrOutHandler(dt.Rows(i)("Location")) & ", ")
        'Sql.Append("'800-810-2200', ")
        'Sql.Append("'800-810-2200', ")
        'Sql.Append("(" & Common.StrOutHandler(dt.Rows(i)("UserName") & "@BenefitVision.com") & ", ")
        'Sql.Append("'" & Common.GetServerDateTime & "', ")
        'Sql.Append("'ACTIVE', ")
        'Sql.Append("'PA', ")
        'Sql.Append("'PA', ")
        'Sql.Append("'" & Common.GetServerDateTime & "', ")
        'Sql.Append("'" & Common.GetServerDateTime & "') ")

        'Sql.Append("(" & Common.StrOutHandler(txtCarrierID.Text) & ", ")
        'Sql.Append(Common.StrOutHandler(txtDescription.Text) & ", ")
        'Sql.Append("'" & Common.GetServerDateTime & "') ")

        'Else
        'Sql.Append("UPDATE Codes_CarrierID Set ")
        'Sql.Append("CarrierID = '" & txtCarrierID.Text & "', ")
        'Sql.Append("Description = '" & txtDescription.Text & "', ")
        'Sql.Append("ChangeDate = '" & Common.GetServerDateTime & "' ")
        'Sql.Append(" WHERE CarrierID = '" & cCarrierID & "'")
        'End If

        'Dim SqlConnection1 As New SqlClient.SqlConnection(cEnviro.ConnectionString)

        'NOT TESTED PROBABLY NEEDS WORK !!!!!!!!!!!!!!!!!!!!!!!!!!!!!1
        Dim SqlConnection1 As New SqlClient.SqlConnection(cEnviro.GetConnectionString)

        Dim SqlCmd As New System.Data.SqlClient.SqlCommand(Sql.ToString, SqlConnection1)
        SqlCmd.CommandType = System.Data.CommandType.Text
        SqlConnection1.Open()
        SqlCmd.ExecuteNonQuery()

        SqlCmd.Dispose()
        SqlConnection1.Close()
    End Sub

    Private Function GetExcelData() As DataTable
        'Dim ConnStr As String
        'ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\ImportExportManagement\MasterList.xls;Extended Properties=Excel 8.0;"


        Dim dt As New DataTable
        'Dim da As New OleDbDataAdapter("SELECT * FROM [Feed_DTS$]", ConnStr)
        Dim da As New OleDbDataAdapter("SELECT * FROM [Sheet1$]", ExcelConnectionString)
        da.Fill(dt)
        Return dt
    End Function

    Public Function ExcelConnectionString() As String
        ' http://www.connectionstrings.com/?carrier=excel
        'If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
        '  Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\UserManagement\ProdBVIUSERSUSERScEnrollerStatusA.xls;Extended Properties=Excel 8.0;"
        Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\UserManagement\ProdBVIUSERSUSERScEnrollerStatusA.xls;Extended Properties=""Excel 8.0;HDR=Yes"";"
        'Else
        ' Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\UserManagement\ProdBVIUSERSUSERScEnrollerStatusA.xls;Extended Properties=Excel 8.0;"
        'End If
    End Function

End Class

