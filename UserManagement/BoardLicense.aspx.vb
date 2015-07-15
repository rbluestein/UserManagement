Imports System.Data.SqlClient
Imports System.Runtime.InteropServices.Marshal
Imports Microsoft.Office.Interop
'http://www.aspnetpro.com/NewsletterArticle/2003/09/asp200309so_l/asp200309so_l.asp


Public Class BoardLicense
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
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private cCommon As New Common

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        '   ImportAppointments()


        '  UpdateRole()

        UpdateCarriers()


        '   ImportTableToExcel("Prod", "BVI", "License", "License")

        '   WriteToLicenceProduction()
        ' WriteToAppointments()

        '  Codes_Role()
        ' Codes_CarrierID()
        'Codes_LocationID()
        ' Codes_State()
        'Codes_CompanyID()
        'Codes_Status()
        'Codes_ClientID()

        ' CodeTableEntries()
        Exit Sub

        'Dim dt As DataTable
        'dt = GetUserTable()
        'CodeTableEntries()

        '     WriteToUserProduction(dt)

    End Sub



#Region " Appointments "
    Private Sub UpdateCarriers()
        Dim i As Integer
        Dim dtCarriers As DataTable
        Dim dt As DataTable
        Dim CarrierID As String
        Dim SqlStr As String
        Dim Sql As New System.Text.StringBuilder
        dtCarriers = ExcelToDT("SELECT DISTINCT CARRIERID FROM [ApptLicFieldsOnly$]", "License.xls", True)

        For i = 0 To dtCarriers.Rows.Count - 1

            CarrierID = dtCarriers.Rows(i)("CarrierID")
            SqlStr = "Select count (*) from Codes_CarrierID where lower(carrierid) = '" & dtCarriers.Rows(i)("CarrierID").ToString.ToLower & "'"
            dt = cCommon.GetDT(SqlStr, "HBG-TST", "UserManagement")


            If dt.Rows(0)(0) = 0 Then

                Sql.Length = 0
                Sql.Append("INSERT INTO Codes_CarrierID (CarrierID, Description, AddDate, ChangeDate)")
                Sql.Append(" Values ")
                Sql.Append("(" & cCommon.StrOutHandler(CarrierID, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(CarrierID, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("'" & cCommon.GetServerDateTime & "', ")
                Sql.Append("'" & cCommon.GetServerDateTime & "') ")
                'ExecuteNonQueryAllDB(Sql.ToString, "Test", "UserManagement")
                cCommon.ExecuteNonQuery("HBG-TST", "UserManagement", Sql.ToString)
            End If
        Next


        'DataGrid1.DataSource = dtCarriers
        'DataGrid1.DataBind()

    End Sub

    Private Sub UpdateRole()
        Dim dtLicensees As DataTable
        Dim i As Integer
        Dim Sql As String

        'dtLicensees = GetDTAllDB("select distinct userid  from userlicenses", "Test", "UserManagement")
        dtLicensees = cCommon.GetDT("select distinct userid  from userlicenses", "HBG-TST", "UserManagement")
        For i = 0 To dtLicensees.Rows.Count - 1
            Sql = "Update Users set Role='ENROLLER' WHERE UserID = '" & dtLicensees.Rows(i)("UserID") & "'"
            'ExecuteNonQueryAllDB(Sql, "Test", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-TST", "UserManagement", Sql.ToString)
        Next




    End Sub

    Private Function GetDTMapping() As DataTable
        Dim dtMapping As New DataTable
        Dim dtCarriers As DataTable
        Dim dtWorking As DataTable
        Dim i As Integer

        'Dim StatusCodeFldName As String
        'Dim AppointmentNumberFldName As String
        'Dim EffectiveDateFldName As String
        'Dim ExpirationDateFldName As String
        'Dim JITFldName As String


        dtMapping.Columns.Add("CarrierID", GetType(System.String))
        dtMapping.Columns.Add("StatusCodeFldName", GetType(System.String))
        dtMapping.Columns.Add("AppointmentNumberFldName", GetType(System.String))
        dtMapping.Columns.Add("EffectiveDateFldName", GetType(System.String))
        dtMapping.Columns.Add("ExpirationDateFldName", GetType(System.String))
        dtMapping.Columns.Add("JITFldName", GetType(System.String))

        ' ___ Add a row for each carrier
        dtCarriers = ExcelToDT("SELECT DISTINCT CARRIERID FROM [ApptLicFieldsOnly$]", "License.xls", True)
        For i = 0 To dtCarriers.Rows.Count - 1
            Dim dr As DataRow
            dr = dtMapping.NewRow
            dtMapping.Rows.Add(dr)
            dtMapping.Rows(i)("CarrierID") = dtCarriers.Rows(i)(0)
        Next

        ' ___ Map the field names between the production table and the new table
        dtCarriers = Nothing
        dtCarriers = ExcelToDT("SELECT * FROM [ApptLicFieldsOnly$]", "License.xls", True)

        For i = 0 To dtMapping.Rows.Count - 1
            dtWorking = Nothing
            dtWorking = ExcelToDT("SELECT Value FROM  [ApptLicFieldsOnly$] WHERE CarrierID = '" & dtMapping.Rows(i)(0) & "' AND FldName = 'plain'", "License.xls", True)
            If dtWorking.Rows.Count = 1 Then
                dtMapping.Rows(i)("StatusCodeFldName") = dtWorking.Rows(0)(0)
            Else
                dtMapping.Rows(i)("StatusCodeFldName") = "|none|"
            End If

            dtWorking = Nothing
            dtWorking = ExcelToDT("SELECT Value FROM  [ApptLicFieldsOnly$] WHERE CarrierID = '" & dtMapping.Rows(i)(0) & "' AND FldName = 'sharp'", "License.xls", True)
            If dtWorking.Rows.Count = 1 Then
                dtMapping.Rows(i)("AppointmentNumberFldName") = dtWorking.Rows(0)(0)
            Else
                dtMapping.Rows(i)("AppointmentNumberFldName") = "|none|"
            End If


            dtWorking = Nothing
            dtWorking = ExcelToDT("SELECT Value FROM  [ApptLicFieldsOnly$] WHERE CarrierID = '" & dtMapping.Rows(i)(0) & "' AND FldName = 'eff'", "License.xls", True)
            If dtWorking.Rows.Count = 1 Then
                dtMapping.Rows(i)("EffectiveDateFldName") = dtWorking.Rows(0)(0)
            Else
                dtMapping.Rows(i)("EffectiveDateFldName") = "|none|"
            End If


            dtWorking = Nothing
            dtWorking = ExcelToDT("SELECT Value FROM  [ApptLicFieldsOnly$] WHERE CarrierID = '" & dtMapping.Rows(i)(0) & "' AND FldName = 'exp'", "License.xls", True)
            If dtWorking.Rows.Count = 1 Then
                dtMapping.Rows(i)("ExpirationDateFldName") = dtWorking.Rows(0)(0)
            Else
                dtMapping.Rows(i)("ExpirationDateFldName") = "|none|"
            End If


            dtWorking = Nothing
            dtWorking = ExcelToDT("SELECT Value FROM  [ApptLicFieldsOnly$] WHERE CarrierID = '" & dtMapping.Rows(i)(0) & "' AND FldName = 'jit'", "License.xls", True)
            If dtWorking.Rows.Count = 1 Then
                dtMapping.Rows(i)("JITFldName") = dtWorking.Rows(0)(0)
            Else
                dtMapping.Rows(i)("JITFldName") = "|none|"
            End If
        Next

        Return dtMapping


    End Function

    Private Sub ImportAppointments()
        Dim dtProdLicenseTableRow As Integer
        Dim dtResults As New DataTable
        Dim dtProdLicenseTable As DataTable
        Dim dtCarriers As DataTable
        Dim dtMapping As DataTable

        ' ___ Get a list of the carriers in the ApptLicFieldsOnly worksheet
        '  This is a usable listing of carrier columns from the Prod.BVI..License table
        dtCarriers = ExcelToDT("SELECT DISTINCT CARRIERID FROM [ApptLicFieldsOnly$]", "License.xls", True)

        dtMapping = GetDTMapping()
        'DataGrid1.DataSource = dtMapping
        'DataGrid1.DataBind()
        'Exit Sub

        dtResults.Columns.Add("UserID", GetType(System.String))
        dtResults.Columns.Add("State", GetType(System.String))
        dtResults.Columns.Add("CarrierID", GetType(System.String))
        dtResults.Columns.Add("AppointmentNumber", GetType(System.String))
        dtResults.Columns.Add("StatusCode", GetType(System.String))
        dtResults.Columns.Add("EffectiveDate", GetType(System.DateTime))
        dtResults.Columns.Add("ExpirationDate", GetType(System.DateTime))
        dtResults.Columns.Add("JIT", GetType(System.String))
        dtResults.Columns.Add("AddDate", GetType(System.DateTime))
        dtResults.Columns.Add("ChangeDate", GetType(System.DateTime))

        ' dtProdLicenseTable = GetDTAllDB("SELECT * FROM License", "Prod", "BVI")

        '  dtProdLicenseTable = GetDTAllDB("SELECT * FROM License WHERE lower(left(userid, 1)) >= 'a' and lower(left(userid, 1)) <= 'c'", "Prod", "BVI")
        'dtProdLicenseTable = GetDTAllDB("SELECT * FROM License WHERE lower(left(userid, 1)) = 'z'", "Prod", "BVI")
        dtProdLicenseTable = cCommon.GetDT("SELECT * FROM License WHERE lower(left(userid, 1)) = 'z'", "HBG-SQL", "BVI")

        For dtProdLicenseTableRow = 0 To dtProdLicenseTable.Rows.Count - 1

            ' Each row is unique to user and state
            ProcessUserStateRow(dtProdLicenseTableRow, dtProdLicenseTable, dtResults, dtCarriers, dtMapping)

        Next

        WriteToUserAppointments(dtResults)

        'DTToExcel(dtResults, "DTResults.xls")

        DataGrid1.DataSource = dtResults
        DataGrid1.DataBind()



        'DataGrid1.DataSource = dt
        'DataGrid1.DataBind()
    End Sub

    Private Sub WriteToUserAppointments(ByRef dtResults As DataTable)
        Dim i As Integer
        Dim Sql As New System.Text.StringBuilder


        For i = 0 To dtResults.Rows.Count - 1
            Sql.Length = 0
            Sql.Append("INSERT INTO UserAppointments (UserID, State, CarrierID, AppointmentNumber, StatusCode, ")
            Sql.Append("StatusCodeLastChangeDate, ApplicationDate, EffectiveDate, ExpirationDate, JIT, AddDate, ChangeDate) ")
            Sql.Append(" VALUES (")
            Sql.Append(cCommon.StrOutHandler(dtResults.Rows(i)("UserID"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dtResults.Rows(i)("State"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dtResults.Rows(i)("CarrierID"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dtResults.Rows(i)("AppointmentNumber"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dtResults.Rows(i)("StatusCode"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append("'1/1/1950', ")   'StatusCodeLastChangeDate
            Sql.Append("null, ")   'ApplicationDate
            Sql.Append(cCommon.DateOutHandler(dtResults.Rows(i)("EffectiveDate"), True, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(dtResults.Rows(i)("ExpirationDate"), True, True) & ", ")
            Sql.Append("0, ")
            Sql.Append(cCommon.DateOutHandler(dtResults.Rows(i)("AddDate"), False, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(dtResults.Rows(i)("ChangeDate"), False, True) & ") ")
            'ExecuteNonQueryAllDB(Sql.ToString, "Test", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-TST", "UserManagement", Sql.ToString)
        Next


    End Sub

    Private Sub ProcessUserStateRow(ByVal dtProdLicenseTableRow As Integer, ByRef dtProdLicenseTable As DataTable, ByRef dtResults As DataTable, ByRef dtCarriers As DataTable, ByRef dtMapping As DataTable)
        Dim i As Integer

        'For dtCarriersRow = 0 To dtCarriers.Rows.Count - 1
        '    ProcessUserStateCarrier(dtProdLicenseTableRow, dtProdLicenseTable, dtResults, dtCarriers, dtCarriersRow, dtMapping)
        'Next

        For i = 0 To dtCarriers.Rows.Count - 1
            ProcessUserStateCarrier(dtProdLicenseTableRow, dtProdLicenseTable, dtResults, dtCarriers.Rows(i)(0), dtMapping)
        Next


    End Sub

    Private Sub ProcessUserStateCarrier(ByVal dtProdLicenseTableRow As Integer, ByRef dtProdLicenseTable As DataTable, ByRef dtResults As DataTable, ByVal CarrierID As String, ByRef dtMapping As DataTable)
        Dim StatusCodeFldName As String
        Dim AppointmentNumberFldName As String
        Dim EffectiveDateFldName As String
        Dim ExpirationDateFldName As String
        Dim JITFldName As String

        Dim i As Integer
        Dim UserID As String
        Dim State As String
        Dim StatusCode As String
        Dim AppointmentNumber As String
        Dim EffectiveDate As DateTime
        Dim ExpirationDate As DateTime
        Dim JIT As String
        Dim AddDate As DateTime
        Dim ChangeDate As DateTime

        UserID = dtProdLicenseTable.Rows(dtProdLicenseTableRow)("UserID")
        State = dtProdLicenseTable.Rows(dtProdLicenseTableRow)("State")
        AddDate = dtProdLicenseTable.Rows(dtProdLicenseTableRow)("AddDate")
        ChangeDate = dtProdLicenseTable.Rows(dtProdLicenseTableRow)("ChangeDate")

        StatusCodeFldName = GetFldName(dtMapping, CarrierID, "StatusCodeFldName")
        AppointmentNumberFldName = GetFldName(dtMapping, CarrierID, "AppointmentNumberFldName")
        EffectiveDateFldName = GetFldName(dtMapping, CarrierID, "EffectiveDateFldName")
        ExpirationDateFldName = GetFldName(dtMapping, CarrierID, "ExpirationDateFldName")
        JITFldName = GetFldName(dtMapping, CarrierID, "JITFldName")

        If StatusCodeFldName <> "|none|" Then
            Try
                StatusCode = dtProdLicenseTable.Rows(dtProdLicenseTableRow)(StatusCodeFldName)
            Catch
            End Try
        End If

        If AppointmentNumberFldName <> "|none|" Then
            Try
                AppointmentNumber = dtProdLicenseTable.Rows(dtProdLicenseTableRow)(AppointmentNumberFldName)
            Catch
            End Try
        End If

        If EffectiveDateFldName <> "|none|" Then
            Try
                EffectiveDate = dtProdLicenseTable.Rows(dtProdLicenseTableRow)(EffectiveDateFldName)
            Catch
            End Try
        End If

        If ExpirationDateFldName <> "|none|" Then
            Try
                ExpirationDate = dtProdLicenseTable.Rows(dtProdLicenseTableRow)(ExpirationDateFldName)
            Catch
            End Try
        End If

        If JITFldName <> "|none|" Then
            Try
                JIT = dtProdLicenseTable.Rows(dtProdLicenseTableRow)(JITFldName)
            Catch
            End Try
        End If

        AddResultsRow(dtResults, UserID, State, CarrierID, StatusCode, AppointmentNumber, EffectiveDate, ExpirationDate, JIT, AddDate, ChangeDate)
    End Sub

    Private Sub AddResultsRow(ByRef dtResults As DataTable, ByVal UserID As String, _
        ByVal State As String, ByVal CarrierID As String, ByVal StatusCode As String, _
        ByVal AppointmentNumber As String, ByVal EffectiveDate As DateTime, _
        ByVal ExpirationDate As DateTime, ByVal JIT As String, ByVal AddDate As DateTime, ByVal ChangeDate As DateTime)
        Dim dr As DataRow
        Dim RowNum As Integer

        If StatusCode = Nothing Then
            Exit Sub
        End If

        If StatusCode.ToLower = "x" Or StatusCode.ToLower = "p" Then
            dr = dtResults.NewRow
            dtResults.Rows.Add(dr)

            RowNum = dtResults.Rows.Count - 1

            dtResults.Rows(RowNum)("UserID") = cCommon.StrXferHandler(UserID, False)
            dtResults.Rows(RowNum)("State") = cCommon.StrXferHandler(State, False)
            dtResults.Rows(RowNum)("CarrierID") = cCommon.StrXferHandler(CarrierID, False)
            dtResults.Rows(RowNum)("AppointmentNumber") = cCommon.StrXferHandler(AppointmentNumber, False)
            dtResults.Rows(RowNum)("StatusCode") = cCommon.StrXferHandler(StatusCode.ToUpper, True)
            dtResults.Rows(RowNum)("AddDate") = cCommon.StrXferHandler(AddDate, False)
            dtResults.Rows(RowNum)("ChangeDate") = cCommon.StrXferHandler(ChangeDate, False)

            If StatusCode.ToLower = "x" Then
                dtResults.Rows(RowNum)("EffectiveDate") = cCommon.DateXferHandler(EffectiveDate, False)
                dtResults.Rows(RowNum)("ExpirationDate") = cCommon.DateXferHandler(ExpirationDate, True)
                dtResults.Rows(RowNum)("JIT") = cCommon.StrXferHandler(JIT, True)
            End If

        End If

    End Sub

    Private Function GetFldName(ByRef dtMapping As DataTable, ByVal CarrierID As String, ByVal ResultsFldName As String) As String
        Dim i As Integer
        For i = 0 To dtMapping.Rows.Count - 1
            If dtMapping.Rows(i)("CarrierID") = CarrierID Then
                Return dtMapping.Rows(i)(ResultsFldName)
            End If
        Next
    End Function

#End Region

#Region " License table "
    Private Function IsCurrentCarrier(ByVal CarrierID As String) As Boolean

    End Function


    Private Sub WriteToAppointments()
        Dim i As Integer
        Dim Sql As String
        Dim Results As Results
        Dim dt As DataTable

        ' ___ Create a table to hold the records from BVI..Licenses
        Dim dtAppt As New DataTable
        dtAppt.Columns.Add("UserID", GetType(System.String))
        dtAppt.Columns.Add("State", GetType(System.String))
        dtAppt.Columns.Add("CarrierID", GetType(System.String))
        dtAppt.Columns.Add("AppointmentNumber", GetType(System.String))

        ' ___ Get the source table from BVI..Licenses
        ' dt = GetDTAllDB("SELECT * FROM " & "License", "Prod", "BVI")
        dt = cCommon.GetDT("SELECT * FROM " & "License", "HBG-SQL", "BVI")

        ' ___ For each row in the table, determine which, if any of the licenses are (1) with current carriers and (2) are pending or effective
        For i = 0 To dt.Rows.Count - 1

            '     Dim Arr() as String()  {"Trustmark", "Combined", "NWB", "	ING", "CONSECO", "KANAWHA", "TransAmerica", "CNA", "SecurityMutua", "Monumental", "AFLAC", "AIG", 	"FloridaLife", "ProtLife", "MONY", "CA_LTC", "CO_LTC",  "NC_LTC", "IL_LTC", "WA_LTC",  "AmGen", "KANAWHA_ARAMARK", "TransAmOcc", "BosMut", "UnumProv", 


            If IsCurrentCarrier("Trustmark") Then

            End If

            '	Combined	Combined#	NWB#	ING	CONSECO	CONSECO#	KANAWHA	KANAWHA#	TransAmerica	TransAmerica#	FOUND1	FOUND1#	CNA	CNA#	SecurityMutual	Monumental	Monumental#	AFLAC	AFLAC#	AIG	AIG#	FloridaLife	ProtLife	ProtLife#	MONY	MONY#	CA_LTC	CO_LTC	NC_LTC	IL_LTC	WA_LTC	WA_LTC_ExpDate	RenewalSentDate	RenewalRecDate	Check#	Notes	ApplicationDate	License#	AmGen	KANAWHA_ARAMARK#	TransAmOcc	BosMut	UnumProv	UnumProv#	Trustmark_EffectiveDate	Combined_EffectiveDate	AmGen_EffectiveDate	Kanawha_EffectiveDate	TransAmerica_EffectiveDate	TransAmOcc_EffectiveDate	BosMut_EffectiveDate	UnumProv_EffectiveDate	Unum_JIT	CA_LTC_ExpDate	CO_LTC_ExpDate	NC_LTC_ExpDate	IL_LTC_ExpDate	AFLAC_EffectiveDate	JIT	SecurityMutual_EffectiveDate	HM_Life	HM_Life#	HM_Effective	Colonial	Colonial#	Colonial_Effective	AllState	AllState#	AllState_Effective	Active	AddDate	ChangeDate	rowguid


        Next


    End Sub



#Region " Write license table records "
    Private Sub WriteToLicenceProduction()
        Dim i As Integer
        Dim Sql As String
        Dim Results As Results
        Dim dt As DataTable

        'dt = GetDTAllDB("SELECT * FROM " & "License", "Prod", "BVI")
        dt = cCommon.GetDT("SELECT * FROM " & "License", "HBG-SQL", "BVI")

        For i = 0 To dt.Rows.Count - 1
            InsertLicenseRow(dt, i)
        Next

    End Sub

    Private Sub InsertLicenseRow(ByVal dt As DataTable, ByVal i As Integer)
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results

        Try

            If cCommon.IsBlank(dt.Rows(i)("UserID")) Then
                Exit Sub
            End If

            Sql.Append("INSERT INTO [UserManagement].[dbo].[UserLicenses]([UserID], [State], [LicenseNumber], [Status], [ApplicationDate], [EffectiveDate], [RenewalDateSent], [RenewalDateRecd], [ExpirationDate], [LongTermCareInd], [Notes], [AddDate], [ChangeDate])")
            Sql.Append(" VALUES (")

            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("UserID"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("State"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("License#"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.BitOutHandler(dt.Rows(i)("Active"), False) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("ApplicationDate"), True, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("EffectiveDate"), False, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("RenewalSentDate"), True, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("RenewalRecDate"), True, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("ExpirationDate"), True, True) & ", ")
            Sql.Append("0 , ")
            Sql.Append(cCommon.StrOutHandler(dt.Rows(i)("Notes"), False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("AddDate"), False, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(dt.Rows(i)("ChangeDate"), False, True) & ")")

            'ExecuteNonQueryAllDB(Sql.ToString, "Test", "UserManagement")
            cCommon.ExecuteNonQuery("HBG-TST", "UserManagement", Sql.ToString)

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
        ' dt = GetDTAllDB("SELECT * FROM Codes_ClientID", "Test", "UserManagement")
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
        ' dt = GetDTAllDB("SELECT * FROM Codes_Role", "Test", "UserManagement")
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
            ' ExecuteNonQueryAllDB(Sql.ToString, "Prod", "UserManagement")
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
        ' dt = GetDTAllDB("SELECT * FROM Codes_Status", "Test", "UserManagement")
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
            ' dtRoleID = GetDTAllDB("SELECT DISTINCT RoleID FROM BVIUser Order By RoleID", "Prod", "BVI")
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
    Private Function ExcelToDT(ByVal Sql As String, ByVal Filename As String, ByVal Header As Boolean) As DataTable
        'Dim ConnStr As String
        'ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\ImportExportManagement\MasterList.xls;Extended Properties=Excel 8.0;"


        Dim dt As New DataTable
        'Dim da As New OleDbDataAdapter("SELECT * FROM [Feed_DTS$]", ConnStr)
        '   Dim da As New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [ApptLicFieldsOnly$]", GetExcelConnectionString("License.xls", True))
        ' Dim da As New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & WorksheetName & "$]", GetExcelConnectionString(Filename, Header))
        Dim da As New System.Data.OleDb.OleDbDataAdapter(Sql, GetExcelConnectionString(Filename, Header))
        da.Fill(dt)
        Return dt
    End Function


    Private Function GetExcelConnectionString(ByVal FileName As String, ByVal Header As Boolean) As String
        ' http://www.connectionstrings.com/?carrier=excel

        Dim Results As String
        Results = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<fullpath>;Extended Properties=""Excel 8.0;HDR=<header>"";"

        If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
            '  Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\UserManagement\ProdBVIUSERSUSERScEnrollerStatusA.xls;Extended Properties=Excel 8.0;"
            ' Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\UserManagement\ProdBVIUSERSUSERScEnrollerStatusA.xls;Extended Properties=""Excel 8.0;HDR=Yes"";"
            'Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\UserManagement\ProdBVIUSERSUSERScEnrollerStatusA.xls;Extended Properties=""Excel 8.0;HDR=Yes"";"

            Results = Replace(Results, "<fullpath>", Server.MapPath(Request.ApplicationPath) & "\" & FileName)
            If Header Then
                Results = Replace(Results, "<header>", "Yes")
            Else
                Results = Replace(Results, "<header>", "No")
            End If
            Return Results

            '   Server.MapPath(Request.ApplicationPath) & "\MyTemplate.xls"

        Else
            ' Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\UserManagement\ProdBVIUSERSUSERScEnrollerStatusA.xls;Extended Properties=Excel 8.0;"
        End If
    End Function

    Public Sub DTToExcel(ByVal dt As DataTable, ByVal FileName As String)
        Dim oExcel As New Microsoft.Office.Interop.Excel.Application
        Dim oBooks As Microsoft.Office.Interop.Excel.Workbooks
        Dim oBook As Microsoft.Office.Interop.Excel.Workbook
        Dim oSheets As Microsoft.Office.Interop.Excel.Sheets
        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim oCells As Microsoft.Office.Interop.Excel.Range
        Dim sFile As String
        Dim sTemplate As String

        ' sFile = Server.MapPath(Request.ApplicationPath) & "\MigrationOutput.xls"
        sFile = Server.MapPath(Request.ApplicationPath) & "\" & FileName
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
        PopulateSpreadsheet(dt, oCells) 'Fill in the data
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

    Public Sub OutputToExcel(ByVal dt As DataTable)
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
        PopulateSpreadsheet(dt, oCells) 'Fill in the data
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
    Private Function PopulateSpreadsheet(ByVal dt As DataTable, ByVal oCells As Microsoft.Office.Interop.Excel.Range) As String
        Dim dr As DataRow
        Dim ary() As Object
        Dim iRow As Integer, iCol As Integer

        Try

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

        Catch ex As Exception
            Stop
        End Try
    End Function

    Private Sub ImportTableToExcel(ByVal SqlServer As String, ByVal DBName As String, ByVal Tablename As String, ByVal ExcelName As String)
        Dim oExcel As New Microsoft.Office.Interop.Excel.Application
        Dim oBooks As Microsoft.Office.Interop.Excel.Workbooks
        Dim oBook As Microsoft.Office.Interop.Excel.Workbook
        Dim oSheets As Microsoft.Office.Interop.Excel.Sheets
        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim oCells As Microsoft.Office.Interop.Excel.Range
        Dim sFile As String
        Dim sTemplate As String
        Dim dt As DataTable

        'dt = GetDTAllDB("SELECT top 100 * FROM " & Tablename, SqlServer, DBName)
        dt = cCommon.GetDT("SELECT top 100 * FROM " & Tablename, SqlServer, DBName)


        sFile = Server.MapPath(Request.ApplicationPath) & "\" & ExcelName & ".xls"
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
        PopulateSpreadsheet(dt, oCells) 'Fill in the data
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
    '        SqlCmd.Connection = New Data.SqlClient.SqlConnection(cEnviro.SqlServerConnectionString(SqlServer, DB))
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
