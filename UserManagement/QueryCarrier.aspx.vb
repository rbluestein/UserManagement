Imports System.Data.SqlClient

Public Class QueryCarrier
        Inherits System.Web.UI.Page

#Region " Declarations "
    Protected PageCaption As HtmlGenericControl
    Protected WithEvents litDG As System.Web.UI.WebControls.Literal
    Protected WithEvents litFilterHiddens As System.Web.UI.WebControls.Literal
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Private cQueryCarrierSess As QueryCarrierSession
    Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
#End Region

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PageMode As PageMode
        Dim Results As Results
        Dim Action As String
        Dim DG As DG

        Try

            cEnviro = Session("Enviro")
            cCommon = New Common

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right check
            cRights = New RightsClass(Page)
            Dim RightsRqd(0) As String
            RightsRqd.SetValue(cRights.LicenseView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the querycarrier session object
            cQueryCarrierSess = Session("QueryCarrierSession")

            ' ___ Get the page mode 
            PageMode = cCommon.GetPageMode(Page, cQueryCarrierSess)

            ' ___ Load the page session variables
            LoadVariables(PageMode)

            ' ___ Initialize the datagrid
            DG = DefineDataGrid()

            ' ___ Execute action
            Select Case PageMode
                Case PageMode.Initial
                    DisplayPage(PageMode, "View", DG, DG.OrderByType.Initial)

                Case PageMode.Postback
                    Action = Request.Form("hdAction")
                    Select Case Action
                        Case "Sort"
                            DisplayPage(PageMode, "View", DG, DG.OrderByType.Field, Request.Form("hdSortField"))

                        Case "ApplyFilter"
                            DisplayPage(PageMode, "View", DG, DG.OrderByType.Recurring)

                        Case "Download"
                            DisplayPage(PageMode, "Download", DG, DG.OrderByType.Recurring)

                        Case "ExistingRecord"
                            Response.Redirect("UserMaintain.aspx?CallType=Existing")

                        Case "Permissions"
                            Response.Redirect("Permissions.aspx?CalledByOther&CallType=Existing")

                        Case "License"
                            Response.Redirect("LicenseWorklist.aspx?CalledBy=Other&CallType=Existing")

                        Case "NewRecord"
                            Response.Redirect("UserMaintain.aspx?CallType=New")
                    End Select

                Case PageMode.ReturnFromChild, PageMode.CalledByOther
                    DisplayPage(PageMode, "View", DG, DG.OrderByType.ReturnToPage)
                    If cQueryCarrierSess.PageReturnOnLoadMessage <> Nothing Then
                        litMsg.Text = "<script language='javascript'>alert('" & cQueryCarrierSess.PageReturnOnLoadMessage & "')</script>"
                        cQueryCarrierSess.PageReturnOnLoadMessage = Nothing
                    End If
            End Select

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'>"

        Catch ex As Exception
            Throw New Exception("Error #1102: QueryCarrier Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Sub LoadVariables(ByVal PageMode As PageMode)
        Try

            Select Case PageMode
                Case PageMode.Initial
                    ' No action
                Case PageMode.CalledByOther, PageMode.ReturnFromChild
                    ' No action
                Case PageMode.Postback
                    ' ___ Update session variables with those that the user may have changed
                    cQueryCarrierSess.UserID = Replace(Request.Form("txtua.UserID"), "~", "'")
                    cQueryCarrierSess.UserIDFilter = Request.Form("txtua.UserID")   'txtua.UserID
                    cQueryCarrierSess.FullNameFilter = Request.Form("txtFullName")    ' txtFullName
                    cQueryCarrierSess.EmpStatusCodeFilter = Request.Form("ddu.StatusCode")   ' ddu.StatusCode
                    cQueryCarrierSess.LocationIDFilter = Request.Form("ddu.LocationID")
                    cQueryCarrierSess.LicStatusFilter = Request.Form("ddLicStatus")
                    cQueryCarrierSess.CarStatusCodeFilter = Request.Form("ddua.StatusCode")
                    cQueryCarrierSess.StateFilter = Request.Form("ddua.State")              ' ddua.State                 ddua.CarrierID
                    cQueryCarrierSess.CarrierIDFilter = Request.Form("ddua.CarrierID")
            End Select
        Catch ex As Exception
            Throw New Exception("Error #1104: QueryCarrier LoadVariables. " & ex.Message)
        End Try
    End Sub

    Private Function DefineDataGrid() As DG
        Try

            '  DG.AddDataBoundColumn("ItemName", "DataFldName", "HeaderText", "SortExpression", True

            Dim DG As New DG("UserID", cCommon, cRights, True, "EmbeddedTableDef", "ua.UserID")
            'DG.AddNewButton(cRights.UserEdit)
            DG.AddDataBoundColumn("UserID", "UserID", "User ID", "ua.UserID", True, Nothing, Nothing, "align='left'")
            ' DG.AddDataBoundColumn("FullName", "FullName", "Name", "u.LastName+u.FirstName+u.MI", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("FullName", "FullName", "Name", "FullName", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("EmpStatusCode", "StatusCode", "Emp Status", "u.StatusCode", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("LocationID", "LocationID", "Loc", "u.LocationID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("LicStatus", "LicStatus", "Lic Status", "LicStatus", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("CarStatusCode", "StatusCodeStr", "Car Status", "ua.StatusCodeStr", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("State", "State", "State", "ua.State", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("CarrierID", "CarrierID", "Carrier", "ua.CarrierID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("LicEffectiveDate", "LicEffectiveDate", "License<br>Eff Date", "ul.LicEffectiveDate", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("LicExpirationDate", "LicExpirationDate", "License<br>Exp Date", "ul.LicExpirationDate", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("CarEffectiveDate", "CarEffectiveDate", "Carrier<br>Eff Date", "ua.CarEffectiveDate", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("AppointmentNumber", "AppointmentNumber", "Appt<br>Num", "ua.AppointmentNumber", True, Nothing, Nothing, "align='left'")

            ' ___ Build the menu
            Dim Menu As DG.Menu
            Menu = DG.AttachMenu(7)
            Menu.AddItem(DG.Menu.ObjectTypeEnum.IsButton, "View", "View", cRights.LicenseView)
            Menu.AddItem(DG.Menu.ObjectTypeEnum.IsButton, "Download", "Download", cRights.LicenseView)


            ' ___ Build the filter
            Dim Filter As DG.Filter
            'Filter = DG.AttachFilter(DG.FilterOperationMode.FilterSwitchable, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialHide)
            Filter = DG.AttachFilter(DG.FilterOperationMode.FilterAlwaysOn, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialHide)
            Filter.AddTextbox("UserID", "ua.UserID", 50)
            Filter.AddTextbox("FullName", "FullName", 48, "LTrim(RTrim(u.LastName)) + ', ' + LTrim(RTrim(u.FirstName))", Nothing)
            Filter.AddDropdown("EmpStatusCode", "u.StatusCode")
            Filter.AddDropdown("LocationID", "u.LocationID")
            Filter.AddExtendedDropdown("LicStatus", "LicStatus")
            Filter.AddDropdown("CarStatusCode", "ua.StatusCode")
            Filter.AddDropdown("State", "ua.State")
            Filter.AddDropdown("CarrierID", "ua.CarrierID")

            Return DG

        Catch ex As Exception
            Throw New Exception("Error #1106: QueryCarrier DefineDataGrid. " & ex.Message)
        End Try
    End Function

    Private Sub DisplayPage(ByVal PageMode As PageMode, ByVal ViewOrDownload As String, ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Try

            ' ___ Handle the filter
            HandleFilter(DG, PageMode)

            ' ___ Handle the sort
            If cQueryCarrierSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cQueryCarrierSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            ' ___ Handle the data
            HandleData(DG, PageMode, OrderByType, ViewOrDownload)

            ' ___ Set the FilterOnOffState
            cQueryCarrierSess.FilterOnOffState = "on"

            ' ___ Set the last field sorted and sort direction in the sort reference
            cQueryCarrierSess.SortReference = DG.GetSortReference

        Catch ex As Exception
            Throw New Exception("Error #1107: QueryCarrier DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleData(ByRef DG As DG, ByVal PageMode As PageMode, ByVal OrderByType As DG.OrderByType, ByVal ViewOrDownload As String)
        Dim dt As DataTable
        Dim RecordCount As Integer
        Dim Coll As Collection
        Dim SuppressDisplayData As Boolean
        Dim EmbeddedMessage As String
        Dim SPBuildExcel As Boolean
        Dim DownloadPathColl As Collection
        Dim HTMLReport As Boolean
        Dim DownloadReport As Boolean
        Dim ExceedsRecordMaximum As Boolean
        Dim ExcessiveRecords As Boolean
        Dim IgnoreExcessiveRecords As Boolean
        Dim NewQuery As Boolean
        Dim PerformNextTest As Boolean

        Try

            ' ___ #1: FIGURE OUT WHAT'S GOING ON

            ' ___ Get the recordcount
            Coll = GetQueryInfo("RecordCount", DG, OrderByType)
            RecordCount = Coll("RecordCount")


            ' ___ Test #1: InitialReportDataSuppress
            ' User opens this page. The datagrid suppresses the initial display of the data. 
            ' If the user navigates to a different page and then returns to this page, PageMode is no longer set
            ' to initial and the initial data is permitted to display. The Sql attempts to return all of the records the sql statement
            ' with no restrictions. A postback is required to enable the display of the data.
            If PageMode = PageMode.Initial AndAlso DG.RecordsInitialShowHide = DG.RecordsInitialShowHideEnum.RecordsInitialHide Then
                cQueryCarrierSess.InitialReportDataSuppressInEffect = True
            ElseIf PageMode = PageMode.Postback AndAlso cQueryCarrierSess.InitialReportDataSuppressInEffect Then
                cQueryCarrierSess.InitialReportDataSuppressInEffect = False
            End If
            If Not cQueryCarrierSess.InitialReportDataSuppressInEffect Then
                PerformNextTest = True
            End If

            ' ___ Test #2: Exceeds record maximum test
            If PerformNextTest Then
                If RecordCount > cEnviro.RecordMaximum Then
                    ExceedsRecordMaximum = True
                    PerformNextTest = False
                End If
            End If

            ' ___ Test #3: Ignore excessive records warning
            ' If an excessive records warning was put into effect and the user has chosen to proceed with the same query, 
            ' ignore and reset the warning.
            If PerformNextTest Then
                If cQueryCarrierSess.ExcessiveRecordsWarningInEffect Then
                    Coll = GetQueryInfo("Sql", DG, OrderByType)
                    If StrComp(Coll("Sql"), cQueryCarrierSess.Sql, CompareMethod.Text) = 0 Then
                        IgnoreExcessiveRecords = True
                        PerformNextTest = False
                    Else
                        NewQuery = True
                    End If
                Else
                    NewQuery = True
                End If

                ' ___ Reset the excessive record warning properties
                If cQueryCarrierSess.ExcessiveRecordsWarningInEffect Then
                    cQueryCarrierSess.ExcessiveRecordsWarningInEffect = False
                    cQueryCarrierSess.Sql = String.Empty
                End If
            End If

            ' ___ Test #4: Excessive records test
            If PerformNextTest Then
                If RecordCount > cEnviro.ExcessiveRecordAmount Then
                    Coll = GetQueryInfo("Sql", DG, OrderByType)
                    cQueryCarrierSess.ExcessiveRecordsWarningInEffect = True
                    cQueryCarrierSess.Sql = Coll("Sql")
                End If
            End If


            ' ___ DIRECT THE REPORT
            If ViewOrDownload = "Download" Then
                DownloadReport = True
            Else
                HTMLReport = True
            End If
            If cQueryCarrierSess.InitialReportDataSuppressInEffect Then
                HTMLReport = True
                DownloadReport = False
                SuppressDisplayData = True
            ElseIf ExceedsRecordMaximum Then
                HTMLReport = True
                DownloadReport = False
                SuppressDisplayData = True
                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">&nbsp;&nbsp;Report contains " & RecordCount & " records. Respecify report.</td>"  '"<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">" & EmbeddedMessage & "</td>"
            ElseIf IgnoreExcessiveRecords Then
                HTMLReport = False
                DownloadReport = True
            ElseIf cQueryCarrierSess.ExcessiveRecordsWarningInEffect Then
                HTMLReport = True
                DownloadReport = False
                SuppressDisplayData = True
                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">&nbsp;&nbsp;Report contains " & RecordCount & " records. Proceed or respecify report.</td>"  '"<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">" & EmbeddedMessage & "</td>"
            End If
            If DownloadReport Then
                DownloadPathColl = cCommon.GetDownloadPath(Page)
                SuppressDisplayData = True
                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif"">&nbsp;&nbsp;Click here to <a href=""" & DownloadPathColl("RelPath") & """>download</a> your CSV file.</td>"
            End If


            ' ___ EXECUTE THE REPORT

            ' ___ Get the data
            If (HTMLReport And (Not SuppressDisplayData)) OrElse DownloadReport Then
                Coll = GetQueryInfo("Data", DG, OrderByType)
                dt = Coll("Data")
            End If

            ' ___ Process the download
            If DownloadReport Then
                cCommon.PrintCSVVersionLocal(dt, DownloadPathColl("AbsPath"))
            End If

            ' ___ Process the html
            If SuppressDisplayData Then
                dt = Nothing
            End If
            litDG.Text = DG.GetText(dt, EmbeddedMessage)

        Catch ex As Exception
            Throw New Exception("Error #1210: QueryLicense DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleFilter(ByRef DG As DG, ByVal PageMode As PageMode)
        Dim i As Integer
        Dim dt As DataTable

        Try

            ' ___ Get a filter reference
            Dim Filter As DG.Filter
            Filter = DG.GetFilter

            ' ___ UserID, FullName
            If PageMode <> PageMode.Initial Then
                Filter.Coll("UserID").SetFilterValue(cQueryCarrierSess.UserIDFilter)
                Filter.Coll("FullName").SetFilterValue(cQueryCarrierSess.FullNameFilter)
            End If

            ' ___ EmpStatusCode
            Filter("EmpStatusCode").AddDropdownItem("", "")
            If PageMode = PageMode.Initial Then
                Filter("EmpStatusCode").AddDropdownItem("Active", "Active", True)
            Else
                Filter("EmpStatusCode").AddDropdownItem("Active", "Active")
            End If
            Filter("EmpStatusCode").AddDropdownItem("Inactive", "Inactive")
            Filter("EmpStatusCode").AddDropdownItem("Terminated", "Terminated")
            If PageMode <> PageMode.Initial Then
                Filter.Coll("EmpStatusCode").SetFilterValue(cQueryCarrierSess.EmpStatusCodeFilter)
            End If

            ' ___ LocationID
            dt = cCommon.GetDT("Select LocationID from Codes_LocationID WHERE LocationID <> 'CLIENT' Order By LocationID")
            If PageMode = PageMode.Initial Then
                Filter("LocationID").AddDropdownItem("", "", True)
            Else
                Filter("LocationID").AddDropdownItem("", "")
            End If
            For i = 0 To dt.Rows.Count - 1
                Filter("LocationID").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0))
            Next
            If PageMode <> PageMode.Initial Then
                Filter.Coll("LocationID").SetFilterValue(cQueryCarrierSess.LocationIDFilter)
            End If

            ' ___ LicStatus
            Filter("LicStatus").AddExtendedDropdownItem("", "", "", True)
            Filter("LicStatus").AddExtendedDropdownItem("1", "Application", " ul.EffectiveDate = '1/1/1950' ", False)
            Filter("LicStatus").AddExtendedDropdownItem("2", "Effective", " ul.EffectiveDate <> '1/1/1950' AND (ul.ExpirationDate is null OR ul.ExpirationDate >= getDate()) ", False)
            Filter("LicStatus").AddExtendedDropdownItem("3", "Expired", " ul.ExpirationDate < getDate() ", False)
            If PageMode <> PageMode.Initial Then
                Filter.Coll("LicStatus").SetFilterValue(cQueryCarrierSess.LicStatusFilter)
            End If

            ' ___ CarStatusCode
            If PageMode = PageMode.Initial Then
                Filter("CarStatusCode").AddDropdownItem("", "", True)
            Else
                Filter("CarStatusCode").AddDropdownItem("", "")
            End If
            Filter("CarStatusCode").AddDropdownItem("P", "Pending")
            Filter("CarStatusCode").AddDropdownItem("X", "Effective")
            Filter("CarStatusCode").AddDropdownItem("T", "Terminated")
            If PageMode <> PageMode.Initial Then
                Filter.Coll("CarStatusCode").SetFilterValue(cQueryCarrierSess.CarStatusCodeFilter)
            End If

            ' ___ State
            dt = cCommon.GetDT("Select StateCode from Codes_State Order By StateCode")
            Filter("State").AddDropdownItem("", "", True)
            If PageMode = PageMode.Initial Then
                For i = 0 To dt.Rows.Count - 1
                    If i = 0 Then
                        Filter("State").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), True)
                    Else
                        Filter("State").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
                    End If
                Next
            Else
                For i = 0 To dt.Rows.Count - 1
                    Filter("State").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
                Next
            End If
            If PageMode <> PageMode.Initial Then
                Filter.Coll("State").SetFilterValue(cQueryCarrierSess.StateFilter)
            End If


            ' ___ Carrier
            dt = cCommon.GetDT("Select CarrierID, Description from Codes_CarrierID Order By Description")
            Filter("CarrierID").AddDropdownItem("", "", True)
            If PageMode = PageMode.Initial Then
                For i = 0 To dt.Rows.Count - 1
                    If i = 0 Then
                        Filter("CarrierID").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), True)
                    Else
                        Filter("CarrierID").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
                    End If
                Next
            Else
                For i = 0 To dt.Rows.Count - 1
                    Filter("CarrierID").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
                Next
            End If
            If PageMode <> PageMode.Initial Then
                Filter.Coll("CarrierID").SetFilterValue(cQueryCarrierSess.CarrierIDFilter)
            End If

        Catch ex As Exception
            Throw New Exception("Error #1109: QueryCarrier HandleFilter. " & ex.Message)
        End Try
    End Sub


    Private Function GetQueryInfo(ByVal InfoType As String, ByVal DG As DG, ByVal OrderByType As DG.OrderByType) As Collection
        Dim sb As New System.Text.StringBuilder
        Dim Sql As String
        Dim dt As DataTable
        Dim ShowFilter As Boolean
        Dim Coll As New Collection
        Dim Pos As Integer
        Dim SecurityWhereClause As String
        Dim AccessLevel As RightsClass.AccessLevelEnum
        Dim Role As String
        Dim LocationID As String

        Try

            ' ___ Get the security settings for the logged in user
            cRights.GetSecurityFlds(AccessLevel, Role, LocationID)

            If InfoType = "RecordCount" Then
                sb.Append("SELECT Count (*) ")

            ElseIf InfoType = "Data" Or InfoType = "Sql" Then
                sb.Append("SELECT ua.UserId, ")
                sb.Append("LTrim(RTrim(u.LastName)) + ', ' + LTrim(RTrim(u.FirstName))  FullName,  ")
                sb.Append("u.StatusCode, ")
                'sb.Append("EmpStatusStr = case when u.StatusCode='ACTIVE' then 'Active' else 'Inactive' end, ")
                sb.Append("EmpStatusStr = case when u.StatusCode='ACTIVE' then 'Active'  when u.StatusCode='TERMINATED' then 'Terminated' else 'Inactive' end, ")
                sb.Append("u.LocationID, ")


                sb.Append("LicStatus = case ")
                sb.Append(" when ul.ExpirationDate < getDate() then 'Expired' ")
                sb.Append(" when ul.EffectiveDate = '1/1/1950' then 'Application' ")
                sb.Append(" else 'Effective' ")
                sb.Append(" end, ")

                sb.Append("StatusCodeStr = case when ua.StatusCode='X' then 'Effective' when ua.StatusCode='P' then 'Pending' when ua.StatusCode = 'T' then 'Terminated' end, ")
                sb.Append("ua.State, ")
                sb.Append("ul.EffectiveDate as LicEffectiveDate, ")
                sb.Append("ul.ExpirationDate as LicExpirationDate, ")
                sb.Append("ua.EffectiveDate as CarEffectiveDate, ")
                sb.Append("ua.CarrierID, ")
                sb.Append("ua.AppointmentNumber ")
            End If

            sb.Append("FROM UserAppointments ua ")
            sb.Append("INNER JOIN UserLicenses ul on ua.UserID = ul.UserID and ua.State = ul.State ")
            sb.Append("INNER JOIN Users u on ua.UserID = u.UserID")
            Sql = sb.ToString

            ' ___ Where clause: Get the restrictions imposed by the security rules.
            SecurityWhereClause = cCommon.GetSecurityWhereClause("ul", AccessLevel, LocationID)

            'Select Case AccessLevel
            '    Case RightsClass.AccessLevelEnum.AllAccess
            '        SecurityWhereClause = String.Empty
            '    Case RightsClass.AccessLevelEnum.SupervisorAccess
            '        'SecurityWhereClause = " (Role = 'SUPERVISOR' or Role = 'ENROLLER') and LocationID = '" & LocationID & "'"
            '        SecurityWhereClause = " (ul.UserID = '" & cEnviro.LoggedInUserID & " ' or u.Role = 'ENROLLER') and u.LocationID = '" & LocationID & "'"
            '    Case RightsClass.AccessLevelEnum.EnrollerAccess
            '        SecurityWhereClause = "ul.UserID = '" & cEnviro.LoggedInUserID & "'"
            'End Select

            DG.GenerateSQL(Sql, ShowFilter, SecurityWhereClause, OrderByType, Request, cQueryCarrierSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"))

            ' ___ Eliminate order by clause from recordcount query
            If InfoType = "RecordCount" Then
                Pos = InStr(Sql, "ORDER BY", CompareMethod.Binary)
                If Pos > 0 Then
                    Sql = Sql.Substring(0, Pos - 1)
                End If
            End If

            If InfoType <> "Sql" Then
                dt = cCommon.GetDT(Sql, cEnviro.DBHost, "UserManagement", True)
            End If

            If InfoType = "RecordCount" Then
                Coll.Add(dt.Rows(0)(0).Value, "RecordCount")
            ElseIf InfoType = "Data" Then
                Coll.Add(dt, "Data")
            ElseIf InfoType = "Sql" Then
                Coll.Add(Sql, "Sql")
            End If
            Return Coll

        Catch ex As Exception
            Throw New Exception("Error #1108: QueryCarrier GetRecordCountOrData. " & ex.Message)
        End Try
    End Function
End Class