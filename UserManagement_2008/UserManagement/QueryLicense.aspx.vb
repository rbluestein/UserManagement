Imports System.Data
Imports System.Data.SqlClient

Partial Class QueryLicense
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Private cQueryLicenseSess As QueryLicenseSession
#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PageMode As PageMode
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
            RightsRqd.SetValue(RightsClass.LicenseView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the LicenseCarrier session object
            cQueryLicenseSess = Session("QueryLicenseSession")

            ' ___ Get the page mode 
            PageMode = cCommon.GetPageMode(Page, cQueryLicenseSess)

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
                    If cQueryLicenseSess.PageReturnOnLoadMessage <> Nothing Then
                        litMsg.Text = "<script language='javascript'>alert('" & cQueryLicenseSess.PageReturnOnLoadMessage & "')</script>"
                        cQueryLicenseSess.PageReturnOnLoadMessage = Nothing
                    End If
            End Select

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'>"

        Catch ex As Exception
            Throw New Exception("Error #1202: LicenseCarrier Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Sub LoadVariables(ByVal PageMode As PageMode)
        Try

            Select Case PageMode
                Case PageMode.Initial
                    cQueryLicenseSess.DateRange = "||||"
                Case PageMode.CalledByOther, PageMode.ReturnFromChild
                    ' No action
                Case PageMode.Postback
                    ' ___ Update session variables with those that the user may have changed
                    cQueryLicenseSess.UserID = Replace(Request.Form("txtul.UserID"), "~", "'")
                    cQueryLicenseSess.UserIDFilter = Request.Form("txtul.UserID")
                    cQueryLicenseSess.FullNameFilter = Request.Form("txtFullName")
                    cQueryLicenseSess.EmpStatusCodeFilter = Request.Form("ddu.StatusCode")
                    cQueryLicenseSess.LocationIDFilter = Request.Form("ddu.LocationID")
                    cQueryLicenseSess.StateFilter = Request.Form("ddul.State")
                    cQueryLicenseSess.CatgyFilter = Request.Form("ddCatgy")
                    cQueryLicenseSess.OKToRenewIndFilter = Request.Form("ddul.OKToRenewInd")
                    cQueryLicenseSess.DateRange = Request.Form("hdDateRange")
            End Select
        Catch ex As Exception
            Throw New Exception("Error #1204: QueryLicense LoadVariables. " & ex.Message)
        End Try
    End Sub

    Private Function DefineDataGrid() As DG
        Try

            Dim DG As New DG("UserID", cCommon, cRights, True, "EmbeddedTableDef", "ul.UserID")
            DG.AddDataBoundColumn("UserID", "UserID", "User<br>ID", "ul.UserID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("FullName", "FullName", "Name", "FullName", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("EmpStatusCode", "StatusCode", "Emp<br>Status", "u.StatusCode", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("LocationID", "LocationID", "Loc", "u.LocationID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("State", "State", "State", "ul.State", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("Catgy", "Catgy", "Lic<br>Status", "Catgy", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("LicenseNumber", "LicenseNumber", "Lic<br>Num", "ul.LicenseNumber", True, Nothing, Nothing, "align='left'")
            DG.AddDateColumn("ApplicationDate", "ApplicationDate", "App<br>Date", "ul.ApplicationDate", True, "MM/dd/yyyy", Nothing, "align='left'")
            DG.AddDateColumn("EffectiveDate", "EffectiveDate", "Eff<br>Date", "ul.EffectiveDate", True, "MM/dd/yyyy", Nothing, "align='left'")
            DG.AddDateColumn("RenewalDateSent", "RenewalDateSent", "Renew<br>Date<br>Sent", "ul.RenewalDateSent", True, "MM/dd/yyyy", Nothing, "align='left'")
            DG.AddDateColumn("RenewalDateRecd", "RenewalDateRecd", "Renew<br>Date<br>Recd", "ul.RenewalDateRecd", True, "MM/dd/yyyy", Nothing, "align='left'")
            DG.AddDateColumn("ExpirationDate", "ExpirationDate", "Exp<br>Date", "ul.ExpirationDate", True, "MM/dd/yyyy", Nothing, "align='left'")
            DG.AddFreeFormColumn("Spacer", "&nbsp;&nbsp", Nothing, Nothing, True, Nothing)
            DG.AddBooleanColumn("OKToRenewInd", "OKToRenewInd", "OK To<br>Renew", "ul.OKToRenewInd", True, "1", "Yes", "No", Nothing, "align='left'")
            'DG.AddFreeFormColumn("Spacer", "&nbsp;&nbsp", Nothing, Nothing, True, Nothing)
            DG.AddBooleanColumn("LongTermCareCert", "LongTermCareCert", "LTC<br>Cert", "LongTermCareCert", True, 1, "Yes", "No", Nothing, "align='left'")
            DG.AddDateColumn("LongTermCareCertExpirationDate", "LongTermCareCertExpirationDate", "LTC<br>Exp<br>Date", "u.LongTermCareCertExpirationDate", True, "MM/dd/yyyy", Nothing, "align='left'")
            DG.AddDateColumn("LongTermCareStateSpecificExpirationDate", "LongTermCareStateSpecificExpirationDate", "LTC<br>St Spec Exp<br>Date", "ul.LongTermCareStateSpecificExpirationDate", True, "MM/dd/yyyy", Nothing, "align='left'")


            ' ___ Build the menu
            Dim Menu As DG.Menu
            Menu = DG.AttachMenu(7)
            Menu.AddItem(DG.Menu.ObjectTypeEnum.IsButton, "View", "View", RightsClass.LicenseView)
            Menu.AddItem(DG.Menu.ObjectTypeEnum.IsButton, "Download", "Download", RightsClass.LicenseView)

            ' ___ Build the filter
            Dim Filter As DG.Filter
            Filter = DG.AttachFilter(DG.FilterOperationModeEnum.FilterAlwaysOn, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialHide)
            Filter.AddTextbox("UserID", "ul.UserID", 50)
            Filter.AddTextbox("FullName", "FullName", 48, "LTrim(RTrim(u.LastName)) + ', ' + LTrim(RTrim(u.FirstName))", Nothing)
            Filter.AddDropdown("EmpStatusCode", "u.StatusCode")
            Filter.AddDropdown("LocationID", "u.LocationID")
            Filter.AddDropdown("State", "ul.State")
            Filter.AddExtendedDropdown("Catgy", "Catgy")
            Filter.AddDateCtl("ExpirationDate", "ul.ExpirationDate", "Get Date", "GetDateRange")
            Filter.AddDropdown("OKToRenewInd", "ul.OKToRenewInd")

            DG.SecondaryOrderBy = "ul.State"

            Return DG

        Catch ex As Exception
            Throw New Exception("Error #1206: QueryLicense DefineDataGrid. " & ex.Message)
        End Try
    End Function

    Private Sub DisplayPage(ByVal PageMode As PageMode, ByVal ViewOrDownload As String, ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Try

            ' ___ Handle the filter
            HandleFilter(DG, PageMode)

            ' ___ Handle the sort
            If cQueryLicenseSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cQueryLicenseSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            ' ___ Handle the data
            HandleData(DG, PageMode, OrderByType, ViewOrDownload)

            ' ___ Set the FilterOnOffState
            cQueryLicenseSess.FilterOnOffState = "on"

            ' ___ Set the last field sorted and sort direction in the sort reference
            cQueryLicenseSess.SortReference = DG.GetSortReference

            '  ___ Set up the interface to the expiration date javascript filter control
            litFilterHiddens.Text = "<input type='hidden' name='hdDateRange' id='hdDateRange' value='" & cQueryLicenseSess.DateRange & "'>"

        Catch ex As Exception
            Throw New Exception("Error #1207: QueryLicense DisplayPage. " & ex.Message)
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
                Filter.Coll("UserID").SetFilterValue(cQueryLicenseSess.UserIDFilter)
                Filter.Coll("FullName").SetFilterValue(cQueryLicenseSess.FullNameFilter)
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
                Filter.Coll("EmpStatusCode").SetFilterValue(cQueryLicenseSess.EmpStatusCodeFilter)
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
                Filter.Coll("LocationID").SetFilterValue(cQueryLicenseSess.LocationIDFilter)
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
                Filter.Coll("State").SetFilterValue(cQueryLicenseSess.StateFilter)
            End If

            ' ___ Catgy
            Filter("Catgy").AddExtendedDropdownItem("", "", "", True)
            'Filter("Catgy").AddExtendedDropdownItem("1", "Application", " ul.EffectiveDate = '1/1/1950' ", False)
            'Filter("Catgy").AddExtendedDropdownItem("2", "Effective", " ul.EffectiveDate <> '1/1/1950' ", False)
            'Filter("Catgy").AddExtendedDropdownItem("3", "Expired", " ul.ExpirationDate < getDate() ", False)

            Filter("Catgy").AddExtendedDropdownItem("1", "Application", " ul.EffectiveDate = '1/1/1950' ", False)
            Filter("Catgy").AddExtendedDropdownItem("2", "Effective", " ul.EffectiveDate <> '1/1/1950' AND (ul.ExpirationDate Is Null OR ul.ExpirationDate > getDate()) ", False)
            Filter("Catgy").AddExtendedDropdownItem("3", "Expired", " ul.ExpirationDate Is Not Null AND DateAdd(day, 1, ul.ExpirationDate) <= getDate() ", False)

            ' ___ OKToRenewInd
            Filter("OKToRenewInd").AddDropdownItem("", "")
            Filter("OKToRenewInd").AddDropdownItem("1", "Yes")
            Filter("OKToRenewInd").AddDropdownItem("0", "No")
            If PageMode <> PageMode.Initial Then
                Filter.Coll("OKToRenewInd").SetFilterValue(cQueryLicenseSess.OKToRenewIndFilter)
            End If


            If PageMode <> PageMode.Initial Then
                Filter.Coll("Catgy").SetFilterValue(cQueryLicenseSess.CatgyFilter)
            End If

            ' ___ Expiration date
            Filter.Coll("ExpirationDate").SetFilterValue(cQueryLicenseSess.DateRange)

        Catch ex As Exception
            Throw New Exception("Error #1209: QueryLicense HandleFilter. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleData(ByRef DG As DG, ByVal PageMode As PageMode, ByVal OrderByType As DG.OrderByType, ByVal ViewOrDownload As String)
        Dim dt As DataTable = Nothing
        Dim RecordCount As Integer
        Dim Coll As Collection
        Dim SuppressDisplayData As Boolean
        Dim EmbeddedMessage As String = Nothing
        Dim DownloadPathColl As Collection = Nothing
        Dim HTMLReport As Boolean
        Dim DownloadReport As Boolean
        Dim ExceedsRecordMaximum As Boolean
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
                cQueryLicenseSess.InitialReportDataSuppressInEffect = True
            ElseIf PageMode = PageMode.Postback AndAlso cQueryLicenseSess.InitialReportDataSuppressInEffect Then
                cQueryLicenseSess.InitialReportDataSuppressInEffect = False
            End If
            If Not cQueryLicenseSess.InitialReportDataSuppressInEffect Then
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
                If cQueryLicenseSess.ExcessiveRecordsWarningInEffect Then
                    Coll = GetQueryInfo("Sql", DG, OrderByType)
                    If StrComp(Coll("Sql"), cQueryLicenseSess.Sql, CompareMethod.Text) = 0 Then
                        IgnoreExcessiveRecords = True
                        PerformNextTest = False
                    Else
                        NewQuery = True
                    End If
                Else
                    NewQuery = True
                End If

                ' ___ Reset the excessive record warning properties
                If cQueryLicenseSess.ExcessiveRecordsWarningInEffect Then
                    cQueryLicenseSess.ExcessiveRecordsWarningInEffect = False
                    cQueryLicenseSess.Sql = String.Empty
                End If
            End If

            ' ___ Test #4: Excessive records test
            If PerformNextTest Then
                If RecordCount > cEnviro.ExcessiveRecordAmount Then
                    Coll = GetQueryInfo("Sql", DG, OrderByType)
                    cQueryLicenseSess.ExcessiveRecordsWarningInEffect = True
                    cQueryLicenseSess.Sql = Coll("Sql")
                End If
            End If

            ' ___ DIRECT THE REPORT
            If ViewOrDownload = "Download" Then
                DownloadReport = True
            Else
                HTMLReport = True
            End If
            If cQueryLicenseSess.InitialReportDataSuppressInEffect Then
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
            ElseIf cQueryLicenseSess.ExcessiveRecordsWarningInEffect Then
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

    Private Function GetQueryInfo(ByVal InfoType As String, ByVal DG As DG, ByVal OrderByType As DG.OrderByType) As Collection
        Dim sb As New System.Text.StringBuilder
        Dim Sql As String
        Dim dt As DataTable = Nothing
        Dim ShowFilter As Boolean
        Dim Coll As New Collection
        Dim Pos As Integer
        Dim SecurityWhereClause As String
        Dim AccessLevel As RightsClass.AccessLevelEnum
        Dim Role As String = Nothing
        Dim LocationID As String = Nothing

        Try

            ' ___ Get the security settings for the logged in user
            cRights.GetSecurityFlds(AccessLevel, Role, LocationID)

            If InfoType = "RecordCount" Then
                sb.Append("SELECT Count (*) ")

            ElseIf InfoType = "Data" Or InfoType = "Sql" Then
                sb.Append("SELECT ul.UserId, ")
                sb.Append("LTrim(RTrim(u.LastName)) + ', ' + LTrim(RTrim(u.FirstName))  FullName,  ")
                sb.Append("u.StatusCode, ")
                'sb.Append("EmpStatusStr = case when u.StatusCode='ACTIVE' then 'Active' else 'Inactive' end, ")
                sb.Append("EmpStatusStr = case when u.StatusCode='ACTIVE' then 'Active' when u.StatusCode='TERMINATED' then 'Terminated' else 'Inactive' end, ")
                sb.Append("u.LocationID, ")
                sb.Append("u.LongTermCareCertExpirationDate, ")
                sb.Append("ul.State, ")
                sb.Append("ul.LongTermCareStateSpecificExpirationDate, ")


                'Filter("Catgy").AddExtendedDropdownItem("1", "Application", " ul.EffectiveDate = '1/1/1950' ", False)
                'Filter("Catgy").AddExtendedDropdownItem("2", "Effective", " ul.EffectiveDate <> '1/1/1950' AND (ul.ExpirationDate Is Null OR ul.ExpirationDate > getDate()) ", False)
                'Filter("Catgy").AddExtendedDropdownItem("3", "Expired", " ul.ExpirationDate Is Not Null AND DateAdd(day, 1, ul.ExpirationDate) <= getDate() ", False)

                sb.Append("Catgy = case ")
                sb.Append(" when ul.EffectiveDate = '1/1/1950' then 'Application' ")
                sb.Append(" when ul.EffectiveDate <> '1/1/1950' AND (ul.ExpirationDate Is Null OR ul.ExpirationDate > getDate())  then 'Effective' ")
                sb.Append(" when ul.ExpirationDate Is Not Null AND DateAdd(day, 1, ul.ExpirationDate) <= getDate()  then 'Expired' ")
                ' sb.Append(" when ul.ExpirationDate < getDate() then 'Expired' ")
                sb.Append(" else '' ")
                sb.Append(" end, ")

                sb.Append("ul.LicenseNumber, ")
                sb.Append("ul.ApplicationDate, ")
                sb.Append("ul.EffectiveDate, ")
                sb.Append("ul.RenewalDateSent, ")
                sb.Append("ul.RenewalDateRecd, ")
                sb.Append("ul.ExpirationDate, ")
                sb.Append("ul.OKToRenewInd, ")
                sb.Append("LongTermCareCert = dbo.ufn_getLTCCert(getDate(), u.UserID, ul.State) ")
            End If

            sb.Append("FROM UserLicenses ul ")
            sb.Append("INNER JOIN Users u on ul.UserID = u.UserID ")
            sb.Append("INNER JOIN Codes_State cs on ul.State = cs.StateCode")
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


            DG.GenerateSQL(Sql, ShowFilter, SecurityWhereClause, OrderByType, Request, cQueryLicenseSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"))

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
            Throw New Exception("Error #1208: QueryLicense GetRecordCountOrData. " & ex.Message)
        End Try
    End Function

    Private Function SqlIndFalse(ByVal IndName As String) As String
        Return "(" & IndName & " is null or " & IndName & " = 0)"
    End Function

    Private Function SqlIndTrue(ByVal IndName As String, ByVal FromDate As String, ByVal ToDate As String) As String
        Return " (" & IndName & " = 1 and " & FromDate & " is not null and " & FromDate & " <= getDate() and " & ToDate & " is not null and DateAdd(Day, 1,  " & ToDate & ") >= getDate()) "
    End Function

    Private Function SqlLicTest() As String
        'Return "(ul.EffectiveDate is not null and ul.EffectiveDate <= getDate() and (ul.ExpirationDate is null or ul.ExpirationDate >= getDate()))"
        Return "(ul.EffectiveDate is not null and ul.EffectiveDate <= getDate() and (ul.ExpirationDate is null or DateAdd(Day, 1, ul.ExpirationDate) >= getDate()))"
    End Function
End Class