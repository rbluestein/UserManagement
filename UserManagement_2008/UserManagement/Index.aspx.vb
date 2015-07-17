Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

' document.getElementsByTagName("form")[0].submit()


Partial Class Index
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Private cLicSess As LicenseWorklistSession
    Private cUserSess As UserWorklistSession
#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PageMode As PageMode
        Dim Results As Results
        Dim Action As String
        Dim DG As DG

        Try

            cEnviro = Session("Enviro")
            cCommon = New Common


            ' ___ Load the application settings
            If Not cEnviro.Init Then
                LoadEnviro()
                cEnviro.MakeCookie(Me)
            Else
                'cEnviro.ValidateAppTimeout()
                cEnviro.AuthenticateRequest(Page)
            End If

            ' AdHoc()

            '  Testit()

            'cCommon.GenerateError()


            ' ___ Right check
            cRights = New RightsClass(Page)
            Dim RightsRqd(0) As String
            RightsRqd.SetValue(RightsClass.UserView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the user session object
            cUserSess = Session("UserWorklistSession")

            ' ___ Get the license session object
            cLicSess = Session("LicenseWorklistSession")

            ' ___ Get the page mode 
            PageMode = cCommon.GetPageMode(Page, cUserSess)

            ' ___ ClearDowloads
            ClearDownloads(PageMode)

            ' ___ ClearLog
            ClearLog(PageMode)

            ' ___ UpdateData
            cCommon.UpdateTimeSensitiveData(Nothing, Nothing)

            ' ___ Load the page session variables
            LoadVariables(PageMode)

            ' ___ Initialize the datagrid
            DG = DefineDataGrid()

            ' ___ Execute action
            Select Case PageMode
                Case PageMode.Initial
                    DisplayPage(PageMode, DG, DG.OrderByType.Initial)

                Case PageMode.Postback
                    Action = Request.Form("hdAction")
                    Select Case Action
                        Case "Sort"
                            DisplayPage(PageMode, DG, DG.OrderByType.Field, Request.Form("hdSortField"))

                        Case "ApplyFilter"
                            DisplayPage(PageMode, DG, DG.OrderByType.Recurring)

                        Case "Delete"
                            Results = DeleteRecord()
                            DisplayPage(PageMode, DG, DG.OrderByType.Recurring)
                            litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"

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
                    DisplayPage(PageMode, DG, DG.OrderByType.ReturnToPage)
                    If cUserSess.PageReturnOnLoadMessage <> Nothing Then
                        litMsg.Text = "<script language='javascript'>alert('" & cUserSess.PageReturnOnLoadMessage & "')</script>"
                        'litMsg.Text = LittleMessage()
                        cUserSess.PageReturnOnLoadMessage = Nothing
                    End If
            End Select

            ' ___ Record logged in user data
            cCommon.RecordLoggedInUserData(cEnviro.LoggedInUserID, cUserSess.SessionID, Request.ServerVariables("REMOTE_ADDR"))

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            'litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'/><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'/>"
            litEnviro.Text = cCommon.GetlitEnviroText(cEnviro.LoggedInUserID, cEnviro.DBHost)

        Catch ex As Exception
            Dim ErrorObj As New ErrorObj(ex, "Error #102: Index Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Sub AdHoc()
        Dim Users As DataTable
        Dim UserID As String
        Dim i As Integer
        Dim Querypack As DBase.QueryPack
        Dim ChangeDate As Date

        ChangeDate = cCommon.GetServerDateTime


        Querypack = cCommon.GetDTWithQueryPack("SELECT UserID FROM Users WHERE StatusCode = 'INACTIVE'")
        If Not Querypack.Success Then
            Stop
        End If

        Users = Querypack.dt
        For i = 0 To Users.Rows.Count - 1
            System.Diagnostics.Debug.WriteLine(i.ToString)
            UserID = Users.Rows(i)("UserID")
            Querypack = cCommon.GetDTWithQueryPack("Update UserAppointments set StatusCode = 'T', ChangeDate = '" & ChangeDate & "', ExpirationDate = IsNull(ExpirationDate, '1/14/2010') WHERE UserID = '" & UserID & "' AND StatusCode <> 'T'")
            If Not Querypack.Success Then
                Stop
            End If
        Next

        Querypack = cCommon.GetDTWithQueryPack("Update Users Set StatusCode = 'TERMINATED', LastStatusChangeDate = '" & ChangeDate & "', ChangeDate = '" & ChangeDate & "' WHERE StatusCode = 'INACTIVE'")
        If Not Querypack.Success Then
            Stop
        End If

    End Sub

    Private Sub ClearLog(ByVal PageMode As PageMode)
        Dim FileInfo As FileInfo
        Dim FileDays As Integer
        Dim NowDays As Integer

        Try

            If PageMode <> PageMode.Initial Then
                Exit Sub
            End If

            ' ___ Delete the log file if it is more than 10 days old..
            If System.IO.File.Exists(cEnviro.LogFileFullPath) Then

                FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)
                FileDays = FileInfo.CreationTimeUtc.Subtract(#1/1/2007#).Days()
                NowDays = Date.Now.Subtract(#1/1/2007#).Days()

                If NowDays - FileDays > 9 Then
                    Dim FileStream As New System.IO.FileStream(cEnviro.LogFileFullPath, FileMode.Create, FileAccess.Write, FileShare.None)
                    FileStream.Close()
                    FileInfo.CreationTime = Date.Now.ToUniversalTime.AddHours(-5)
                End If

            End If

        Catch ex As Exception
            Throw New Exception("Error #108: Index ClearLog. " & ex.Message)
        End Try
    End Sub

    Private Sub ClearDownloads(ByVal PageMode As PageMode)
        Dim i As Integer
        Dim DirInfo As DirectoryInfo
        Dim FileDays As Integer
        Dim NowDays As Integer
        Dim DownloadPathColl As Collection
        Dim FileArr As System.IO.FileInfo()
        Dim DirPath As String
        Dim FullPath As String

        Try

            If PageMode <> PageMode.Initial Then
                Exit Sub
            End If

            ' ___ Get a list of the download fiels
            DownloadPathColl = cCommon.GetDownloadPath(Page)
            DirPath = DownloadPathColl("DirPath")
            DirInfo = New DirectoryInfo(DirPath)
            FileArr = DirInfo.GetFiles

            ' ___ Delete the file if it is more than 1 days old.
            NowDays = Date.Now.Subtract(#1/1/2008#).Days()
            For i = 0 To FileArr.GetUpperBound(0)
                FullPath = DirPath & FileArr(i).Name
                FileDays = FileArr(i).CreationTimeUtc.Subtract(#1/1/2008#).Days()
                If NowDays - FileDays > 1 Then
                    File.Delete(FullPath)
                End If
            Next

        Catch ex As Exception
            Throw New Exception("Error #109: Index ClearDownloads. " & ex.Message)
        End Try
    End Sub

    Private Function LittleMessage() As String
        Dim sb As New System.Text.StringBuilder

        sb.Append("<script language='javascript'>")
        sb.Append("window.open('LittleMessage.aspx?Message=" & cUserSess.PageReturnOnLoadMessage & "', '','width=465,height=140,status=no,resizable=no,top=200,left=200,dependent=yes,alwaysRaised=yes')")
        sb.Append("</script>")
        Return sb.ToString

    End Function

    Private Sub LoadEnviro()
        Dim LoggedInUserID As String
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader

        Try

            cEnviro.Init = True
            cEnviro.SessionID = Guid.NewGuid.ToString
            cEnviro.LastPageLoad = Date.Now

            If System.Net.Dns.GetHostName().ToUpper = "WADEV" Then
                'If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
                'LoggedInUserID = "rbluestein"
                LoggedInUserID = "awickenheiser"
                'LoggedInUserID = "eagle"
                ' LoggedInUserID = "earchambeault"
                'LoggedInUserID = "rsmith"
                'LoggedInUserID = "jpenny"
                'LoggedInUserID = "bhake"
            Else
                LoggedInUserID = HttpContext.Current.User.Identity.Name.ToString
                LoggedInUserID = LoggedInUserID.Substring(InStr(LoggedInUserID, "\", CompareMethod.Binary))
            End If
            cEnviro.LoggedInUserID = LoggedInUserID

            cEnviro.DBHost = ConfigurationManager.AppSettings("DBHost").ToUpper

            cEnviro.DBHost = CType(configurationAppSettings.GetValue("DBHost", GetType(System.String)), String).ToUpper
            cEnviro.LoginIP = Request.ServerVariables("REMOTE_ADDR")
            cEnviro.ApplicationPath = Page.Server.MapPath(Page.Request.ApplicationPath)
            cEnviro.LogFileFullPath = Page.Server.MapPath(Page.Request.ApplicationPath) & "\UserManagementLogFile.txt"
            cEnviro.QueryDownloadDir = Page.Server.MapPath(Page.Request.ApplicationPath) & "\Temp"

        Catch ex As Exception
            Throw New Exception("Error #103: Index LoadEnviro. " & ex.Message)
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
                    cUserSess.UserID = Replace(Request.Form("hdUserID"), "~", "'")
                    cUserSess.UserIDFilter = Request.Form("txtUserID")
                    cUserSess.FullNameFilter = Request.Form("txtFullName")
                    cUserSess.StatusCodeFilter = Request.Form("ddStatusCode")
                    cUserSess.RoleFilter = Request.Form("ddRole")
                    cUserSess.CompanyIDFilter = Request.Form("ddCompanyID")
                    cUserSess.LocationIDFilter = Request.Form("ddLocationID")
                    cLicSess.UserID = Request.Form("hdUserID")
                    ' Note:  hdFilterShowHideToggle processed elsewhere
            End Select
        Catch ex As Exception
            Throw New Exception("Error #104: Index LoadVariables. " & ex.Message)
        End Try
    End Sub

    Private Function DeleteRecord() As Results
        Dim MyResults As New Results
        Dim QueryPack As DBase.QueryPack
        Dim Sql As String

        Try

            Sql = "exec usp_DeleteUser @UserID = '" & cUserSess.UserID & "'"
            QueryPack = cCommon.GetDTWithQueryPack(Sql, False)

            MyResults.Success = QueryPack.Success
            If QueryPack.Success Then
                MyResults.Msg = "Record deleted."
            Else
                MyResults.Msg = "Unable to delete record."
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #105: Index DeleteRecord. " & ex.Message)
        End Try



        'Sql = "DELETE Users WHERE UserID = '" & cUserSess.UserID & "'"
        'Querypack = cCommon.ExecuteNonQueryWithQuerypack(cEnviro.DBHost, cEnviro.DefaultDatabase, Sql)
        'MyResults.Success = Querypack.Success
        'If Querypack.Success Then
        '    MyResults.Msg = "Record deleted."
        'Else
        '    MyResults.Msg = "Unable to delete record."
        'End If
        'Return MyResults
    End Function

    Private Function DefineDataGrid() As DG
        Try

            Dim DG As New DG("UserID", cCommon, cRights, True, "EmbeddedTableDef", "UserID")
            DG.AddNewButton(RightsClass.UserEdit)
            DG.AddDataBoundColumn("UserID", "UserID", "User ID", "UserID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("FullName", "FullName", "Name", "LastName+FirstName+MI", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("StatusCode", "StatusCode", "Status", "StatusCode", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("Role", "Role", "Role", "Role", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("CompanyID", "CompanyID", "Company", "CompanyID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("LocationID", "LocationID", "Location", "LocationID", True, Nothing, Nothing, "align='left'")

            ' ___ Build the filter
            Dim Filter As DG.Filter
            Filter = DG.AttachFilter(DG.FilterOperationModeEnum.FilterSwitchable, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialHide)
            Filter.AddTextbox("UserID", "UserID", 50)
            Filter.AddTextbox("FullName", "FullName", 48, "LastName", Nothing)
            Filter.AddDropdown("StatusCode", "StatusCode")
            Filter.AddDropdown("Role", "Role")
            Filter.AddDropdown("CompanyID", "CompanyID")
            Filter.AddDropdown("LocationID", "LocationID")

            Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, False, Nothing, True)
            TemplateCol.AddDefaultTemplateItem("View", "ExistingRecord", "StandardView", "User record", RightsClass.UserView, Nothing)
            TemplateCol.AddDefaultTemplateItem("Permissions", "Permissions", "StandardClip", "Permissions", RightsClass.PermissionsView, Nothing)
            TemplateCol.AddDefaultTemplateItem("License", "License", "StandardNotes", "Licenses", RightsClass.LicenseView, "IsSupvOrEnroller", Nothing)
            TemplateCol.AddDefaultTemplateItem("Delete", "Delete", "StandardDelete", "Delete", RightsClass.UserEdit, Nothing, "FirstLastName")
            DG.AttachTemplateCol(TemplateCol)

            Return DG

        Catch ex As Exception
            Throw New Exception("Error #106: Index DefineDataGrid. " & ex.Message)
        End Try
    End Function

    Private Sub DisplayPage(ByVal PageMode As PageMode, ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Dim ViewOrDownload As String = "View"

        Try

            ' ___ Handle the filter
            HandleFilter(DG, PageMode)

            ' ___ Handle the sort
            If cUserSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cUserSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            ' ___ Handle the data
            HandleData(DG, PageMode, OrderByType, ViewOrDownload)

            ' ___ Set the FilterOnOffState
            cUserSess.FilterOnOffState = "on"

            ' ___ Set the last field sorted and sort direction in the sort reference
            cUserSess.SortReference = DG.GetSortReference

        Catch ex As Exception
            Throw New Exception("Error #110: Index DisplayPage. " & ex.Message)
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
                Filter.Coll("UserID").SetFilterValue(cUserSess.UserIDFilter)
                Filter.Coll("FullName").SetFilterValue(cUserSess.FullNameFilter)
            End If

            ' ___ Role
            dt = cCommon.GetDT("Select Role from Codes_Role Order By Role")
            Filter("Role").AddDropdownItem("", "", True)
            For i = 0 To dt.Rows.Count - 1
                Filter("Role").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
            Next
            If PageMode <> PageMode.Initial Then
                Filter.Coll("Role").SetFilterValue(cUserSess.RoleFilter)
            End If

            ' ___ StatusCode
            dt = cCommon.GetDT("Select Status from Codes_Status Order By Status")
            Filter("StatusCode").AddDropdownItem("", "", True)
            For i = 0 To dt.Rows.Count - 1
                Filter("StatusCode").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
            Next
            If PageMode <> PageMode.Initial Then
                Filter.Coll("StatusCode").SetFilterValue(cUserSess.StatusCodeFilter)
            End If

            ' ___ Company
            dt = cCommon.GetDT("Select ClientID from Codes_ClientID Order By ClientID")
            Filter("CompanyID").AddDropdownItem("", "", True)
            For i = 0 To dt.Rows.Count - 1
                Filter("CompanyID").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
            Next
            If PageMode <> PageMode.Initial Then
                Filter.Coll("CompanyID").SetFilterValue(cUserSess.CompanyIDFilter)
            End If

            ' ___ Location
            dt = cCommon.GetDT("Select LocationID from Codes_LocationID Order By LocationID")
            Filter("LocationID").AddDropdownItem("", "", True)
            For i = 0 To dt.Rows.Count - 1
                Filter("LocationID").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
            Next
            If PageMode <> PageMode.Initial Then
                Filter.Coll("LocationID").SetFilterValue(cUserSess.LocationIDFilter)
            End If

        Catch ex As Exception
            Throw New Exception("Error #111: Index HandleFilter. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleData(ByRef DG As DG, ByVal PageMode As PageMode, ByVal OrderByType As DG.OrderByType, ByVal ViewOrDownload As String)
        Dim dt As DataTable = Nothing
        Dim RecordCount As Integer
        Dim Coll As Collection
        Dim SuppressDisplayData As Boolean
        Dim EmbeddedMessage As String = Nothing
        Dim sb As New System.Text.StringBuilder
        Dim SPBuildExcel As Boolean
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
                cUserSess.InitialReportDataSuppressInEffect = True
            ElseIf PageMode = PageMode.Postback AndAlso cUserSess.InitialReportDataSuppressInEffect Then
                cUserSess.InitialReportDataSuppressInEffect = False
            End If
            If Not cUserSess.InitialReportDataSuppressInEffect Then
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
                If cUserSess.ExcessiveRecordsWarningInEffect Then
                    Coll = GetQueryInfo("Sql", DG, OrderByType)
                    If StrComp(Coll("Sql"), cUserSess.Sql, CompareMethod.Text) = 0 Then
                        IgnoreExcessiveRecords = True
                        PerformNextTest = False
                    Else
                        NewQuery = True
                    End If
                Else
                    NewQuery = True
                End If

                ' ___ Reset the excessive record warning properties
                If cUserSess.ExcessiveRecordsWarningInEffect Then
                    cUserSess.ExcessiveRecordsWarningInEffect = False
                    cUserSess.Sql = String.Empty
                End If
            End If

            ' ___ Test #4: Excessive records test
            If PerformNextTest Then
                If RecordCount > cEnviro.ExcessiveRecordAmount Then
                    Coll = GetQueryInfo("Sql", DG, OrderByType)
                    cUserSess.ExcessiveRecordsWarningInEffect = True
                    cUserSess.Sql = Coll("Sql")
                End If
            End If

            ' ___ DIRECT THE REPORT
            If ViewOrDownload = "Download" Then
                DownloadReport = True
            Else
                HTMLReport = True
            End If
            If cUserSess.InitialReportDataSuppressInEffect Then
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
            ElseIf cUserSess.ExcessiveRecordsWarningInEffect Then
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
                If SPBuildExcel Then
                    ' not used
                Else
                    cCommon.PrintCSVVersionLocal(dt, DownloadPathColl("AbsPath"))
                End If
            End If

            ' ___ Process the html
            If SuppressDisplayData Then
                dt = Nothing
            End If
            litDG.Text = DG.GetText(dt, EmbeddedMessage)

        Catch ex As Exception
            Throw New Exception("Error #112: Index DisplayPage. " & ex.Message)
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
                sb.Append("SELECT UserId, ")
                sb.Append("LTrim(RTrim(LastName)) + ', ' + LTrim(RTrim(FirstName))  FullName,  ")
                sb.Append("LTrim(RTrim(FirstName)) + ' ' + LTrim(RTrim(LastName))  FirstLastName, ")
                sb.Append("StatusCode, Role, CompanyID, LocationID, ")
                sb.Append("IsSupvOrEnroller = case when Role='ENROLLER' then '1' when Role='SUPERVISOR' then '1'   else '0' end  ")
            End If

            sb.Append("FROM Users ")
            Sql = sb.ToString

            ' ___ Where clause: Get the restrictions imposed by the security rules.
            SecurityWhereClause = cCommon.GetSecurityWhereClause(Nothing, AccessLevel, LocationID)

            'Select Case AccessLevel
            '    Case RightsClass.AccessLevelEnum.AllAccess
            '        SecurityWhereClause = String.Empty
            '    Case RightsClass.AccessLevelEnum.SupervisorAccess
            '        'SecurityWhereClause = " (Role = 'SUPERVISOR' or Role = 'ENROLLER') and LocationID = '" & LocationID & "'"
            '        SecurityWhereClause = " (UserID = '" & cEnviro.LoggedInUserID & " ' or Role = 'ENROLLER') and LocationID = '" & LocationID & "'"
            '    Case RightsClass.AccessLevelEnum.EnrollerAccess
            '        SecurityWhereClause = "UserID = '" & cEnviro.LoggedInUserID & "'"
            'End Select

            DG.GenerateSQL(Sql, ShowFilter, SecurityWhereClause, OrderByType, Request, cUserSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"))


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
            Throw New Exception("Error #113: Index GetRecordCountOrData. " & ex.Message)
        End Try
    End Function
End Class
