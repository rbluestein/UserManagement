Imports System.Data.SqlClient

Public Class LicenseWorklist
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected PageCaption As HtmlGenericControl
    Protected WithEvents litDG As System.Web.UI.WebControls.Literal
    Protected WithEvents dgOrgList As System.Web.UI.WebControls.DataGrid
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents litFilterHiddens As System.Web.UI.WebControls.Literal
    Protected WithEvents litHeading As System.Web.UI.WebControls.Literal
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Protected WithEvents litKeyValues As System.Web.UI.WebControls.Literal
    Protected WithEvents lblEnrollerName As System.Web.UI.WebControls.Label
    Private DGAppointments As DG
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
    Private cSess As LicenseWorklistSession
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

    ' ___ Event raised by the data grid.
    Public Sub HandleChildDTRequest(ByRef ChildText As String, ByVal DataFldName As String, ByVal Value As String)
        Dim Sql As String
        Dim dt As DataTable


        'dt = cCommon.GetDT("SELECT *, Catgy = case when EffectiveDate is null then 'Pending' else 'Appointed'  end  from UserAppointments WHERE State =  '" & Value & "' AND UserID = '" & cSess.UserID & "'")
        'Sql = "SELECT *, Catgy = case when EffectiveDate is null then 'Pending' else 'Appointed'  end  from UserAppointments WHERE State =  '" & Value & "' AND UserID = '" & cSess.UserID & "'"
        Sql = "SELECT *, Catgy = case when StatusCode = 'P' then 'Pending' when StatusCode = 'X' then 'Appointed' when StatusCode = 'T' then 'Terminated' else ''  end  from UserAppointments WHERE State =  '" & Value & "' AND UserID = '" & cSess.UserID & "'"
        dt = cCommon.GetDT(Sql, cEnviro.DBHost, "UserManagement", True)


        ChildText = DGAppointments.GetText(dt, Nothing)
        cSess.SubTableInd = "1"
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PageMode As PageMode
        Dim Results As Results
        Dim Action As String
        Dim DG As DG

        Try

            ' ___ Instantiate objects
            cEnviro = Session("Enviro")
            cCommon = New Common

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right Check
            Dim RightsRqd(0) As String
            cRights = New RightsClass(Page)
            RightsRqd.SetValue(cRights.LicenseView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the license worklist session object
            cSess = Session("LicenseWorklistSession")

            ' ___ Get the page mode 
            PageMode = cCommon.GetPageMode(Page, cSess)

            ' ___ Load the page session variables
            LoadVariables(PageMode)

            ' ___ Initialize the datagrid
            DG = DefineDataGrid("state")
            AddHandler DG.ChildDTRequest, AddressOf HandleChildDTRequest
            DGAppointments = DefineDataGrid("appointments")

            If cSess.SubTableInd = "1" Then
                DG.ChildTables.ChildTableSelectColumn.ParmColl("DataFldName").Value = cSess.SubTableState
            End If

            ' ___ Execute action
            Select Case PageMode
                Case PageMode.Initial
                    DisplayPage(PageMode, DG, DG.OrderByType.Initial)

                Case PageMode.Postback
                    Action = Request.Form("hdAction")
                    Select Case Action
                        Case "return"
                            Response.Redirect("Index.aspx?CalledBy=Child")

                        Case "Sort"
                            DisplayPage(PageMode, DG, DG.OrderByType.Field, Request.Form("hdSortField"))

                        Case "ApplyFilter"
                            DisplayPage(PageMode, DG, DG.OrderByType.Recurring)

                        Case "DeleteLic"
                            Results = DeleteLicense()
                            DisplayPage(PageMode, DG, DG.OrderByType.Recurring)
                            litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"

                        Case "DeleteAppt"
                            Results = DeleteAppointment()
                            DisplayPage(PageMode, DG, DG.OrderByType.Recurring)
                            litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"

                        Case "NewLicense"
                            Response.Redirect("License.aspx?CallType=New")

                        Case "ExistingLicense"
                            Response.Redirect("License.aspx?CallType=Existing")

                        Case "NewAppointment"
                            Response.Redirect("Appointment.aspx?CallType=New")

                        Case "ExistingAppointment"
                            Response.Redirect("Appointment.aspx?CallType=Existing")

                        Case "ShowHideSubTable"
                            DisplayPage(PageMode, DG, DG.OrderByType.Recurring)

                    End Select

                Case PageMode.ReturnFromChild, PageMode.CalledByOther
                    DisplayPage(PageMode, DG, DG.OrderByType.ReturnToPage)
                    If cSess.PageReturnOnLoadMessage <> Nothing Then
                        litMsg.Text = "<script language='javascript'>alert('" & cSess.PageReturnOnLoadMessage & "')</script>"
                        cSess.PageReturnOnLoadMessage = Nothing
                    End If
            End Select

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'>"

        Catch ex As Exception
            Throw New Exception("Error #402: LicenseWorklist Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Sub LoadVariables(ByVal PageMode As PageMode)

        Try

            ' ___ Page called by:
            ' 1. Index.aspx either as a new or existing record.
            '    PageMode: Initial upon the very first invocation, CalledByOther thereafter.
            ' 2. Child: License.aspx. PageMode: CalledByChild.
            ' 3. Postback. PageMode: Postback.
            ' 4. Page is NOT called by main menu.

            Select Case PageMode
                Case PageMode.Initial

                    ' ___ Active record
                    'LoadQueryStringValues()

                    ' ___ Initialize SubTable
                    cSess.SubTableInd = "0"

                Case PageMode.CalledByOther

                    ' ___ Active record
                    'LoadQueryStringValues()

                    '' ___ Filter selections
                    'If Sess.LicenseNumber = String.Empty Then
                    '    Sess.StateFilter = String.Empty
                    '    Sess.CatgyFilter = String.Empty
                    'End If

                Case PageMode.ReturnFromChild
                    ' No action
                Case PageMode.Postback

                    ' ___ Update session variables with those that the user may have changed

                    ' ___ Active
                    'Sess.UserID = Request.Form("hdUserID")
                    cSess.State = Request.Form("hdState")
                    cSess.StateDropdown = Request.Form("hdState")
                    cSess.LicenseNumber = Request.Form("hdLicenseNumber")
                    cSess.EffectiveDate = Request.Form("hdEffectiveDate")
                    cSess.CarrierID = Request.Form("hdCarrierID")
                    cSess.AppointmentNumber = Request.Form("hdAppointmentNumber")

                    ' ___ SubTable
                    cSess.SubTableInd = Request.Form("hdSubTableInd")
                    cSess.SubTableState = Request.Form("hdSubTableState")

                    ' ___ Filter
                    cSess.StateFilter = Request.Form("ddState")
                    cSess.CatgyFilter = Request.Form("ddCatgy")
                    ' cFilterOnResponse = Request.Form("hdFilterOnResponse")                ' not changed in client
                    ' Note:  hdFilterShowHideToggle processed elsewhere

            End Select

        Catch ex As Exception
            Throw New Exception("Error #403: LicenseWorklist LoadVariables. " & ex.Message)
        End Try
    End Sub

    Private Function DeleteLicense()
        Dim MyResults As New Results
        Dim QueryPack As DBase.QueryPack
        Dim Sql As String

        Try

            Sql = "exec usp_DeleteLicense @UserID = '" & cSess.UserID & "', @State = '" & cSess.State & "'"
            QueryPack = cCommon.GetDTWithQueryPack(Sql, False)

            MyResults.Success = QueryPack.Success
            If QueryPack.Success Then
                MyResults.Msg = "Record deleted."
            Else
                MyResults.Msg = "Unable to delete record."
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #404: LicenseWorklist DeleteLicense. " & ex.Message)
        End Try


        'Sql = "DELETE Users WHERE UserID = '44'"       ' FIX THIS !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1
        'Querypack = cCommon.ExecuteNonQueryWithQuerypack(cEnviro.DBHost, cEnviro.DefaultDatabase, Sql)
        'MyResults.Success = Querypack.Success
        'If Querypack.Success Then
        '    MyResults.Msg = "Record deleted."
        'Else
        '    MyResults.Msg = "Unable to delete record."
        'End If
        'Return MyResults






        '''Dim MyResults As New Results
        '''Dim QueryPack As DBase.QueryPack
        '''CmdAsst.AddVarChar("UserID", cSess.UserID, 50)
        '''CmdAsst.AddVarChar("State", cSess.State, 2)
        '''CmdAsst.AddVarChar("LicenseNumber", cSess.LicenseNumber, 50)
        '''Try
        '''    QueryPack = CmdAsst.Execute
        '''Catch
        '''    QueryPack.Success = False
        '''End Try
        '''MyResults.Success = QueryPack.Success
        '''If QueryPack.Success Then
        '''    MyResults.Msg = "Record deleted."
        '''Else
        '''    MyResults.Msg = "Unable to delete record."
        '''End If
        '''Return MyResults
    End Function

    Private Function DeleteAppointment() As Results
        Dim MyResults As New Results
        'Dim SqlConnection1 As New SqlClient.SqlConnection(cEnviro.GetConnectionString)
        'Dim SqlCmd As System.Data.SqlClient.SqlCommand
        Dim Sql As String

        Try

            MyResults.Success = True
            MyResults.Msg = "Record deleted."

            'SqlConnection1.Open()

            Try
                Sql = "DELETE FROM UserAppointments WHERE UserID = '" & cSess.UserID & "' AND State='" & cSess.State & "' AND CarrierID='" & cSess.CarrierID & "' AND AppointmentNumber='" & cSess.AppointmentNumber & "'"
                'SqlCmd = New System.Data.SqlClient.SqlCommand(Sql, SqlConnection1)
                'SqlCmd.CommandType = System.Data.CommandType.Text
                'SqlCmd.ExecuteNonQuery()
                cCommon.ExecuteNonQuery(Sql)
                'SqlCmd.Dispose()
                cSess.State = String.Empty
                cSess.CarrierID = String.Empty
                cSess.EffectiveDate = String.Empty
                cSess.ExpirationDate = String.Empty
            Catch
                MyResults.Success = False
                MyResults.Msg = "Problem deleting record. No deletions made."
            End Try

            'SqlConnection1.Close()
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #405: LicenseWorklist DeleteAppointment. " & ex.Message)
        End Try
    End Function

    Private Function DefineDataGrid(ByVal Entity As String) As DG
        Dim DG As DG

        Try

            If Entity = "state" Then
                DG = New DG("State", cCommon, cRights, True, "EmbeddedTableDef", "StateEffectiveDate")
                DG.AddDataBoundColumn("State", "State", "State", "StateEffectiveDate", True, Nothing, Nothing, "align='left'")
                DG.AddDataBoundColumn("LicenseNumber", "LicenseNumber", "License<br>Number", "LicenseNumber", True, Nothing, Nothing, "align='left'")
                DG.AddDateColumn("EffectiveDate", "EffectiveDate", "Effective<br>Date", "EffectiveDate", True, "MM/dd/yyyy", Nothing, "align='left'")
                DG.AddDateColumn("ExpirationDate", "ExpirationDate", "Expiration<br>Date", "ExpirationDate", True, "MM/dd/yyyy", Nothing, "align='left'")
                DG.AddBooleanColumn("OKToRenewInd", "OKToRenewInd", "OK To<br>Renew", "OKToRenewInd", True, "1", "Yes", "No", Nothing, "align='left'")
                DG.AddDataBoundColumn("Catgy", "Catgy", "Category", "Catgy", True, Nothing, Nothing, "align='left'")
                DG.AddBooleanColumn("LongTermCareCert", "LongTermCareCert", "LTC<br>Cert", "LongTermCareCert", True, 1, "Yes", "No", Nothing, "align='left'")
                DG.AddDataBoundColumn("NumAppointments", "NumAppointments", "Num<br>Appts", "NumAppointments", True, Nothing, Nothing, "align='left'")

                ' ___ Build the menu
                Dim Menu As DG.Menu
                Menu = DG.AttachMenu(10)
                Menu.AddItem(DG.Menu.ObjectTypeEnum.IsLink, "NewLicense", "New", cRights.LicenseEdit)
                'Menu.AddItem(DG.Menu.ObjectTypeEnum.IsLink, "NewAppointment", "New Appointment", Rights.LicenseEdit)

                Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, False, Nothing, True)
                TemplateCol.AddDefaultTemplateItem("View", "ExistingLicense", "StandardView", "License Detail", cRights.LicenseView, Nothing, "LicenseNumber", "EffectiveDate")
                TemplateCol.AddDefaultTemplateItem("Appointment", "NewAppointment", "StandardClip", "New Appointment", cRights.LicenseEdit, Nothing)
                TemplateCol.AddDefaultTemplateItem("DeleteLic", "DeleteLic", "StandardDelete", "Delete", cRights.LicenseEdit, Nothing, "LicenseNumber", "EffectiveDate")
                DG.AttachTemplateCol(TemplateCol)

                Dim ChildTables As DG.ChildTablesClass
                ChildTables = DG.AttachChildTables("State", "State", "Active")
                DG.AddChildTableSelectColumn("StateSubTable", "State", "Active", Nothing, Nothing)

                ' ___ Build the filter
                Dim Filter As DG.Filter
                Filter = DG.AttachFilter(DG.FilterOperationMode.FilterSwitchable, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialShow)
                Filter.AddDropdown("State", "State")
                Filter.AddExtendedDropdown("Catgy", "Catgy")

            ElseIf Entity = "appointments" Then
                DG = New DG("State", cCommon, cRights, True, "EmbeddedTableDef", "CarrierID")
                DG.AddDataBoundColumn("CarrierID", "CarrierID", "Car ID", Nothing, True, Nothing, Nothing, "align='left'")
                DG.AddDataBoundColumn("AppointmentNumber", "AppointmentNumber", "Appt Num", Nothing, True, Nothing, Nothing, "align='left'")
                DG.AddDateColumn("EffectiveDate", "EffectiveDate", "Eff Date", Nothing, True, "MM/dd/yyyy", Nothing, "align='left'")
                'DG.AddDateColumn("ExpirationDate", "ExpirationDate", "Exp Date", Nothing, True, "MM/dd/yyyy", Nothing, "align='left'")
                'DG.AddDataBoundColumn("JIT", "JIT", "JIT", Nothing, True, Nothing, Nothing, "align='left'")
                DG.AddDataBoundColumn("Catgy", "Catgy", "Category", Nothing, True, Nothing, Nothing, "align='left'")

                Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, False, Nothing, True)
                TemplateCol.AddDefaultTemplateItem("View", "ExistingAppointment", "StandardView", "Appointment Detail", cRights.LicenseView, Nothing, "CarrierID", "EffectiveDate", "AppointmentNumber")
                TemplateCol.AddDefaultTemplateItem("DeleteAppt", "DeleteAppt", "StandardDelete", "Delete", cRights.LicenseEdit, Nothing, "CarrierID", "EffectiveDate", "AppointmentNumber")
                DG.AttachTemplateCol(TemplateCol)

                DG.AddFreeFormColumn("Spacer", Nothing, Nothing, Nothing, True, "width='80px'")
                DG.FormatAsSubTable = True
            End If

            Return DG

        Catch ex As Exception
            Throw New Exception("Error #406: LicenseWorklist DefineDataGrid. " & ex.Message)
        End Try
    End Function

    Private Sub DisplayPage(ByVal PageMode As PageMode, ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Dim i As Integer
        Dim dt As DataTable
        Dim Sql As String
        Dim sb As New System.Text.StringBuilder
        Dim WhereClause As New System.Text.StringBuilder
        Dim ShowFilter As Boolean
        Dim ShowRecords As Boolean

        Try

            ' ___ Get a filter reference
            Dim Filter As DG.Filter
            Filter = DG.GetFilter

            ' ___ State
            dt = cCommon.GetDT("SELECT StateCode FROM Codes_State ORDER BY StateCode")
            Filter("State").AddDropdownItem("", "", True)
            For i = 0 To dt.Rows.Count - 1
                Filter("State").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
            Next
            If PageMode <> PageMode.Initial Then
                Filter.Coll("State").SetFilterValue(cSess.StateFilter)
            End If

            ' ___ Catgy
            Filter("Catgy").AddExtendedDropdownItem("", "", "", True)
            Filter("Catgy").AddExtendedDropdownItem("1", "Application", " ul1.EffectiveDate = '1/1/1950' ", False)
            Filter("Catgy").AddExtendedDropdownItem("2", "Effective", " ul1.EffectiveDate <> '1/1/1950' ", False)
            Filter("Catgy").AddExtendedDropdownItem("3", "Expired", " ul1.ExpirationDate < getDate() ", False)
            If PageMode <> PageMode.Initial Then
                Filter.Coll("Catgy").SetFilterValue(cSess.CatgyFilter)
            End If
            'Filter("Catgy").AddExtendedDropdownItem("2", "Effective", " (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSubjUserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate ", False)
            '  Filter("Catgy").AddExtendedDropdownItem("3", "Superseded", " ul1.EffectiveDate is not null and (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSubjUserID & "' and ul2.state =  ul1.State) <> ul1.EffectiveDate ", False)

            ' ___ Heading/UserID
            If cRights.HasThisRight(cRights.LicenseEdit) Then
                litHeading.Text = "Edit License"
            Else
                litHeading.Text = "View License"
            End If

            ' ___ Enroller name heading
            'Sql = "SELECT LastName + ', ' +  FirstName + ' ' + MI FullName FROM Users WHERE UserID ='" & cSess.UserID & "'"
            dt = cCommon.GetDT("SELECT LastName + ', ' +  FirstName + ' ' + MI FullName FROM Users WHERE UserID ='" & cSess.UserID & "'")
            lblEnrollerName.Text = "Name:&nbsp;&nbsp;" & dt.Rows(0)("FullName")

            ' ___ Handle the sort
            If cSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            ' ___ Load the parameters and execute the query
            'Sql = "SELECT UserID, State, LicenseNumber, EffectiveDate, ExpirationDate, LongTermCareInd, Selected = case when 0=1 then '1' else '0' end, Unselected = '0', Notes FROM UserLicenses WHERE UserID ='" & cSubjUserID & "'"

            'Sql = "SELECT ul1.UserID, ul1.State, ul1.LicenseNumber, ul1.EffectiveDate, ul1.ExpirationDate, ul1.LongTermCareInd, "
            'Sql &= " StateEffectiveDate = ul1.State + Cast(100 + DATEPART(Year, ul1.EffectiveDate) as varchar(3)) + Cast(100 + DATEPART(Month, ul1.EffectiveDate) as varchar(3)) + Cast(100 + DATEPART(Day, ul1.EffectiveDate) as varchar(3)), "
            'Sql &= "Active = case when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSubjUserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then '1' else '0' end, "
            'Sql &= "NumAppointments = case when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSubjUserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then Cast((Select Count (*) FROM UserAppointments WHERE UserID =ul1.UserID AND State=ul1.State) as varchar(10)) else '' end, "
            'Sql &= "ul1.Notes FROM UserLicenses ul1 WHERE ul1.UserID ='" & cSubjUserID & "'"

            'Sql = "SELECT ul1.UserID, ul1.State, ul1.LicenseNumber, ul1.EffectiveDate, ul1.ExpirationDate, ul1.LongTermCareInd, "
            'Sql &= " StateEffectiveDate = ul1.State + Cast(100 + DATEPART(Year, ul1.EffectiveDate) as varchar(3)) + Cast(100 + DATEPART(Month, ul1.EffectiveDate) as varchar(3)) + Cast(100 + DATEPART(Day, ul1.EffectiveDate) as varchar(3)), "
            'Sql &= "Catgy = case "
            'Sql &= " when ul1.EffectiveDate is null then 'Pending' "
            'Sql &= " when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSubjUserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then 'Last License' "
            'Sql &= " else 'Superseded' end, "

            'Sql = "SELECT ul1.UserID, ul1.State, ul1.LicenseNumber, ul1.EffectiveDate, ul1.ExpirationDate, ul1.LongTermCareInd, "
            'Sql &= " StateEffectiveDate = ul1.State + Cast(100 + DATEPART(Year, ul1.EffectiveDate) as varchar(3)) + Cast(100 + DATEPART(Month, ul1.EffectiveDate) as varchar(3)) + Cast(100 + DATEPART(Day, ul1.EffectiveDate) as varchar(3)), "
            'Sql &= "Active = case when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSubjUserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then '1' else '0' end, "
            'Sql &= "Catgy = case "
            'Sql &= " when ul1.EffectiveDate is null then 'Pending' "
            'Sql &= " when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSubjUserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then 'Last License' "
            'Sql &= " else 'Superseded' end, "
            'Sql &= " NumAppointments = case when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSubjUserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then Cast((Select Count (*) FROM UserAppointments WHERE UserID =ul1.UserID AND State=ul1.State) as varchar(10)) else '' end, "
            'Sql &= "ul1.Notes FROM UserLicenses ul1 WHERE ul1.UserID ='" & cSubjUserID & "'"

            sb.Append("SELECT ul1.UserID, ul1.State, ul1.LicenseNumber, ul1.EffectiveDate, ul1.ExpirationDate, ul1.OKToRenewInd, ")

            ' used for sorting
            sb.Append(" StateEffectiveDate = case ")
            sb.Append(" when ul1.EffectiveDate is null then ul1.State")
            sb.Append(" else ul1.State + Cast(DATEPART(Year, ul1.EffectiveDate) as varchar(4)) + Cast(100 + DATEPART(Month, ul1.EffectiveDate) as varchar(3)) + Cast(100 + DATEPART(Day, ul1.EffectiveDate) as varchar(3)) end, ")

            ' displays plus sign
            sb.Append("Active='1', ")
            'Sql &= "Active = case "
            'Sql &= " when (SELECT Count (*) FROM UserLicenses ul3 WHERE ul3.UserID = '" & cSess.UserID & "' and ul3.state = ul1.State) = 1 then '1' "
            'Sql &= " when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSess.UserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then '1' "
            'Sql &= " else '0' end, "

            'Sql &= "Catgy = case "
            'Sql &= " when ul1.EffectiveDate = '1/1/1950' then 'Application' "
            'Sql &= " when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSess.UserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then 'Effective' "
            ''Sql &= " else 'Superseded' end, " '
            'Sql &= " end, "

            sb.Append("Catgy = case ")
            sb.Append(" when ul1.ExpirationDate < getDate() then 'Expired' ")
            sb.Append(" when ul1.EffectiveDate = '1/1/1950' then 'Application' ")
            sb.Append(" else 'Effective' ")
            sb.Append(" end, ")

            sb.Append("LongTermCareCert = dbo.ufn_getLTCCert(getDate(), u.UserID, ul1.State), ")
            sb.Append(" NumAppointments = case ")
            sb.Append(" when (SELECT Count (*) FROM UserLicenses ul3 WHERE ul3.UserID = '" & cSess.UserID & "' and ul3.state = ul1.State) = 1 then  Cast((Select Count (*) FROM UserAppointments WHERE UserID =ul1.UserID AND State=ul1.State) as varchar(10) )")
            sb.Append(" when (Select Max(ul2.EffectiveDate) From UserLicenses ul2 WHERE ul2.UserID ='" & cSess.UserID & "' and ul2.state =  ul1.State) = ul1.EffectiveDate then Cast((Select Count (*) FROM UserAppointments WHERE UserID =ul1.UserID AND State=ul1.State) as varchar(10)) ")
            sb.Append(" else '' ")
            sb.Append(" end, ")
            sb.Append(" ul1.Notes ")
            sb.Append(" FROM Users u ")
            sb.Append(" INNER JOIN UserLicenses ul1 on u.UserID = ul1.UserID ")
            sb.Append(" INNER JOIN Codes_State cs on ul1.State = cs.StateCode ")
            sb.Append(" WHERE ul1.UserID ='" & cSess.UserID & "'")
            Sql = sb.ToString

            'Sql = "SELECT cc.ClientID, cc.ClientID hdClientID, cc.ClientID ClientIDStr, HasPermissionStr = case when (SELECT Count (*) FROM UserPermissions up WHERE UserID = '" & UserID & "' and cc.ClientID = up.ClientID) = 0 then 'No' else '<b>Yes</b>'  End,  HasPermission = case when (SELECT Count (*) FROM UserPermissions up WHERE UserID = '" & UserID & "' and cc.ClientID = up.ClientID) = 0 then '0' else '1'  End FROM Codes_ClientID cc"

            '   DG.GenerateSQL(Sql, ShowFilter, Nothing, OrderByType, Request, True)
            DG.GenerateSQL(Sql, ShowFilter, Nothing, OrderByType, Request, cSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"), True)


            'dt = cCommon.GetDT(Sql)
            dt = cCommon.GetDT(Sql, cenviro.DBHost, "UserManagement", True)

            ' ___ Write the datagrid to the page
            litDG.Text = DG.GetText(dt, Nothing)

            ' ___ Set the FilterOnOffState
            If DG.FilterOperationMode = DG.FilterOperationModeEnum.FilterSwitchable AndAlso ShowFilter Then
                cSess.FilterOnOffState = "on"
            Else
                cSess.FilterOnOffState = "off"
            End If

            ' ___ Set the last field sorted and sort direction in the viewstate strings
            '  Viewstate("SortData") = DG.GetViewStateString()

            cSess.SortReference = DG.GetSortReference

            litKeyValues.Text = "<input type='hidden' id='hdState' name='hdState' value=""" & cSess.State & """><input type='hidden' id='hdSubTableState' name='hdSubTableState' value=""" & cSess.SubTableState & """><input type='hidden' id='hdEffectiveDate' name='hdEffectiveDate' value=""" & cSess.EffectiveDate & """><input type='hidden' id='hdExpirationDate' name='hdExpirationDate' value=""" & cSess.ExpirationDate & """><input type='hidden' id='hdLicenseNumber' name='hdLicenseNumber' value=""" & cSess.LicenseNumber & """><input type='hidden' id='hdCarrierID' name='hdCarrierID' value=""" & cSess.CarrierID & """><input type='hidden' id='hdSubTableInd' name='hdSubTableInd' value=""" & cSess.SubTableInd & """>"

            '' ___ Write the filter hiddens to the page
            'If DG.FilterOperationMode = DG.FilterOperationModeEnum.FilterSwitchable Then
            '    If ShowFilter Then
            '        litFilterHiddens.Text = "<input type='hidden' name='hdFilterOnResponse' value='on'><input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0'>"
            '    Else
            '        litFilterHiddens.Text = "<input type='hidden' name='hdFilterOnResponse' value='off'><input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0'>"
            '    End If
            'End If

        Catch ex As Exception
            Throw New Exception("Error #407: LicenseWorklist DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Function SqlIndFalse(ByVal IndName) As String
        Return "(" & IndName & " is null or " & IndName & " = 0)"
    End Function

    Private Function SqlIndTrue(ByVal IndName As String, ByVal FromDate As String, ByVal ToDate As String) As String
        Return " (" & IndName & " = 1 and " & FromDate & " is not null and " & FromDate & " <= getDate() and " & ToDate & " is not null and DateAdd(Day, 1,  " & ToDate & ") >= getDate()) "
    End Function

    Private Function SqlLicTest() As String
        'Return "(ul1.EffectiveDate is not null and ul1.EffectiveDate <= getDate() and (ul1.ExpirationDate is null or ul1.ExpirationDate >= getDate()))"
        Return "(ul1.EffectiveDate is not null and ul1.EffectiveDate <= getDate() and (ul1.ExpirationDate is null or DateAdd(Day, 1, ul1.ExpirationDate) >= getDate()))"
    End Function

End Class