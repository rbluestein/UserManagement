Imports System.Data.SqlClient

Public Class ClientWorklist
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
    Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
    Private cSess As ClientWorklistSession
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

            ' ___ Instantiate objects
            cEnviro = Session("Enviro")
            cCommon = New Common

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right check
            cRights = New RightsClass(Page)
            Dim RightsRqd(0) As String
            RightsRqd.SetValue(cRights.ClientView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the page session object
            cSess = Session("ClientWorklistSession")

            ' ___ Get the page mode
            PageMode = cCommon.GetPageMode(Page, cSess)

            ' ___ Load the page session variables
            Select Case PageMode
                Case PageMode.Initial
                    ' No action
                Case PageMode.ReturnFromChild, PageMode.CalledByOther
                    cSess.ClientID = cSess.ClientIDSelectedFilterValue
                Case PageMode.Postback
                    cSess.ClientID = Replace(Request.Form("hdClientID"), "~", "'")
                    cSess.ClientIDSelectedFilterValue = Replace(Request.Form("txtClientID"), "~", "'")
            End Select

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
                            Response.Redirect("ClientMaintain.aspx?CallType=Existing")

                        Case "NewRecord"
                            Response.Redirect("ClientMaintain.aspx?CallType=New")
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
            Throw New Exception("Error #1402: ClientWorklist Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function DeleteRecord() As Results
        Dim MyResults As New Results
        Dim Sql As String
        Dim dt As DataTable

        Try

            ' ___ Are there any users associated with this client?
            dt = cCommon.GetDT("SELECT Count (*) FROM UserPermissions up  INNER JOIN UserManagement..Users u on up.UserID = u.UserID WHERE up.ClientID = " & cCommon.StrOutHandler(cSess.ClientID, False, StringTreatEnum.SideQts_SecApost) & " AND u.StatusCode='ACTIVE'")

            If dt.Rows(0)(0) = 0 Then

                ' ___ No current users. Delete the company.
                Sql = "UPDATE  Codes_ClientID SET LogicalDelete = 1 WHERE ClientID = " & cCommon.StrOutHandler(cSess.ClientID, False, StringTreatEnum.SideQts_SecApost)
                cCommon.ExecuteNonQuery(Sql)
                MyResults.Success = True
                MyResults.Msg = "Record deleted."
            Else

                ' ___ Error message
                MyResults.Success = False
                MyResults.Msg = "Cannot delete client while there are users associated with this client."
            End If

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #1403: ClientWorklist DeleteRecord. " & ex.Message)
        End Try
    End Function

    Private Function DefineDataGrid() As DG
        Try

            Dim DG As New DG("ClientID", cCommon, cRights, True, "EmbeddedTableDef", "ClientID")
            DG.AddNewButton(cRights.ClientEdit)
            DG.AddDataBoundColumn("ClientID", "ClientID", "Client", "ClientID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("Name", "Name", "Name", Nothing, True, Nothing, Nothing, "align='left'")


            ' ___ Build the filter
            Dim Filter As DG.Filter
            Filter = DG.AttachFilter(DG.FilterOperationMode.FilterSwitchable, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialShow)
            Filter.AddTextbox("ClientID", "ClientID", 50)

            Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, False, Nothing, True)
            TemplateCol.AddDefaultTemplateItem("View", "ExistingRecord", "StandardView", "View", cRights.ClientView, Nothing)
            TemplateCol.AddDefaultTemplateItem("Delete", "Delete", "StandardDelete", "Delete", cRights.ClientView, Nothing)
            DG.AttachTemplateCol(TemplateCol)

            Return DG

        Catch ex As Exception
            Throw New Exception("Error #1404: ClientWorklist DefineDataGrid. " & ex.Message)
        End Try
    End Function

    Private Sub DisplayPage(ByVal PageMode As PageMode, ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Dim dt As DataTable
        Dim Sql As String
        Dim ShowFilter As Boolean
        Dim SecurityWhereClause As String

        Try

            ' ___ Get a filter reference
            Dim Filter As DG.Filter
            Filter = DG.GetFilter

            ' ___ ClientID
            If PageMode <> PageMode.Initial Then
                Filter.Coll("ClientID").SetFilterValue(cSess.ClientIDSelectedFilterValue)
            End If

            ' ___ Handle the sort
            If cSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            Sql = "SELECT * FROM Codes_ClientID"
            SecurityWhereClause = "LogicalDelete = 0"

            DG.GenerateSQL(Sql, ShowFilter, SecurityWhereClause, OrderByType, Request, cSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"))
            dt = cCommon.GetDT(Sql, cenviro.DBHost, "UserManagement", True)

            ' ___ Write the datagrid to the page
            litDG.Text = DG.GetText(dt, Nothing)

            ' ___ Set the FilterOnOffState
            If DG.FilterOperationMode = DG.FilterOperationModeEnum.FilterSwitchable AndAlso ShowFilter Then
                cSess.FilterOnOffState = "on"
            Else
                cSess.FilterOnOffState = "off"
            End If

            cSess.SortReference = DG.GetSortReference

        Catch ex As Exception
            Throw New Exception("Error #1405: ClientWorklist DisplayPage. " & ex.Message)
        End Try
    End Sub

End Class