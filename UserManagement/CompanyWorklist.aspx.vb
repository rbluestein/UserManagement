Imports System.Data.SqlClient

Public Class CompanyWorklist
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected PageCaption As HtmlGenericControl
    Protected WithEvents litDG As System.Web.UI.WebControls.Literal
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Protected WithEvents litHiddens As System.Web.UI.WebControls.Literal
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
    Private cSess As CompanyWorklistSession
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
            RightsRqd.SetValue(cRights.CompanyView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the page session object
            cSess = Session("CompanyWorklistSession")

            ' ___ Get the page mode
            PageMode = cCommon.GetPageMode(Page, cSess)

            ' ___ Load the page session variables
            Select Case PageMode
                Case PageMode.Initial
                    ' No action
                Case PageMode.ReturnFromChild, PageMode.CalledByOther
                    ' No action
                Case PageMode.Postback
                    cSess.CompanyID = Replace(Request.Form("hdCompanyID"), "~", "'")
                    cSess.CompanyIDSelectedFilterValue = Replace(Request.Form("txtCompanyID"), "~", "'")
                    cSess.PrimaryContactNameSelectedFilterValue = Replace(Request.Form("txtPrimaryContactName"), "~", "'")
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
                            Response.Redirect("CompanyMaintain.aspx?CallType=Existing")

                        Case "NewRecord"
                            Response.Redirect("CompanyMaintain.aspx?CallType=New")
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
            Throw New Exception("Error #302: CompanyWorklist Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function DeleteRecord() As Results
        Dim MyResults As New Results
        Dim Sql As String
        Dim dt As DataTable

        Try

            ' ___ Are there any users associated with this company?
            ' dt = cCommon.GetDT("SELECT Count (*) FROM Users WHERE CompanyID = '" & Request.Form("hdCompanyID") & "'")
            dt = cCommon.GetDT("SELECT Count (*) FROM Users WHERE CompanyID = " & cCommon.StrOutHandler(cSess.CompanyID, False, StringTreatEnum.SideQts_SecApost))
            If dt.Rows(0)(0) = 0 Then

                ' ___ No current users. Delete the company.
                Sql = "DELETE FROM Codes_CompanyID WHERE CompanyID = " & cCommon.StrOutHandler(cSess.CompanyID, False, StringTreatEnum.SideQts_SecApost)
                cCommon.ExecuteNonQuery(Sql)
                MyResults.Success = True
                MyResults.Msg = "Record deleted."
            Else

                ' ___ Error message
                MyResults.Success = False
                MyResults.Msg = "Cannot delete company while there are users associated with this company."
            End If

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #303: CompanyWorklist DeleteRecord. " & ex.Message)
        End Try
    End Function

    Private Function DefineDataGrid() As DG
        Try

            Dim DG As New DG("CompanyID", cCommon, cRights, True, "EmbeddedTableDef", "CompanyID")
            DG.AddNewButton(cRights.CompanyEdit)
            DG.AddDataBoundColumn("CompanyID", "CompanyID", "Company", "CompanyID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("PrimaryContactName", "PrimaryContactName", "Primary Contact", "PrimaryContactName", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("PrimaryContactPhone", "PrimaryContactPhone", "Phone", Nothing, True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("PrimaryContactEmal", "PrimaryContactEmail", "Email", Nothing, True, Nothing, Nothing, "align='left'")

            ' ___ Build the filter
            Dim Filter As DG.Filter
            Filter = DG.AttachFilter(DG.FilterOperationMode.FilterSwitchable, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialShow)
            Filter.AddTextbox("CompanyID", "CompanyID", 50)
            Filter.AddTextbox("PrimaryContactName", "PrimaryContactName", 100, "PrimaryContactName", Nothing)

            Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, False, Nothing, True)
            TemplateCol.AddDefaultTemplateItem("View", "ExistingRecord", "StandardView", "User record", cRights.CompanyView, Nothing)
            TemplateCol.AddDefaultTemplateItem("Delete", "Delete", "StandardDelete", "Delete", cRights.CompanyEdit, Nothing)
            DG.AttachTemplateCol(TemplateCol)

            Return DG

        Catch ex As Exception
            Throw New Exception("Error #304: CompanyWorklist DefineDataGrid. " & ex.Message)
        End Try
    End Function

    Private Sub DisplayPage(ByVal PageMode As PageMode, ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Dim dt As DataTable
        Dim Sql As String
        Dim ShowFilter As Boolean
        Dim SecurityWhereClause As String
        Dim sbHiddens As New System.Text.StringBuilder

        Try

            ' ___ Get a filter reference
            Dim Filter As DG.Filter
            Filter = DG.GetFilter

            ' ___ CompanyID
            If PageMode <> PageMode.Initial Then
                Filter.Coll("CompanyID").SetFilterValue(cSess.CompanyIDSelectedFilterValue)
            End If

            ' ___ PrimaryContactName
            If PageMode <> PageMode.Initial Then
                Filter.Coll("PrimaryContactName").SetFilterValue(cSess.PrimaryContactNameSelectedFilterValue)
            End If

            ' ___ Handle the sort
            If cSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            Sql = "SELECT * FROM Codes_CompanyID"

            DG.GenerateSQL(Sql, ShowFilter, SecurityWhereClause, OrderByType, Request, cSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"))
            dt = cCommon.GetDT(Sql, cEnviro.DBHost, "UserManagement", True)

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
            Throw New Exception("Error #305: CompanyWorklist DisplayPage. " & ex.Message)
        End Try
    End Sub
End Class
