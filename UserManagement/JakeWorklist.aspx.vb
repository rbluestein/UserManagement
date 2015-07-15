Imports System.Data.SqlClient

Public Class JakeWorklist
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected WithEvents litDG As System.Web.UI.WebControls.Literal
    Protected WithEvents dgOrgList As System.Web.UI.WebControls.DataGrid
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Private AppSession As New AppSession
    Private Rights As RightsClass
    Private Common As New Common()
    Private cDefaultSortField As String = "TicketNum"
    Protected WithEvents litFilterHiddens As System.Web.UI.WebControls.Literal
    Private cLoggedInUserID As String
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

#Region " Enums "
    Private Enum QAStatus
        Unresolved = 1
        Resolved = 2
    End Enum
#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Results As Results
        Dim Action As String
        Dim DG As DG

        Try

            ' ___ Restore  session
            cLoggedInUserID = AppSession.RestoreSession(Page)

            ' ___ Right check
            Rights = New RightsClass(cLoggedInUserID, Page)
            Dim RightsRqd(0) As String
            RightsRqd.SetValue(Rights.JakeEdit, 0)
            Rights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = Common.GetCurRightsHidden(Rights.RightsColl)


            If Page.IsPostBack AndAlso Request.Form("__EVENTTARGET") = "" Then
                DG = DefineDataGrid()
                Action = Request.Form("hdAction")

                Select Case Action
                    Case "Sort"
                        DisplayPage(DG, DG.OrderByType.Field, Request.Form("hdSortField"))

                    Case "ApplyFilter"
                        DisplayPage(DG, DG.OrderByType.Recurring)

                    Case "Delete"
                        'Results = DeleteRecord()
                        'DisplayPage(DG, DG.OrderByType.Recurring)
                        'litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"

                    Case "View"
                        Response.Redirect("JakeMaintain.aspx?JakeId=" & Request.Form("hdJakeId"))

                    Case "NewRecord"
                        Response.Redirect("JakeMaintain.aspx?JakeId=")
                End Select

            Else

                DG = DefineDataGrid()
                DisplayPage(DG, DG.OrderByType.Initial)

            End If

        Catch ex As Exception
            Throw New Exception("Error", ex)
        End Try
    End Sub

#Region " Process selection "
    Sub ProcessAction(ByVal Action As String)
    End Sub
#End Region

#Region " Datagrid "
    Private Function DefineDataGrid() As DG
        Dim DG As New DG("JakeId", Rights, True, "EmbeddedTableDef", "TicketNum")
        DG.AddNewButton(Rights.JakeEdit)
        DG.AddDataBoundColumn("JakeID", "JakeID", "JakeID", "JakeID", False, Nothing, Nothing, "align='left'")
        DG.AddDataBoundColumn("TicketNum", "TicketNum", "Ticket<br> Num", "TicketNum", True, Nothing, Nothing, "align='left' width='60px'")
        DG.AddDataBoundColumn("IssueSummary", "IssueSummary", "Issue", "IssueSummary", True, Nothing, Nothing, "align='left'")

        DG.AddFreeFormColumn("Spacer", Nothing, Nothing, Nothing, True, "width='10px'")


        DG.AddDataBoundColumn("Originator", "Originator", "Originator", "Originator", True, Nothing, Nothing, "align='left'")
        DG.AddDataBoundColumn("TechStatus", "TechStatusStr", "Tech<br>Status", "TechStatus", True, Nothing, Nothing, "align='left'")
        DG.AddDataBoundColumn("QAStatus", "QAStatusStr", "QA<br>Status", "QAStatus", True, Nothing, Nothing, "align='left'")
        DG.AddDateColumn("OpenDt", "OpenDt", "Open<br>Date", "OpenDt", True, "MM/dd/yyyy", Nothing, "align='left' width='60px'")
        DG.AddDateColumn("CloseDt", "CloseDt", "Close<br>Date", "CloseDt", True, "MM/dd/yyyy", Nothing, "align='left'")

        Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, False, Nothing, True)
        TemplateCol.AddDefaultTemplateItem("View", "View", "StandardView", "Edit ticket", Rights.JakeEdit, Nothing, Nothing, Nothing)
        DG.AttachTemplateCol(TemplateCol)

        ' ___ Build the filter
        Dim Filter As DG.Filter
        Filter = DG.AttachFilter(DG.FilterOperationMode.FilterSwitchable, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialShow)
        Filter.AddDropdown("TechStatus", "TechStatus")
        Filter.AddDropdown("QAStatus", "QAStatus")

        Return DG
    End Function

    Private Sub DisplayPage(ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Dim dt As New Data.DataTable
        Dim Sql As String
        Dim NewSortDirection As String
        Dim NewOrderByField As String
        Dim ShowFilter As Boolean
        Dim sb As New System.Text.StringBuilder

        ' ___ Get a filter reference
        Dim Filter As DG.Filter
        Filter = DG.GetFilter

        ' ___ QAStatusStr
        Filter("QAStatus").AddDropdownItem("", "", True)
        Filter("QAStatus").AddDropdownItem("1", "Unresolved", False)
        Filter("QAStatus").AddDropdownItem("2", "Resolved", False)

        ' ___ TechStatusStr
        Filter("TechStatus").AddDropdownItem("", "", True)
        Filter("TechStatus").AddDropdownItem("1", "Not started")
        Filter("TechStatus").AddDropdownItem("2", "In progress")
        Filter("TechStatus").AddDropdownItem("3", "Unable to duplicate")
        Filter("TechStatus").AddDropdownItem("4", "Completed")

        ' ___ Handle the sort
        If OrderByType <> DG.OrderByType.Initial Then
            DG.UpdateViewstate(Viewstate("SortData"))
        End If
        DG.SetSortElements(OrderByField, OrderByType, NewOrderByField, NewSortDirection)

        ' ___ Load the parameters and execute the query
        sb.Append("Select *, ")
        sb.Append(" TechStatusStr = case  when TechStatus = 1 then 'Not started' when TechStatus = 2 then 'In progress' when TechStatus = 3 then 'Unable to duplicate' when TechStatus = 4 then 'Completed' else  ''  end,  ")
        sb.Append(" QAStatusStr = case  when QAStatus = 1 then 'Unresolved' when QAStatus = 2 then 'Resolved' else  ''  end,  ")
        sb.Append("OpenDt,  CloseDt ")
        sb.Append("FROM Jake ")



        Sql = sb.ToString
        DG.GenerateSQL(Sql, ShowFilter, Nothing, OrderByType, Request)
        dt = Common.GetDT(Sql)

        'Dim CmdAsst As New CmdAsst(CommandType.StoredProcedure, "JakesRead")
        'Dim QueryPack As CmdAsst.QueryPack
        ''  CmdAsst.AddInt("Filter", "QAStatusStr")
        'CmdAsst.AddVarChar("OrderByField", NewOrderByField, 20)
        'If NewSortDirection = "A" Then
        '    CmdAsst.AddBit("OrderByDir", 0)
        'Else
        '    CmdAsst.AddBit("OrderByDir", 1)
        'End If
        'QueryPack = CmdAsst.Execute
        'dt = QueryPack.dt

        ' ___ Write the datagrid to the page
        litDG.Text = DG.GetText(dt)

        ' ___ Set the last field sorted and sort direction in the viewstate strings
        Viewstate("SortData") = DG.GetViewStateString()

        ' ___ Write the filter hiddens to the page
        If DG.FilterOperationMode = DG.FilterOperationModeEnum.FilterSwitchable Then
            If ShowFilter Then
                litFilterHiddens.Text = "<input type='hidden' name='hdFilterOnResponse' value='on'><input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0'>"
            Else
                litFilterHiddens.Text = "<input type='hidden' name='hdFilterOnResponse' value='off'><input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0'>"
            End If
        End If
    End Sub


#End Region

End Class
