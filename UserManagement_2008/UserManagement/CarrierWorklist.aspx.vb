Imports System.Data
Imports System.Data.SqlClient

Partial Class CarrierWorklist
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Private cSess As CarrierWorklistSession
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
            RightsRqd.SetValue(RightsClass.CarrierView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the page session object
            cSess = Session("CarrierWorklistSession")

            ' ___ Get the page mode
            PageMode = cCommon.GetPageMode(Page, cSess)

            ' ___ Load the page session variables
            Select Case PageMode
                Case PageMode.Initial
                    ' No action
                Case PageMode.ReturnFromChild, PageMode.CalledByOther
                    cSess.CarrierID = cSess.CarrierIDSelectedFilterValue
                Case PageMode.Postback
                    cSess.CarrierID = Replace(Request.Form("hdCarrierID"), "~", "'")
                    cSess.CarrierIDSelectedFilterValue = Replace(Request.Form("txtCarrierID"), "~", "'")
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
                            litMsg.Text = "<script type=""text/javascript"">alert('" & Results.Msg & "')</script>"

                        Case "ExistingRecord"
                            Response.Redirect("CarrierMaintain.aspx?CallType=Existing")

                        Case "NewRecord"
                            Response.Redirect("CarrierMaintain.aspx?CallType=New")
                    End Select

                Case PageMode.ReturnFromChild, PageMode.CalledByOther
                    DisplayPage(PageMode, DG, DG.OrderByType.ReturnToPage)
                    If cSess.PageReturnOnLoadMessage <> Nothing Then
                        litMsg.Text = "<script type=""text/javascript"">alert('" & cSess.PageReturnOnLoadMessage & "')</script>"
                        cSess.PageReturnOnLoadMessage = Nothing
                    End If
            End Select

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            'litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'/><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'/>"
            litEnviro.Text = cCommon.GetlitEnviroText(cEnviro.LoggedInUserID, cEnviro.DBHost)

        Catch ex As Exception
            'Throw New Exception("Error #202: CarrierWorklist Page_Load. " & ex.Message)
            Dim ErrorObj As New ErrorObj(ex, "Error #202: CarrierWorklist Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function DeleteRecord() As Results
        Dim MyResults As New Results
        Dim Sql As String
        Dim dt As DataTable

        Try

            ' ___ Are there any users associated with this carrier?
            ' dt = cCommon.GetDT("SELECT Count (*) FROM UserAppointments WHERE CarrierID = '" & Request.Form("hdCarrierID") & "'")
            dt = cCommon.GetDT("SELECT Count (*) FROM UserAppointments WHERE CarrierID = " & cCommon.StrOutHandler(cSess.CarrierID, False, StringTreatEnum.SideQts_SecApost))

            If dt.Rows(0)(0) = 0 Then

                ' ___ No current users. Delete the company.
                'Sql = "DELETE FROM Codes_CarrierID WHERE CarrierID = '" & Request.Form("hdCarrierID") & "'"
                'Sql = "DELETE FROM Codes_CarrierID WHERE CarrierID = " & cCommon.StrOutHandler(cSess.CarrierID, False, StringTreatEnum.SideQts_SecApost)
                Sql = "UPDATE  Codes_CarrierID SET LogicalDelete = 1 WHERE   CarrierID = " & cCommon.StrOutHandler(cSess.CarrierID, False, StringTreatEnum.SideQts_SecApost)
                cCommon.ExecuteNonQuery(Sql)
                MyResults.Success = True
                MyResults.Msg = "Record deleted."
            Else

                ' ___ Error message
                MyResults.Success = False
                MyResults.Msg = "Cannot delete carrier while there are users associated with this carrier."
            End If

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #203: CarrierWorklist DeleteRecord. " & ex.Message)
        End Try
    End Function

    Private Function DefineDataGrid() As DG
        Try

            Dim DG As New DG("CarrierID", cCommon, cRights, True, "EmbeddedTableDef", "CarrierID")
            DG.AddNewButton(RightsClass.CarrierEdit)
            DG.AddDataBoundColumn("CarrierID", "CarrierID", "Carrier", "CarrierID", True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("Description", "Description", "Description", Nothing, True, Nothing, Nothing, "align='left'")

            ' ___ Build the filter
            Dim Filter As DG.Filter
            Filter = DG.AttachFilter(DG.FilterOperationModeEnum.FilterSwitchable, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialShow)
            Filter.AddTextbox("CarrierID", "CarrierID", 50)

            Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, False, Nothing, True)
            TemplateCol.AddDefaultTemplateItem("View", "ExistingRecord", "StandardView", "User record", RightsClass.CarrierView, Nothing)
            TemplateCol.AddDefaultTemplateItem("Delete", "Delete", "StandardDelete", "Delete", RightsClass.CarrierView, Nothing)
            DG.AttachTemplateCol(TemplateCol)

            Return DG

        Catch ex As Exception
            Throw New Exception("Error #204: CarrierWorklist DefineDataGrid. " & ex.Message)
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

            ' ___ CarrierID
            If PageMode <> PageMode.Initial Then
                Filter.Coll("CarrierID").SetFilterValue(cSess.CarrierIDSelectedFilterValue)
            End If

            ' ___ Handle the sort
            If cSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            Sql = "SELECT * FROM Codes_CarrierID"
            SecurityWhereClause = "LogicalDelete = 0"

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
            Throw New Exception("Error #205: CarrierWorklist DisplayPage. " & ex.Message)
        End Try
    End Sub

End Class
