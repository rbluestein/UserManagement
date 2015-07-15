Imports System.Data.SqlClient

Public Class Permissions
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected PageCaption As HtmlGenericControl
    Protected WithEvents litDG As System.Web.UI.WebControls.Literal
    Protected WithEvents litResponseAction As System.Web.UI.WebControls.Literal
    Protected WithEvents litHeading As System.Web.UI.WebControls.Literal
    Protected WithEvents lblRole As System.Web.UI.WebControls.Label
    Protected WithEvents lblCompany As System.Web.UI.WebControls.Label
    Protected WithEvents lblFullName As System.Web.UI.WebControls.Label
    Protected WithEvents litUpdate As System.Web.UI.WebControls.Literal
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents litFilterHiddens As System.Web.UI.WebControls.Literal
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
    Private cSess As UserWorklistSession
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
        Dim Results As Results
        Dim RequestAction As RequestAction
        Dim ResponseAction As ResponseAction
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
            RightsRqd.SetValue(cRights.PermissionsView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the session object
            cSess = Session("UserWorklistSession")

            ' ___ Define the datagrid
            DG = DefineDataGrid()

            ' ___ Get RequestAction
            RequestAction = cCommon.GetRequestAction(Page)

            ' ___ Execute the RequestAction
            Results = ExecuteRequestAction(RequestAction)
            ResponseAction = Results.ResponseAction

            ' ___ Execute the ResponseAction
            If ResponseAction = ResponseAction.ReturnToCallingPage Then
                cSess.PageReturnOnLoadMessage = Results.Msg
                Response.Redirect("Index.aspx?CalledBy=Child")
            Else
                'DisplayPage(ResponseAction)
                DisplayPage(DG, DG.OrderByType.Recurring)
                If Not Results.Msg = Nothing Then
                    litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                End If
                litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"
            End If

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'>"


            'If Page.IsPostBack AndAlso Request.Form("__EVENTTARGET") = "" Then
            '    Select Case Request.Form("hdAction")
            '        Case "return"
            '            Response.Redirect("Index.aspx?CalledBy=Child")
            '        Case "Sort"
            '            DisplayPage(DG, DG.OrderByType.Field, Request.Form("hdSortField"))
            '        Case "ApplyFilter"
            '            DisplayPage(DG, DG.OrderByType.Recurring)
            '        Case "Update"
            '            PerformSave()
            '            DisplayPage(DG, DG.OrderByType.Recurring, Request.Form("hdAction"))
            '        Case "ExistingRecord"
            '            Response.Redirect("UserMaintain.aspx?UserID=" & Request.Form("hdUserID"))
            '        Case "NewRecord"
            '            Response.Redirect("UserMaintain.aspx?UserID=")
            '    End Select
            'Else
            '    DisplayPage(DG, DG.OrderByType.Initial)
            'End If

        Catch ex As Exception
            Throw New Exception("Error #902: Permissions Page_Load. " & ex.Message)
        End Try
    End Sub


    Private Function DefineDataGrid() As DG
        Try

            Dim DG As New DG("ClientID", cCommon, cRights, True, "EmbeddedTableDef", "UserID")
            DG.AddHiddenColumn("hdClientID", "HasPermission")
            DG.AddDataBoundColumn("ClientIDStr", "ClientID", "Client", Nothing, True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("HasPermissionStr", "HasPermissionStr", "Has Permission", Nothing, True, Nothing, Nothing, "align='left'")
            DG.AddCheckboxToggleColumn("chkPermission", "ClientID", "Add/Remove Permission", Nothing, True, cRights.PermissionsEdit, "HasPermission", "Remove permission", "Add permission", Nothing, "align='left'")
            Return DG

        Catch ex As Exception
            Throw New Exception("Error #903: Permissions DefineDataGrid. " & ex.Message)
        End Try
    End Function


    Private Function ExecuteRequestAction(ByVal RequestAction As RequestAction) As Results
        Dim ValidationResults As New Results
        Dim SaveResults As New Results
        Dim MyResults As New Results

        Try

            Select Case RequestAction
                Case RequestAction.ReturnToParent
                    MyResults.ResponseAction = ResponseAction.ReturnToCallingPage
                    MyResults.Success = True

                Case RequestAction.CreateNew
                    MyResults.ResponseAction = ResponseAction.DisplayBlank

                Case RequestAction.SaveNew
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then
                            MyResults.ResponseAction = ResponseAction.DisplayExisting
                            MyResults.Msg = SaveResults.Msg
                        Else
                            MyResults.ResponseAction = ResponseAction.DisplayUserInputNew
                            MyResults.Msg = SaveResults.Msg
                        End If
                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = ResponseAction.DisplayUserInputNew
                        MyResults.ObtainConfirm = True
                        MyResults.Msg = ValidationResults.Msg
                    Else
                        MyResults.ResponseAction = ResponseAction.DisplayUserInputNew
                        MyResults.Msg = ValidationResults.Msg
                    End If

                Case RequestAction.LoadExisting
                    MyResults.ResponseAction = ResponseAction.DisplayExisting

                Case RequestAction.SaveExisting
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then
                            MyResults.ResponseAction = ResponseAction.DisplayExisting
                            MyResults.Msg = SaveResults.Msg
                        Else
                            MyResults.ResponseAction = ResponseAction.DisplayUserInputExisting
                            MyResults.Msg = SaveResults.Msg
                        End If
                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = ResponseAction.DisplayUserInputExisting
                        MyResults.ObtainConfirm = True
                        MyResults.Msg = ValidationResults.Msg
                    Else
                        'MyResults.ResponseAction = ResponseAction.DisplayExisting
                        MyResults.ResponseAction = ResponseAction.DisplayUserInputExisting
                        MyResults.Msg = ValidationResults.Msg
                    End If

                Case RequestAction.NoSaveNew
                    MyResults.ResponseAction = ResponseAction.DisplayUserInputNew

                Case RequestAction.NoSaveExisting
                    MyResults.ResponseAction = ResponseAction.DisplayUserInputExisting

                Case RequestAction.Other
                    MyResults.ResponseAction = cCommon.GetResponseActionFromRequestActionOther(Page)
            End Select

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #904: Permissions ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function IsDataValid(ByVal RequestAction As RequestAction) As Results
        Dim MyResults As New Results
        MyResults.Success = True
        Return MyResults
    End Function


    Private Function PerformSave(ByVal RequestAction As RequestAction) As Results
        Dim i As Integer
        Dim CurRight As Integer
        Dim ToggleRight As Integer
        'Dim RightDefId As String
        Dim AddRight As Integer
        Dim DeleteRight As Integer
        Dim EditRecord As Integer
        Dim PerformWrite As Boolean
        Dim SqlCmd As System.Data.SqlClient.SqlCommand
        Dim Arr As Object
        Dim Coll As New Collection
        Dim MyResults As New Results

        'For i = 0 To Count - 1
        '    Arr(i, 0) = Request.Form.AllKeys(i)
        '    Arr(i, 1) = Request.QueryString(Arr(i, 0))
        'Next

        Try

            For i = 0 To Request.Form.AllKeys.Length - 1
                If Request.Form.AllKeys(i).Length > 10 Then
                    If Request.Form.AllKeys(i).Substring(0, 10) = "hdClientID" Then
                        Coll.Add(Request.Form.AllKeys(i))
                    End If
                End If
            Next

            Dim SqlConnection1 As New SqlConnection(cEnviro.GetConnectionString)
            SqlConnection1.Open()

            For i = 1 To Coll.Count

                ' ___ Reset values
                AddRight = 0
                DeleteRight = 0
                EditRecord = 0
                PerformWrite = False

                ' ___ Toggle the right?
                Arr = Split(Coll(i), "|")
                If Request.Form("chkPermission|" & Arr(1)) = Nothing Then
                    ToggleRight = 0
                Else
                    ToggleRight = 1
                End If

                ' ___ Get the current right
                CurRight = Request.Form("hdClientID|" & Arr(1))

                ' ___ Determine add/delete
                If ToggleRight = 1 Then
                    If CurRight = 1 Then
                        DeleteRight = 1
                        PerformWrite = True
                    Else
                        AddRight = 1
                        PerformWrite = True
                    End If
                End If

                If PerformWrite Then
                    SqlCmd = New System.data.SqlClient.SqlCommand
                    SqlCmd.CommandType = System.Data.CommandType.Text
                    SqlCmd.Connection = SqlConnection1
                    If AddRight Then
                        SqlCmd.CommandText = "INSERT INTO UserPermissions (UserID, ClientID, AuthorizedBy) Values ('" & cSess.UserID & "', '" & Arr(1) & "', '" & cEnviro.LoggedInUserID & "')"
                    ElseIf DeleteRight Then
                        SqlCmd.CommandText = "DELETE UserPermissions WHERE UserID = '" & cSess.UserID & "' AND ClientID = '" & Arr(1) & "'"
                    End If
                    SqlCmd.ExecuteNonQuery()
                End If
            Next
            MyResults.Success = True
            MyResults.Msg = "Update complete."

            SqlConnection1.Close()
            If Not SqlCmd Is Nothing Then
                SqlCmd.Dispose()
            End If

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Msg = "Unable to save record."
        End Try

        Return MyResults
    End Function

    Private Sub DisplayPage(ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Dim dt As DataTable
        Dim Sql As String
        Dim NewSortDirection As String
        Dim NewOrderByField As String

        Try

            ' ___ Heading/UserID
            If cRights.HasThisRight(cRights.PermissionsEdit) Then
                litHeading.Text = "Edit Permissions"
            Else
                litHeading.Text = "View Permissions"
            End If

            Sql = "SELECT LastName + ', ' +  FirstName + ' ' + MI FullName, Role, CompanyID, StatusCode FROM Users WHERE UserID ='" & cSess.UserID & "'"
            dt = cCommon.GetDT(Sql)
            lblFullName.Text = dt.Rows(0)("FullName")
            lblRole.Text = dt.Rows(0)("Role")
            lblCompany.Text = dt.Rows(0)("CompanyID")

            ' ___ Update button
            If cRights.HasThisRight(cRights.PermissionsEdit) Then
                litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"
            End If

            ' ___ Handle the sort
            If cSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            ' ___ Load the parameters and execute the query
            Sql = "SELECT cc.ClientID, cc.ClientID hdClientID, cc.ClientID ClientIDStr, HasPermissionStr = case when (SELECT Count (*) FROM UserPermissions up WHERE UserID = '" & cSess.UserID & "' and cc.ClientID = up.ClientID) = 0 then 'No' else '<b>Yes</b>'  End,  HasPermission = case when (SELECT Count (*) FROM UserPermissions up WHERE UserID = '" & cSess.UserID & "' and cc.ClientID = up.ClientID) = 0 then '0' else '1'  End FROM Codes_ClientID cc"
            dt = cCommon.GetDT(Sql, cEnviro.DBHost, "UserManagement", True)

            ' ___ Write the datagrid to the page
            litDG.Text = DG.GetText(dt, Nothing)

            '' ___ Set the last field sorted and sort direction in the viewstate strings
            'Viewstate("UserID") = UserID
            'Viewstate("SortData") = DG.GetViewStateString()

        Catch ex As Exception
            Throw New Exception("Error #907: Permissions DisplayPage. " & ex.Message)
        End Try
    End Sub

End Class