Imports System.Data.SqlClient

Public Class ClientMaintain
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected PageCaption As HtmlGenericControl
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents litHeading As System.Web.UI.WebControls.Literal
    Protected WithEvents litResponseAction As System.Web.UI.WebControls.Literal
    Protected WithEvents litUpdate As System.Web.UI.WebControls.Literal
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
    Protected WithEvents txtClientID As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtName As System.Web.UI.WebControls.TextBox
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
        Dim Results As Results
        Dim RequestAction As RequestAction
        Dim ResponseAction As ResponseAction

        Try

            ' ___ Instantiate objects
            cEnviro = Session("Enviro")
            cCommon = New Common

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right Check
            Dim RightsRqd(0) As String
            cRights = New RightsClass(Page)
            RightsRqd.SetValue(cRights.ClientView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the session object
            cSess = Session("ClientWorklistSession")

            ' ___ Get RequestAction
            RequestAction = cCommon.GetRequestAction(Page)

            ' ___ Execute the RequestAction
            Results = ExecuteRequestAction(RequestAction)
            ResponseAction = Results.ResponseAction

            ' ___ Execute the ResponseAction
            If ResponseAction = ResponseAction.ReturnToCallingPage Then
                cSess.PageReturnOnLoadMessage = Results.Msg
                Response.Redirect("ClientWorklist.aspx?CalledBy=Child")

            Else

                ' ___ Execute the ResponseAction
                DisplayPage(ResponseAction)

                ' ___ Build message if present
                If Not Results.Msg = Nothing Then
                    If Results.ObtainConfirm Then
                        litMsg.Text = "<script language='javascript'>" & Results.Msg & "</script>"
                    Else
                        litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                    End If
                End If

                litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"
            End If

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cenviro.DBHost & "'>"

        Catch ex As Exception
            Throw New Exception("Error #1350: ClientMaintain Page_Load. " & ex.Message)
        End Try
    End Sub

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
                            MyResults.ResponseAction = ResponseAction.ReturnToCallingPage
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
                            MyResults.ResponseAction = ResponseAction.ReturnToCallingPage
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
            Throw New Exception("Error #1351: ClientMaintain ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function IsDataValid(ByVal RequestAction As RequestAction) As Object
        Dim Results As New Results
        Dim ErrColl As New Collection
        Dim dt As DataTable

        Try

            ' ___ Trim the textbox input
            txtName.Text = Trim(txtName.Text)

            cCommon.ValidateStringField(ErrColl, txtClientID.Text, 1, "user id not provided")
            cCommon.ValidateStringField(ErrColl, txtName.Text, 1, "name not provided")
            cCommon.ValidateApostrophe(ErrColl, txtClientID.Text, "userid apostrophe not permitted")
            cCommon.ValidateApostrophe(ErrColl, txtName.Text, "name apostrophe not permitted")

            ' ___ For a new client is the client id already in use?
            If RequestAction = RequestAction.SaveNew Then
                dt = cCommon.GetDT("Select Count (*) FROM Codes_ClientID WHERE lower(ClientID) = " & cCommon.StrOutHandler(txtClientID.Text.ToLower, False, StringTreatEnum.SideQts_SecApost))
                If dt.Rows(0)(0) > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "client id is already  in use")
                End If
            End If

            ' ___ For an existing record are we changing to an existing client id?
            If RequestAction = RequestAction.SaveExisting AndAlso txtClientID.Text.ToLower <> cSess.ClientID.ToLower Then
                dt = cCommon.GetDT("Select Count (*) FROM Codes_ClientID WHERE lower(ClientID) =  " & cCommon.StrOutHandler(txtClientID.Text.ToLower, False, StringTreatEnum.SideQts_SecApost))
                If dt.Rows(0)(0) > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "client id is already  in use")
                End If
            End If

            If ErrColl.Count = 0 Then
                Results.Success = True
            Else
                Results.Success = False
                Results.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
            End If

            ' ___ For an existing record, if we change the client id, are there user records that would be orphaned?
            If Results.Success Then
                If RequestAction = RequestAction.SaveExisting AndAlso txtClientID.Text <> cSess.ClientID Then
                    Dim dtUserCount As DataTable
                    dtUserCount = cCommon.GetDT("SELECT Count (*) FROM UserPermissions WHERE ClientID = " & cCommon.StrOutHandler(cSess.ClientID, False, StringTreatEnum.SideQts_SecApost))
                    If dtUserCount.Rows(0)(0) > 0 Then
                        Results.Success = False
                        'Results.ObtainConfirm = True
                        'Results.Msg = "GetUserConfirmation();"
                        Results.Msg = "There are users currently associated with this client. ClientID cannot be changed."
                    End If
                End If
            End If

            Return Results

        Catch ex As Exception
            Throw New Exception("Error #1352: ClientMaintain IsDataValid. " & ex.Message)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestAction) As Results
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results

        Try

            If RequestAction = RequestAction.SaveNew Then
                Sql.Append("INSERT INTO Codes_ClientID (ClientID, Name, Project, LogicalDelete)")
                Sql.Append(" Values ")
                Sql.Append("(" & cCommon.StrOutHandler(txtClientID.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("'', 0)")
            Else
                Sql.Append("UPDATE Codes_ClientID Set ")
                Sql.Append("ClientID = " & cCommon.StrOutHandler(txtClientID.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("Name = " & cCommon.StrOutHandler(txtName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("LogicalDelete = 0")
                Sql.Append(" WHERE ClientID = " & cCommon.StrOutHandler(cSess.ClientID, False, StringTreatEnum.SideQts_SecApost))
            End If

            cCommon.ExecuteNonQuery(Sql.ToString)

            If Request.Form("hdConfirm") = "yes" Then
                Sql.Length = 0
                Sql.Append("UPDATE Codes_ClientID SET ClientID = " & cCommon.StrOutHandler(txtClientID.Text, False, StringTreatEnum.SideQts_SecApost & " WHERE ClientID = " & cCommon.StrOutHandler(cSess.ClientID, False, StringTreatEnum.SideQts_SecApost)))
                cCommon.ExecuteNonQuery(Sql.ToString)
            End If

            cSess.ClientID = txtClientID.Text

            MyResults.Success = True
            MyResults.Msg = "Update complete."

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Msg = "Unable to save record."
        End Try
        Return MyResults
    End Function

    Private Sub DisplayPage(ByVal ResponseAction As ResponseAction)
        Dim i As Integer
        Dim dt As DataTable

        Try

            ' ___ Set the attributes for controls
            txtClientID.Attributes.Add("onkeyup", "return legallength(this, 50)")
            txtName.Attributes.Add("onkeyup", "return legallength(this, 100)")

            ' ___ Heading/ClientID
            Select Case ResponseAction
                Case ResponseAction.DisplayBlank, ResponseAction.DisplayUserInputNew
                    litHeading.Text = "New Client"
                Case ResponseAction.DisplayExisting, ResponseAction.DisplayUserInputExisting
                    If cRights.HasThisRight(cRights.ClientEdit) Then
                        litHeading.Text = "Edit Client"
                    Else
                        litHeading.Text = "View Client"
                    End If
            End Select

            FormatControls()

            If ResponseAction = ResponseAction.DisplayExisting Then

                ' ___ Get the data
                dt = cCommon.GetDT("Select * From Codes_ClientID Where ClientID = " & cCommon.StrOutHandler(cSess.ClientID, False, StringTreatEnum.SideQts_SecApost))
                txtClientID.Text = cCommon.StrInHandler(dt.Rows(0)("ClientID"))
                txtName.Text = cCommon.StrInHandler(dt.Rows(0)("Name"))
            End If

        Catch ex As Exception
            Throw New Exception("Error #1354: ClientMaintain DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Sub FormatControls()
        Try

            If cRights.HasThisRight(cRights.ClientEdit) Then
                litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"
                Style.AddStyle(txtClientID, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtName, Style.StyleType.NormalEditable, 300, True)
            Else
                Style.AddStyle(txtClientID, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtName, Style.StyleType.NoneditableGrayed, 300, True)
            End If

        Catch ex As Exception
            Throw New Exception("Error #1355: ClientMaintain FormatControls. " & ex.Message)
        End Try
    End Sub


End Class
