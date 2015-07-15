Imports System.Data.SqlClient

Public Class CompanyMaintain
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected PageCaption As HtmlGenericControl
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents litHeading As System.Web.UI.WebControls.Literal
    Protected WithEvents litResponseAction As System.Web.UI.WebControls.Literal
    Protected WithEvents litUpdate As System.Web.UI.WebControls.Literal
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Protected WithEvents txtCompanyID As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtPrimaryContactName As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtPrimaryContactPhone As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtPrimaryContactEmail As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtContactNotes As System.Web.UI.WebControls.TextBox
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
            RightsRqd.SetValue(cRights.CompanyView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the session object
            cSess = Session("CompanyWorklistSession")

            ' ___ Get RequestAction
            RequestAction = cCommon.GetRequestAction(Page)

            ' ___ Execute the RequestAction
            Results = ExecuteRequestAction(RequestAction)
            ResponseAction = Results.ResponseAction

            ' ___ Display page or return to calling page
            If ResponseAction = ResponseAction.ReturnToCallingPage Then
                cSess.PageReturnOnLoadMessage = Results.Msg
                Response.Redirect("CompanyWorklist.aspx?CalledBy=Child")

            Else

                ' ___ Execute the ResponseAction
                DisplayPage(ResponseAction)

                ' ___ Build message if present
                If Not Results.Msg = Nothing Then
                    litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                End If

                litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"
            End If

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'>"

        Catch ex As Exception
            Throw New Exception("Error #350: CompanyMaintain Page_Load. " & ex.Message)
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
            Throw New Exception("Error #351: CompanyMaintain ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function IsDataValid(ByVal RequestAction As RequestAction) As Results
        Dim Results As New Results
        Dim ErrColl As New Collection
        Dim dt As DataTable

        Try

            ' ___ Trim the textbox input
            txtCompanyID.Text = Trim(txtCompanyID.Text)
            txtDescription.Text = Trim(txtDescription.Text)
            txtPrimaryContactName.Text = Trim(txtPrimaryContactName.Text)
            txtPrimaryContactEmail.Text = Trim(txtPrimaryContactEmail.Text)
            txtContactNotes.Text = Trim(txtContactNotes.Text)

            cCommon.ValidateStringField(ErrColl, txtCompanyID.Text, 1, "company id not provided")
            If txtPrimaryContactPhone.Text.Length > 0 Then
                cCommon.ValidatePhoneNumber(ErrColl, txtPrimaryContactPhone.Text, "phone number must be blank if valid phone number with area code not provided")
            End If

            ' ___ For a new record is the company id already in use?
            If RequestAction = RequestAction.SaveNew Then
                ' dt = cCommon.GetDT("Select Count (*) FROM Codes_CompanyID WHERE lower(CompanyID) = '" & txtCompanyID.Text.ToLower & "'")
                dt = cCommon.GetDT("Select Count (*) FROM Codes_CompanyID WHERE lower(CompanyID) = " & cCommon.StrOutHandler(txtCompanyID.Text.ToLower, False, StringTreatEnum.SideQts_SecApost))
                If dt.Rows(0)(0) > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "company id is already  in use")
                End If
            End If

            ' ___ For an existing record are we changing to an existing company id?
            If RequestAction = RequestAction.SaveExisting AndAlso txtCompanyID.Text.ToLower <> cSess.CompanyID.ToLower Then
                dt = cCommon.GetDT("Select Count (*) FROM Codes_CompanyID WHERE lower(CompanyID)  = " & cCommon.StrOutHandler(txtCompanyID.Text.ToLower, False, StringTreatEnum.SideQts_SecApost))
                If dt.Rows(0)(0) > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "company id is already  in use")
                End If
            End If

            ' ___ Email, if included, must be valid
            If txtPrimaryContactEmail.Text.Length > 0 Then
                cCommon.ValidateEmailAddress(ErrColl, txtPrimaryContactEmail.Text, "valid email address not provided")
            End If

            If ErrColl.Count = 0 Then
                Results.Success = True
            Else
                Results.Success = False
                Results.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
            End If

            ' ___ For an existing record, if we change the company id, are there user records that would be orphaned?
            If Results.Success Then
                If RequestAction = RequestAction.SaveExisting AndAlso txtCompanyID.Text <> cSess.CompanyID Then
                    Dim dtUserCount As DataTable
                    ' dtUserCount = cCommon.GetDT("SELECT Count (*) FROM Users WHERE CompanyID = '" & cSess.CompanyID & "'")
                    dtUserCount = cCommon.GetDT("SELECT Count (*) FROM Users WHERE CompanyID = " & cCommon.StrOutHandler(cSess.CompanyID, False, StringTreatEnum.SideQts_SecApost))
                    If dtUserCount.Rows(0)(0) > 0 Then
                        Results.Success = False
                        Results.ObtainConfirm = True
                        Results.Msg = "ConfirmKeyChange('There are users currently associated with this Company ID. If you proceed with this change in Company ID, these users will be updated to the new Company ID value. Do you wish to proceed with this change?');"
                    End If
                End If
            End If

            Return Results

        Catch ex As Exception
            Throw New Exception("Error #352: CompanyMaintain IsDataValid. " & ex.Message)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestAction) As Results
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results


        Try

            If RequestAction = RequestAction.SaveNew Then
                Sql.Append("INSERT INTO Codes_CompanyID (CompanyID, Description, PrimaryContactName, PrimaryContactPhone, PrimaryContactEmail, ContactNotes, AddDate, ChangeDate)")
                Sql.Append(" Values ")
                Sql.Append("(" & cCommon.StrOutHandler(txtCompanyID.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtDescription.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtPrimaryContactName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.PhoneOutHandler(txtPrimaryContactPhone.Text, False, True) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtPrimaryContactEmail.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtContactNotes.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("'" & cCommon.GetServerDateTime & "', ")
                Sql.Append("'" & cCommon.GetServerDateTime & "') ")

            Else
                Sql.Append("UPDATE Codes_CompanyID Set ")
                Sql.Append("CompanyID = " & cCommon.StrOutHandler(txtCompanyID.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("Description = " & cCommon.StrOutHandler(txtDescription.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("PrimaryContactName = " & cCommon.StrOutHandler(txtPrimaryContactName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("PrimaryContactPhone = " & cCommon.PhoneOutHandler(txtPrimaryContactPhone.Text, False, True) & ", ")
                Sql.Append("PrimaryContactEmail = " & cCommon.StrOutHandler(txtPrimaryContactEmail.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("ContactNotes = " & cCommon.StrOutHandler(txtContactNotes.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("ChangeDate = '" & cCommon.GetServerDateTime & "' ")
                ' Sql.Append(" WHERE CompanyID = '" & cSess.CompanyID & "'")
                Sql.Append(" WHERE CompanyID = " & cCommon.StrOutHandler(cSess.CompanyID, False, StringTreatEnum.SideQts_SecApost))
            End If

            cCommon.ExecuteNonQuery(Sql.ToString)

            If Request.Form("hdConfirm") = "yes" Then
                Sql.Length = 0
                ' Sql.Append("UPDATE Users SET CompanyID = '" & txtCompanyID.Text & "' WHERE CompanyID = '" & cSess.CompanyID & "'")
                Sql.Append("UPDATE Users SET CompanyID = " & cCommon.StrOutHandler(txtCompanyID.Text, False, StringTreatEnum.SideQts_SecApost) & " WHERE CompanyID = " & cCommon.StrOutHandler(cSess.CompanyID, False, StringTreatEnum.SideQts_SecApost))
                cCommon.ExecuteNonQuery(Sql.ToString)
            End If

            cSess.CompanyID = txtCompanyID.Text

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
            txtContactNotes.Attributes.Add("onkeyup", "return legallength(this, 512)")

            ' ___ Heading/CompanyID
            Select Case ResponseAction
                Case ResponseAction.DisplayBlank, ResponseAction.DisplayUserInputNew
                    litHeading.Text = "New Company"
                Case ResponseAction.DisplayExisting, ResponseAction.DisplayUserInputExisting
                    If cRights.HasThisRight(cRights.CompanyEdit) Then
                        litHeading.Text = "Edit Company Contact"
                    Else
                        litHeading.Text = "View Company Contact"
                    End If
            End Select

            FormatControls()

            If ResponseAction = ResponseAction.DisplayExisting Then

                ' ___ Get the data
                'dt = cCommon.GetDT("SELECT * FROM Codes_CompanyID WHERE CompanyID = '" & cSess.CompanyID & "'")
                dt = cCommon.GetDT("SELECT * FROM Codes_CompanyID WHERE CompanyID = " & cCommon.StrOutHandler(cSess.CompanyID, False, StringTreatEnum.SideQts_SecApost))
                txtCompanyID.Text = cCommon.StrInHandler(dt.Rows(0)("CompanyID"))
                txtDescription.Text = cCommon.StrInHandler(dt.Rows(0)("Description"))
                txtPrimaryContactName.Text = cCommon.StrInHandler(dt.Rows(0)("PrimaryContactName"))
                txtPrimaryContactPhone.Text = cCommon.StrInHandler(dt.Rows(0)("PrimaryContactPhone"))
                txtPrimaryContactEmail.Text = cCommon.StrInHandler(dt.Rows(0)("PrimaryContactEmail"))
                txtContactNotes.Text = cCommon.StrInHandler(dt.Rows(0)("ContactNotes"))
            End If

        Catch ex As Exception
            Throw New Exception("Error #354: CompanyMaintain DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Sub FormatControls()
        Try

            If cRights.HasThisRight(cRights.CompanyEdit) Then
                litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"
                Style.AddStyle(txtCompanyID, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtDescription, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtPrimaryContactName, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtPrimaryContactPhone, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtPrimaryContactEmail, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtContactNotes, Style.StyleType.NormalEditable, 400, True)
            Else
                Style.AddStyle(txtCompanyID, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtDescription, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtPrimaryContactName, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtPrimaryContactPhone, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtPrimaryContactEmail, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtContactNotes, Style.StyleType.NoneditableGrayed, 400, True)
            End If

        Catch ex As Exception
            Throw New Exception("Error #355: CompanyMaintain FormatControls. " & ex.Message)
        End Try
    End Sub
End Class
